"""Bounded, read-only MathType discovery. Finding a file is not compatibility."""
from dataclasses import dataclass, asdict
import json
import os
from pathlib import Path
import re

DEFAULT_CONFIG = Path(__file__).resolve().parents[1] / 'config' / 'environment.json'


@dataclass
class Candidate:
    path: str
    sources: list[str]


@dataclass
class Discovery:
    status: str
    candidates: list[Candidate]
    searched: list[str]
    errors: list[str]
    installations: list[str]

    def to_dict(self):
        return asdict(self)

    def require(self):
        if self.status == 'found':
            return Path(self.candidates[0].path)
        messages = {
            'selection_required': '找到多个 MathType 接口，请明确选择安装目录或接口文件。',
            'interface_missing': '发现 MathType 安装位置，但没有找到所需的 MathPage/64/MathPage.wll。',
            'not_found': '未找到 MathType 安装位置，请手动指定；这不等于软件一定未安装。',
            'invalid_override': '手动路径、环境变量或已保存路径无效，请修正或清除后重新查找。',
            'config_error': '无法读取路径设置文件，请检查诊断信息。',
        }
        raise RuntimeError(messages[self.status] + ('\n' + '\n'.join(self.errors) if self.errors else ''))


def _location(value):
    text = os.path.expandvars(str(value).strip())
    if text.startswith('"'):
        match = re.match(r'^"([^"]+)"', text)
        if match:
            text = match.group(1)
    text = re.sub(r',\s*-?\d+\s*$', '', text).strip('"')
    return Path(text)


def registry_locations():
    """Read both registry views; never enumerate unrelated directories."""
    try:
        import winreg
    except ImportError:
        return [], []
    found, errors = [], []
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for view in (winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY):
            try:
                with winreg.OpenKey(hive, r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MathType.exe', 0, winreg.KEY_READ | view) as key:
                    for field in ('', 'Path'):
                        try:
                            value = winreg.QueryValueEx(key, field)[0]
                            if value:
                                found.append((_location(value), 'registry:App Paths'))
                        except FileNotFoundError:
                            pass
                        except OSError as error:
                            errors.append('App Paths value: ' + str(error))
            except FileNotFoundError:
                pass
            except OSError as error:
                errors.append('App Paths: ' + str(error))
            try:
                with winreg.OpenKey(hive, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 0, winreg.KEY_READ | view) as uninstall:
                    for index in range(winreg.QueryInfoKey(uninstall)[0]):
                        try:
                            with winreg.OpenKey(uninstall, winreg.EnumKey(uninstall, index)) as key:
                                name = str(winreg.QueryValueEx(key, 'DisplayName')[0])
                                if 'mathtype' not in name.lower():
                                    continue
                                for field in ('InstallLocation', 'DisplayIcon'):
                                    try:
                                        value = winreg.QueryValueEx(key, field)[0]
                                        if value:
                                            found.append((_location(value), 'registry:Uninstall'))
                                    except FileNotFoundError:
                                        pass
                        except FileNotFoundError:
                            pass
                        except OSError as error:
                            errors.append('Uninstall entry: ' + str(error))
            except FileNotFoundError:
                pass
            except OSError as error:
                errors.append('Uninstall: ' + str(error))
    return found, errors


def discover(manual=None, config=DEFAULT_CONFIG, env=None, registry=None):
    env = os.environ if env is None else env
    errors, searched, installations, candidates = [], [], [], {}
    override = manual or env.get('MATHTYPE_DLL')
    source = 'manual' if manual else 'environment:MATHTYPE_DLL'
    if override and not str(override).strip():
        return Discovery('invalid_override', [], [], ['Path is blank'], [])
    if not override and config and Path(config).exists():
        try:
            settings = json.loads(Path(config).read_text(encoding='utf-8'))
            if settings.get('schema_version') != 1:
                raise ValueError('Unsupported settings schema')
            override = settings.get('mathtype_location')
            if override is not None and not isinstance(override, str):
                raise ValueError('mathtype_location must be a string')
            source = 'saved settings'
        except (OSError, ValueError, AttributeError) as error:
            return Discovery('config_error', [], [], [str(error)], [])
    if override:
        locations = [(_location(override), source)]
    else:
        locations, errors = registry_locations() if registry is None else (list(registry), [])
        for key in ('ProgramFiles(x86)', 'ProgramFiles', 'ProgramW6432', 'LOCALAPPDATA'):
            if env.get(key):
                root = Path(env[key])
                for relative in ('MathType', 'Wiris/MathType', 'Programs/MathType'):
                    locations.append((root / relative, 'common:' + key))
    for location, origin in locations:
        location = Path(location)
        if location.suffix.lower() == '.exe':
            location = location.parent
        if location.name.lower() == 'mathpage.wll':
            paths = [location]
        elif location.is_dir():
            installations.append(str(location.resolve()))
            paths = [location / 'MathPage/64/MathPage.wll', location / '64/MathPage.wll', location / 'MathPage.wll']
            # A finite one-level check supports versioned folders under MathType.
            try:
                if location.name.lower() == 'mathtype':
                    paths.extend(child / 'MathPage/64/MathPage.wll' for child in location.iterdir() if child.is_dir())
            except OSError as error:
                errors.append(str(error))
        else:
            paths = [location / 'MathPage/64/MathPage.wll']
        for path in paths:
            searched.append(str(path))
            if path.is_file() and path.name.lower() == 'mathpage.wll':
                key = os.path.normcase(str(path.resolve()))
                entry = candidates.setdefault(key, Candidate(str(path.resolve()), []))
                if origin not in entry.sources:
                    entry.sources.append(origin)
    values = list(candidates.values())
    if override and len(values) != 1:
        status = 'selection_required' if values else 'invalid_override'
    elif len(values) > 1:
        status = 'selection_required'
    elif values:
        status = 'found'
    else:
        status = 'interface_missing' if installations else 'not_found'
    return Discovery(status, values, sorted(set(searched)), errors, sorted(set(installations)))


def save_location(location, config=DEFAULT_CONFIG):
    result = discover(manual=str(location), config=None)
    result.require()
    path = Path(config)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = json.loads(path.read_text(encoding='utf-8')) if path.exists() else {'schema_version': 1}
    if not isinstance(data, dict) or data.get('schema_version') != 1:
        raise ValueError('Unsupported settings schema')
    data['mathtype_location'] = str(Path(location).resolve())
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(temporary, path)
    return result


def configure_conversion(env=None, config=DEFAULT_CONFIG, manual=None):
    """GUI checks and build/MathType adapters use this one resolver."""
    result = discover(manual=manual, config=config, env=env)
    selected = result.require()
    configured = dict(os.environ if env is None else env)
    configured['MATHTYPE_DLL'] = str(selected)
    return configured, result
