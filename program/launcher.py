"""Portable thesis converter UI and command-line entry point."""
from __future__ import annotations
import argparse
from datetime import datetime
import json
import os
from pathlib import Path
import queue
import shutil
import subprocess
import sys
import threading
import time
import traceback

BASE = Path(__file__).resolve().parent.parent
CORE = BASE / 'program'


def environment():
    env = os.environ.copy()
    env.pop('PYTHONHOME', None)
    env.pop('PYTHONPATH', None)
    env['PYTHONUTF8'] = '1'
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PATH'] = os.pathsep.join([str(BASE / 'runtime'), str(BASE / 'tools'), env.get('PATH', '')])
    return env


def check():
    import fitz, docx, lxml, pandoc_tex_numbering
    env = environment()
    missing = [name for name in ('latexmk', 'xelatex', 'biber') if not shutil.which(name, path=env['PATH'])]
    if missing:
        raise RuntimeError('未找到 ' + ', '.join(missing) + '。请安装完整 TeX Live，并将 bin/windows 加入 PATH 后重新打开窗口。')
    word = subprocess.run(['powershell', '-NoProfile', '-Command',
        "if ($null -eq [type]::GetTypeFromProgID('Word.Application')) {exit 1}"],
        capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
    if word.returncode:
        raise RuntimeError('未找到桌面版 Microsoft Word；WPS 或网页版 Word 不能替代目录刷新步骤。')
    if not (BASE / 'tools' / 'pandoc.exe').exists():
        raise RuntimeError('内置 Pandoc 缺失，请重新解压完整发布包。')
    return 'Python、转换库、Pandoc、TeX Live 命令和 Word 注册检查通过。字体和 TeX 宏包由实际编译进一步验证。'


def convert(source: Path, emit=print):
    # Named OS mutex prevents two windows from calibrating one reference concurrently.
    import ctypes
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.CreateMutexW.argtypes = (ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p)
    kernel.CreateMutexW.restype = ctypes.c_void_p
    kernel.CloseHandle.argtypes = (ctypes.c_void_p,)
    handle = kernel.CreateMutexW(None, False, 'Local\\PortableThesisConverter')
    if not handle:
        raise ctypes.WinError(ctypes.get_last_error())
    if ctypes.get_last_error() == 183:
        kernel.CloseHandle(handle)
        raise RuntimeError('已有转换任务运行，请等待完成后重试。')
    try:
        return convert_job(source, emit)
    finally:
        kernel.CloseHandle(handle)


def convert_job(source: Path, emit=print):
    started = time.monotonic()
    source = source.resolve()
    if source.parent != BASE or source.suffix.lower() != '.tex':
        raise ValueError('请将 TeX 文档及对应资源放在发布包根目录，再选择该文件。')
    emit(check())
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    label = source.stem + '_' + stamp
    work = BASE / '.work' / label
    work.mkdir(parents=True)
    for path in BASE.iterdir():
        if path.is_file() and path.suffix.lower() in ('.tex', '.bib', '.sty', '.cls', '.bst'):
            shutil.copy2(path, work / path.name)
    if (BASE / 'figures').exists():
        shutil.copytree(BASE / 'figures', work / 'figures')
    output = BASE / 'output'
    output.mkdir(exist_ok=True)
    log = output / (label + '.log')
    command = [str(BASE / 'runtime' / 'python.exe'), '-u', str(CORE / 'build_thesis.py'),
               '--tex', str(work / source.name), '--word-output', str(work / (label + '.docx'))]
    emit('开始转换：' + source.name)
    with log.open('w', encoding='utf-8') as stream:
        process = subprocess.Popen(command, cwd=BASE, env=environment(), stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace',
            creationflags=subprocess.CREATE_NO_WINDOW)
        for line in process.stdout:
            stream.write(line)
            stream.flush()
            emit(line.rstrip())
        result = process.wait()
    if result:
        raise RuntimeError(f'转换失败，未发布本次结果。日志：{log}\n检查编译错误、字体、宏包，以及 Word 是否占用文件。')
    paths = []
    for origin, suffix in [(work / source.with_suffix('.pdf').name, '.pdf'), (work / (label + '.docx'), '.docx')]:
        destination = output / (label + suffix)
        shutil.copy2(origin, destination)
        paths.append(destination)
    emit('转换完成，耗时 %.1f 秒：\n' % (time.monotonic() - started) + '\n'.join(map(str, paths)))
    # Only this uniquely created job directory is disposable. Keep failed jobs for diagnostics.
    shutil.rmtree(work)
    return paths


def gui(smoke_test=False):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter.scrolledtext import ScrolledText
    root = tk.Tk()
    root.title('LaTeX 论文转 PDF / Word')
    root.geometry('780x550')
    selected = tk.StringVar(value=str(BASE / 'thesis-template.tex'))
    busy = [False]
    events = queue.Queue()
    frame = ttk.Frame(root, padding=14)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='选择本文件夹中的 TeX；自动生成 PDF 和 Word，文件名附时间戳。').pack(anchor='w')
    row = ttk.Frame(frame)
    row.pack(fill='x', pady=10)
    entry = ttk.Entry(row, textvariable=selected)
    entry.pack(side='left', fill='x', expand=True)
    def choose():
        path = filedialog.askopenfilename(initialdir=BASE, filetypes=[('LaTeX', '*.tex')])
        if path: selected.set(path)
    browse = ttk.Button(row, text='选择文件', command=choose)
    browse.pack(side='left', padx=6)
    controls = ttk.Frame(frame)
    controls.pack(fill='x')
    status = tk.StringVar(value='就绪。首次编译或长论文可能需要数分钟。')
    text = ScrolledText(frame, height=22, font=('Consolas', 10))
    text.pack(fill='both', expand=True, pady=10)
    ttk.Label(frame, textvariable=status).pack(anchor='w')
    def worker(source):
        try:
            paths = convert(source, lambda line: events.put(('log', line)))
            events.put(('done', paths))
        except Exception:
            events.put(('error', traceback.format_exc()))
    def start():
        busy[0] = True
        start_button.config(state='disabled')
        browse.config(state='disabled')
        entry.config(state='disabled')
        status.set('转换中，请等待；不要打开正在生成的 Word。')
        threading.Thread(target=worker, args=(Path(selected.get()),), daemon=True).start()
    def preflight():
        try: messagebox.showinfo('环境检查', check())
        except Exception as error: messagebox.showerror('环境检查', str(error))
    start_button = ttk.Button(controls, text='开始转换 PDF + Word', command=start)
    start_button.pack(side='left')
    ttk.Button(controls, text='检查环境', command=preflight).pack(side='left', padx=8)
    ttk.Button(controls, text='打开结果文件夹', command=lambda: os.startfile(BASE / 'output')).pack(side='left')
    def poll():
        while not events.empty():
            kind, value = events.get()
            if kind == 'log':
                text.insert('end', value + '\n')
                text.see('end')
            else:
                busy[0] = False
                start_button.config(state='normal')
                browse.config(state='normal')
                entry.config(state='normal')
                status.set('转换成功，结果已保存到 output。' if kind == 'done' else '转换失败，查看日志。')
                if kind == 'error':
                    text.insert('end', value + '\n')
                    messagebox.showerror('转换失败', value.splitlines()[-1])
                else:
                    messagebox.showinfo('转换完成', '\n'.join(map(str, value)))
        root.after(150, poll)
    def close():
        if busy[0]: messagebox.showinfo('转换中', '请等待本次转换完成后关闭窗口。')
        else: root.destroy()
    root.protocol('WM_DELETE_WINDOW', close)
    root.after(150, poll)
    if smoke_test:
        root.withdraw()
        root.after(400, root.destroy)
    root.mainloop()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--check', action='store_true')
    parser.add_argument('--convert', type=Path)
    parser.add_argument('--gui-smoke-test', action='store_true')
    args = parser.parse_args()
    if args.check: print(check())
    elif args.convert: convert(args.convert)
    else: gui(args.gui_smoke_test)
