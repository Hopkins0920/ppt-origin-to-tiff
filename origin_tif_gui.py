"""Small Windows front end and packaged-worker entry point."""
import os
from pathlib import Path
import sys

if len(sys.argv) > 1 and sys.argv[1] == '--engine':
    import origin_pptx_to_tif as engine
    try:
        sys.exit(engine.main(sys.argv[2:]))
    except BaseException as exc:
        if isinstance(exc, SystemExit): raise
        import traceback
        # Worker logs still exist when a windowed executable has no standard streams.
        if len(sys.argv) > 4 and sys.argv[2] == '--_worker':
            try:
                import json
                job = json.loads(Path(sys.argv[4]).read_text(encoding='utf-8'))
                Path(job['progress']).with_suffix('.error.txt').write_text(traceback.format_exc(), encoding='utf-8')
            except Exception: pass
        if sys.stderr is not None: traceback.print_exc()
        sys.exit(1)

import json
import queue
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import uuid


class OriginTifApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Origin 图表转 TIF · PPT 工具 v1.1')
        self.geometry('1080x780'); self.minsize(860, 650)
        self.configure(background='#f4f6f8')
        style = ttk.Style(self)
        style.theme_use('vista' if 'vista' in style.theme_names() else 'clam')
        style.configure('.', font=('Microsoft YaHei UI', 10))
        style.configure('Treeview', rowheight=28)
        style.configure('Title.TLabel', font=('Microsoft YaHei UI', 20, 'bold'))
        style.configure('Sub.TLabel', foreground='#5e6876')
        self.source = tk.StringVar(); self.destination = tk.StringVar(); self.dpi = tk.StringVar(value='600')
        self.render = tk.BooleanVar(value=True)
        self.skip_failed = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value='选择 PPTX 文件，然后点击“识别图表”。')
        self.summary = tk.StringVar(value='尚未选择文件')
        self.percent = tk.StringVar(value='0%')
        self.messages = queue.Queue(); self.process = None; self.busy = False
        self.rows = {}; self.report_path = None; self.work_dir = None; self.result_path = None
        self.mode = None; self.failed = False; self.cancelled = False; self.close_when_done = False; self.current_object = None
        self.no_success = False
        self.temp = Path(tempfile.gettempdir())/'OriginTifTool'; self.temp.mkdir(exist_ok=True)
        self.buttons = []
        self._build()
        self.protocol('WM_DELETE_WINDOW', self.on_close)
        self.after(100, self.poll)

    def _build(self):
        root = ttk.Frame(self, padding=24); root.pack(fill='both', expand=True)
        ttk.Label(root, text='Origin 图表转 TIF', style='Title.TLabel').pack(anchor='w')
        ttk.Label(root, text='使用本机 Origin 导出，保留 PPT 中图表的尺寸、位置与组合关系。', style='Sub.TLabel').pack(anchor='w', pady=(4, 20))
        form = ttk.Frame(root); form.pack(fill='x')
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text='PPT 文件').grid(row=0, column=0, sticky='w', padx=(0, 14))
        self.source_entry = ttk.Entry(form, textvariable=self.source)
        self.source_entry.grid(row=0, column=1, sticky='ew', pady=5)
        self.button(form, '选择文件…', self.choose_source).grid(row=0, column=2, padx=(10, 0))
        ttk.Label(form, text='另存为').grid(row=1, column=0, sticky='w')
        self.dest_entry = ttk.Entry(form, textvariable=self.destination)
        self.dest_entry.grid(row=1, column=1, sticky='ew', pady=5)
        self.button(form, '保存位置…', self.choose_output).grid(row=1, column=2, padx=(10, 0))
        options = ttk.Frame(root); options.pack(fill='x', pady=(12, 16))
        ttk.Label(options, text='TIF 分辨率').pack(side='left')
        self.dpi_combo = ttk.Combobox(options, textvariable=self.dpi, values=('150','300','600','1200'), state='readonly', width=8)
        self.dpi_combo.pack(side='left', padx=(10, 5)); ttk.Label(options, text='DPI').pack(side='left')
        self.render_check = ttk.Checkbutton(options, text='使用 PowerPoint 生成替换前后预览', variable=self.render)
        self.render_check.pack(side='left', padx=24)
        self.skip_check = ttk.Checkbutton(root, text='跳过失败图表，保留原 Origin 对象并继续', variable=self.skip_failed)
        self.skip_check.pack(anchor='w', pady=(0, 12))
        actions = ttk.Frame(root); actions.pack(fill='x', pady=(0, 12))
        self.scan_button = self.button(actions, '识别图表', self.scan); self.scan_button.pack(side='left')
        self.convert_button = self.button(actions, '开始转换', self.convert); self.convert_button.pack(side='left', padx=10)
        self.cancel_button = ttk.Button(actions, text='取消', command=self.cancel, state='disabled'); self.cancel_button.pack(side='left')
        ttk.Label(actions, textvariable=self.summary, style='Sub.TLabel').pack(side='right')
        table = ttk.Frame(root)
        columns = ('slide','object','type','size','group','state')
        self.tree = ttk.Treeview(table, columns=columns, show='headings', selectmode='browse', height=8)
        widths = (65,110,160,160,85,220)
        for key, label, width in zip(columns, ('页码','对象 ID','Origin 类型','显示尺寸 / cm','组合层数','状态'), widths):
            self.tree.heading(key, text=label); self.tree.column(key, width=width, minwidth=55, stretch=(key=='state'))
        scroll = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side='left', fill='both', expand=True); scroll.pack(side='right', fill='y')
        self.tree.tag_configure('done', foreground='#157344')
        self.tree.tag_configure('review', foreground='#9b6100')
        self.tree.tag_configure('error', foreground='#ac2430')
        self.tree.bind('<Double-1>', self.open_comparison)
        bottom = ttk.Frame(root); bottom.pack(side='bottom', fill='x')
        ttk.Label(bottom, textvariable=self.status, wraplength=970).pack(anchor='w', pady=(14, 8))
        progress = ttk.Frame(bottom); progress.pack(fill='x')
        self.bar = ttk.Progressbar(progress, mode='determinate', maximum=100)
        self.bar.pack(side='left', fill='x', expand=True)
        ttk.Label(progress, textvariable=self.percent, width=6, anchor='e').pack(side='right')
        self.log = tk.Text(bottom, height=4, font=('Microsoft YaHei UI',9), background='#fafbfc', relief='flat', wrap='word', state='disabled')
        self.log.pack(fill='x', pady=(12, 10))
        footer = ttk.Frame(bottom); footer.pack(fill='x')
        self.open_result = ttk.Button(footer, text='打开转换后的 PPT', command=lambda:self.open_path(self.result_path), state='disabled')
        self.open_result.pack(side='left')
        self.open_folder = ttk.Button(footer, text='图表与前后预览', command=lambda:self.open_path(self.work_dir), state='disabled')
        self.open_folder.pack(side='left', padx=10)
        self.open_report = ttk.Button(footer, text='检查报告', command=self.show_report, state='disabled')
        self.open_report.pack(side='left')
        ttk.Label(footer, text='原文件保留 · 跳过项保留原图', style='Sub.TLabel').pack(side='right')
        table.pack(fill='both', expand=True)

    def button(self, parent, text, command):
        widget = ttk.Button(parent, text=text, command=command); self.buttons.append(widget); return widget

    def choose_source(self):
        path = filedialog.askopenfilename(title='选择包含 Origin 嵌入图表的 PPT', filetypes=[('PowerPoint 演示文稿','*.pptx')])
        if path:
            self.source.set(path); self.destination.set(str(self.proposed_output(Path(path))))
            self.scan()

    def proposed_output(self, source):
        path = source.with_name(source.stem + f'_Origin_TIF_{self.dpi.get()}dpi.pptx')
        i = 2
        while path.exists():
            path = source.with_name(source.stem + f'_Origin_TIF_{self.dpi.get()}dpi_{i}.pptx'); i += 1
        return path

    def choose_output(self):
        path = filedialog.asksaveasfilename(title='选择新 PPT 保存位置', defaultextension='.pptx',
                                           initialfile=Path(self.destination.get()).name if self.destination.get() else 'Origin_TIF_600dpi.pptx',
                                           filetypes=[('PowerPoint 演示文稿','*.pptx')])
        if path: self.destination.set(path)

    def source_path(self):
        path = Path(self.source.get().strip())
        if not path.is_file() or path.suffix.lower() != '.pptx':
            messagebox.showerror('请选择文件', '请选择现有的 .pptx 文件。旧版 .ppt 需先另存为 .pptx。', parent=self)
            return None
        return path.resolve()

    def scan(self):
        path = self.source_path()
        if path and not self.busy: self.start('scan', path)

    def convert(self):
        source = self.source_path()
        if not source or self.busy: return
        if not self.destination.get(): self.destination.set(str(self.proposed_output(source)))
        output = Path(self.destination.get()).resolve()
        if output == source or output.exists():
            messagebox.showerror('请选择新文件名', '工具不会覆盖原文件或已有结果，请选择一个新的保存文件名。', parent=self); return
        if output.suffix.lower() != '.pptx':
            messagebox.showerror('文件格式', '保存文件名须以 .pptx 结尾。', parent=self); return
        self.start('convert', source)

    def set_busy(self, busy):
        self.busy = busy
        for button in self.buttons: button.configure(state='disabled' if busy else 'normal')
        for entry in (self.source_entry,self.dest_entry): entry.configure(state='disabled' if busy else 'normal')
        self.dpi_combo.configure(state='disabled' if busy else 'readonly')
        self.render_check.configure(state='disabled' if busy else 'normal')
        self.skip_check.configure(state='disabled' if busy else 'normal')
        self.cancel_button.configure(state='normal' if busy else 'disabled')

    def start(self, mode, source):
        self.mode = mode; self.failed = False; self.cancelled = False; self.report_path = None
        self.current_object = None
        self.no_success = False
        self.work_dir = None; self.result_path = None
        self.log.configure(state='normal');self.log.delete('1.0','end');self.log.configure(state='disabled')
        self.log_line(('识别文件：' if mode=='scan' else '转换文件：')+str(source))
        for button in (self.open_result,self.open_report,self.open_folder): button.configure(state='disabled')
        self.tree.delete(*self.tree.get_children()); self.rows.clear()
        self.set_busy(True); self.set_progress(0)
        self.summary.set('正在读取…'); self.status.set('正在识别 Origin 嵌入图表…')
        token = uuid.uuid4().hex
        self.events_path = self.temp/(token+'.jsonl'); self.cancel_path = self.temp/(token+'.cancel')
        args = [str(source), '--dpi', self.dpi.get(), '--event-file', str(self.events_path), '--cancel-file', str(self.cancel_path)]
        if mode == 'scan': args += ['--dry-run']
        else:
            args += ['--output', self.destination.get()]
            if not self.render.get(): args += ['--no-render']
            if self.skip_failed.get(): args += ['--skip-failed']
        if getattr(sys,'frozen',False): command = [sys.executable,'--engine',*args]
        else: command = [sys.executable,str(Path(__file__).with_name('origin_pptx_to_tif.py')),*args]
        try:
            self.process = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                            env=dict(os.environ,PYTHONUTF8='1'), creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        except Exception as exc:
            self.set_busy(False); messagebox.showerror('启动失败',str(exc),parent=self); return
        threading.Thread(target=self.read_events,args=(self.process,self.events_path),daemon=True).start()

    def read_events(self, process, path):
        offset = 0; buffer = ''
        while True:
            if path.exists():
                with path.open('r',encoding='utf-8') as f:
                    f.seek(offset); buffer += f.read(); offset = f.tell()
                while '\n' in buffer:
                    line,buffer = buffer.split('\n',1)
                    try:self.messages.put(json.loads(line))
                    except json.JSONDecodeError:pass
            if process.poll() is not None:
                # Read once more after process completion so final events cannot be lost.
                if path.exists():
                    with path.open('r',encoding='utf-8') as f:
                        f.seek(offset)
                        for line in f:
                            try:self.messages.put(json.loads(line))
                            except json.JSONDecodeError:pass
                self.messages.put({'event':'exit','code':process.returncode}); return
            import time; time.sleep(.15)

    def poll(self):
        try:
            while True:self.handle(self.messages.get_nowait())
        except queue.Empty:pass
        if self.winfo_exists():self.after(100,self.poll)

    def set_progress(self,value):
        self.bar['value']=value;self.percent.set(f'{round(value)}%')

    def log_line(self,message):
        if not message:return
        self.log.configure(state='normal');self.log.insert('end',message+'\n');self.log.see('end');self.log.configure(state='disabled')

    def object_state(self,name,text,tag=''):
        if name in self.rows:
            self.tree.set(name,'state',text);self.tree.item(name,tags=(tag,) if tag else ())

    def mark_skipped(self,row):
        name=row['name']
        if name not in self.rows:return
        first=self.rows[name].get('status')!='skipped'
        self.rows[name].update(row)
        self.object_state(name,'已跳过 · 保留原 Origin','error')
        if first:
            reason=row.get('error','').split('Attempts:')[0].strip()
            self.log_line(f"跳过：第 {row['slide']} 页，对象 {row['id']}。{reason}")

    def handle(self,event):
        kind=event['event']
        if kind=='inventory':
            self.tree.delete(*self.tree.get_children());self.rows={r['name']:r for r in event['rows']}
            for name,r in self.rows.items():
                w,h=r['display_inches']
                self.tree.insert('', 'end', iid=name, values=(r['slide'],r['id'],r['prog_id'],f'{w*2.54:.2f} × {h*2.54:.2f}',r['group_depth'],'已识别'))
            self.summary.set(f"{event['slides']} 页 · {len(self.rows)} 个 Origin 图表")
            self.work_dir=event['work_dir'];self.set_progress(5 if self.mode=='convert' else 100)
        elif kind=='progress':
            stage=event.get('stage');name=event.get('name','');index=event.get('index',0);total=event.get('total',max(1,len(self.rows)))
            for row in event.get('skipped',[]):self.mark_skipped(row)
            labels={'references':'正在读取原图预览…','starting_origin':'正在启动 Origin…','closing_origin':'Origin 导出结束，正在关闭临时实例…',
                    'packaging':'正在替换并核验尺寸、位置与文件结构…','opening_powerpoint':'正在打开 PowerPoint 进行预览验证…'}
            if stage in ('export','exported'):
                self.current_object = name if stage == 'export' else None
                # Progress snapshots can skip a short "exported" phase. Earlier rows are done.
                names=list(self.rows)
                for prior in names[:max(0,index-1)]:
                    if self.rows[prior].get('status')!='skipped':self.object_state(prior,'已导出，待核验','done')
                self.object_state(name,'正在导出…' if stage=='export' else '已导出，待核验','done' if stage=='exported' else '')
                self.status.set(f'Origin 导出：{index} / {total} · {name}')
                self.set_progress(5+70*(index-(stage=='export'))/max(1,total))
            elif stage=='skipped':
                self.current_object=None
                self.status.set(f'已跳过 {name}，继续处理后续图表…');self.set_progress(5+70*index/max(1,total))
            elif stage=='render':self.status.set('PowerPoint 预览验证：'+name);self.set_progress(90)
            else:
                self.status.set(labels.get(stage,event.get('message','正在处理…')))
                if stage=='packaging':self.set_progress(80)
            if self.cancelled:self.status.set('已请求取消；当前图表或验证步骤结束后停止。')
        elif kind=='skipped_results':
            for row in event['skipped']:self.mark_skipped(row)
        elif kind=='export_results':
            for r in event['results']:
                self.rows[r['name']].update(r)
                self.object_state(r['name'],'已转换 · 建议复核' if r['warnings'] else '已转换，验证通过','review' if r['warnings'] else 'done')
        elif kind=='saved':
            self.result_path=event['output'];self.open_result.configure(state='normal');self.set_progress(100)
            count=event.get('warnings',0)
            skipped=event.get('skipped',0)
            if skipped:
                self.status.set(f"完成：转换 {event['converted']} 个，跳过 {skipped} 个（保留原 Origin 对象）。" + (f'{count} 个转换图表建议复核。' if count else ''))
            else:
                self.status.set(f'完成：已另存为新 PPT。{count} 个图表建议查看前后对照。' if count else '完成：所有图表已转换，原文件保留。')
            self.log_line('已保存：'+self.result_path)
        elif kind=='no_success':
            self.no_success=True;self.set_progress(100)
            self.status.set('所有图表均已跳过，未生成新 PPT。原文件保留，请查看跳过原因。')
        elif kind=='report':
            self.report_path=event['report'];self.open_report.configure(state='normal')
            if self.work_dir:self.open_folder.configure(state='normal')
        elif kind=='error':
            self.failed=True;self.status.set('转换未完成，未生成最终 PPT。请查看下方错误或检查报告。')
            if self.current_object:self.object_state(self.current_object,'导出失败 · 请查看报告','error')
            self.log_line(event['message'])
        elif kind=='cancelled':
            self.cancelled=True;self.status.set('已取消。原文件保留，未生成最终 PPT。')
        elif kind=='exit':
            self.set_busy(False)
            if self.mode=='scan' and not self.failed and not self.cancelled and event['code']==0:
                self.status.set('识别完成，可开始转换。双击已转换的行可打开图表前后对照。' if self.rows else '未发现可转换的 Origin 嵌入图表。普通图片和外部链接不在自动转换范围内。')
                self.set_progress(100)
            elif event['code']!=0 and not self.failed and not self.cancelled:
                self.failed=True;self.status.set('程序异常退出。请检查运行环境或报告。')
            elif self.mode=='convert' and not self.result_path and not self.failed and not self.cancelled and not self.no_success:
                self.status.set('未发现需要转换的 Origin 对象，未生成新文件。')
            if self.close_when_done:self.destroy()

    def cancel(self):
        if self.busy:
            self.cancelled=True;self.cancel_path.write_text('cancel',encoding='ascii')
            self.cancel_button.configure(state='disabled')
            self.status.set('已请求取消；当前图表或验证步骤结束后停止。')

    def on_close(self):
        if self.busy:
            self.close_when_done=True;self.cancel()
        else:self.destroy()

    def open_path(self,path):
        if not path:return
        try:os.startfile(str(path))
        except Exception as exc:messagebox.showerror('无法打开',str(exc),parent=self)

    def open_comparison(self,event=None):
        selected=self.tree.selection()
        if selected and self.work_dir:
            path=Path(self.work_dir)/(selected[0]+'_comparison.png')
            if path.exists():self.open_path(path)
            elif self.rows[selected[0]].get('status')=='skipped':
                messagebox.showinfo('跳过原因',self.rows[selected[0]].get('error','未能生成对照图。'),parent=self)

    def show_report(self):
        if not self.report_path:return
        window=tk.Toplevel(self);window.title('检查报告');window.geometry('860x600')
        text=tk.Text(window,font=('Consolas',10),wrap='word');text.pack(fill='both',expand=True)
        try:text.insert('1.0',Path(self.report_path).read_text(encoding='utf-8'))
        except Exception as exc:text.insert('1.0',str(exc))
        text.configure(state='disabled')


def main():
    app=OriginTifApp()
    if '--smoke-test' in sys.argv:
        app.withdraw();app.update_idletasks()
        require=lambda condition: (_ for _ in ()).throw(RuntimeError('GUI smoke check failed')) if not condition else None
        app.handle({'event':'inventory','slides':2,'work_dir':str(app.temp),'rows':[{'name':'demo','slide':2,'id':'7','prog_id':'Origin95.Graph','display_inches':[2,1],'group_depth':1}]})
        require(len(app.tree.get_children())==1)
        app.mode='convert';app.handle({'event':'progress','stage':'export','index':1,'total':1,'name':'demo'})
        require(app.tree.set('demo','state')=='正在导出…')
        app.handle({'event':'export_results','results':[{'name':'demo','warnings':['review']}]})
        require('复核' in app.tree.set('demo','state'))
        app.handle({'event':'skipped_results','skipped':[{'name':'demo','slide':2,'id':'7','status':'skipped','error':'test mismatch'}]})
        app.handle({'event':'progress','stage':'export','index':2,'total':2,'name':'next'})
        require('跳过' in app.tree.set('demo','state'))
        app.handle({'event':'saved','output':'test.pptx','converted':1,'skipped':1,'warnings':0})
        require('转换 1 个，跳过 1 个' in app.status.get())
        app.handle({'event':'no_success','skipped':1});app.handle({'event':'exit','code':0})
        require('所有图表均已跳过' in app.status.get())
        app.destroy();return
    if len(sys.argv)>1 and Path(sys.argv[1]).is_file() and Path(sys.argv[1]).suffix.lower()=='.pptx':
        source=Path(sys.argv[1]).resolve()
        app.source.set(str(source));app.destination.set(str(app.proposed_output(source)))
        app.after(200,app.scan)
    app.mainloop()


if __name__=='__main__':main()
