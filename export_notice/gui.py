"""本地桌面界面；后台分析不阻塞界面，生成前重新读取文件避免旧预览。"""
from datetime import date
from pathlib import Path
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from .service import analyze_batch, generate


def app_home():
    return Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent.parent


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('出口技术通知助手')
        self.geometry('1120x760')
        self.minsize(900, 620)
        self.home = app_home()
        self.busy = False
        self.paths = {}
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill='both', expand=True)
        ttk.Label(outer, text='出口技术通知助手', font=('Microsoft YaHei', 18, 'bold')).pack(anchor='w')
        ttk.Label(outer, text='选择一批文件 → 检查负责产品与要求 → 生成 Word → 人工审核').pack(anchor='w', pady=(4,14))
        form=ttk.Frame(outer);form.pack(fill='x')
        self.path_row(form,0,'批次文件夹','batch',self.home/'examples'/'样例批次',True)
        self.path_row(form,1,'产品资料表','master',self.home/'data'/'产品资料表.xlsx',False)
        self.path_row(form,2,'Word 模板','template',self.home/'data'/'自动化模板.docx',False)
        self.path_row(form,3,'输出文件夹','output',self.home/'outputs',True)
        metadata=ttk.Frame(outer);metadata.pack(fill='x',pady=10)
        ttk.Label(metadata,text='通知编号（按公司规则填写）').pack(side='left')
        self.number=tk.StringVar();ttk.Entry(metadata,textvariable=self.number,width=24).pack(side='left',padx=8)
        ttk.Label(metadata,text='实施日期').pack(side='left')
        self.effective=tk.StringVar(value=date.today().isoformat())
        ttk.Entry(metadata,textvariable=self.effective,width=16).pack(side='left',padx=8)
        actions=ttk.Frame(outer);actions.pack(fill='x',pady=(0,12))
        self.preview_button=ttk.Button(actions,text='读取并预览',command=lambda:self.run(False))
        self.preview_button.pack(side='left')
        self.generate_button=ttk.Button(actions,text='校验并生成通知',command=lambda:self.run(True))
        self.generate_button.pack(side='left',padx=8)
        ttk.Button(actions,text='打开产品资料表',command=lambda:self.open_path(self.paths['master'].get())).pack(side='left',padx=8)
        ttk.Button(actions,text='打开输出目录',command=lambda:self.open_path(self.paths['output'].get())).pack(side='left')
        self.status=tk.StringVar(value='每次选择一个批次文件夹；不负责的图号只列入排除清单。')
        ttk.Label(outer,textvariable=self.status,wraplength=1000).pack(anchor='w')
        tabs=ttk.Notebook(outer);tabs.pack(fill='both',expand=True,pady=(8,0))
        product_tab=ttk.Frame(tabs);tabs.add(product_tab,text='负责产品')
        columns=['图号','中文名称','英文名称','数量','包装模数','单箱数量','纸箱总数','工位号','纸箱尺寸']
        self.tree=ttk.Treeview(product_tab,columns=columns,show='headings')
        for name in columns:
            self.tree.heading(name,text=name)
            self.tree.column(name,width=125 if name in ['图号','英文名称','纸箱尺寸'] else 95,minwidth=65)
        scroll=ttk.Scrollbar(product_tab,orient='horizontal',command=self.tree.xview)
        self.tree.configure(xscrollcommand=scroll.set)
        self.tree.pack(fill='both',expand=True);scroll.pack(fill='x')
        self.texts={}
        for title in ['出口要求','排除产品','校验与字段来源']:
            frame=ttk.Frame(tabs);tabs.add(frame,text=title)
            text=tk.Text(frame,wrap='word',font=('Microsoft YaHei',10))
            bar=ttk.Scrollbar(frame,command=text.yview);text.configure(yscrollcommand=bar.set)
            bar.pack(side='right',fill='y');text.pack(fill='both',expand=True)
            text.configure(state='disabled');self.texts[title]=text

    def path_row(self,parent,row,label,key,initial,directory):
        variable=tk.StringVar(value=str(initial));self.paths[key]=variable
        ttk.Label(parent,text=label,width=12).grid(row=row,column=0,sticky='w',pady=3)
        ttk.Entry(parent,textvariable=variable).grid(row=row,column=1,sticky='ew',padx=8)
        def browse():
            value=filedialog.askdirectory() if directory else filedialog.askopenfilename(filetypes=[('Excel/Word 文件','*.xlsx *.xls *.docx'),('所有文件','*.*')])
            if value: variable.set(value)
        ttk.Button(parent,text='选择…',command=browse).grid(row=row,column=2)
        parent.columnconfigure(1,weight=1)

    def open_path(self,path):
        try:
            target=Path(path)
            if not target.exists():
                if target.suffix: raise FileNotFoundError('文件不存在')
                target.mkdir(parents=True)
            os.startfile(str(target))
        except Exception as exc: messagebox.showerror('打开失败',str(exc))

    def run(self,write):
        if self.busy: return
        try:
            date.fromisoformat(self.effective.get().strip())
        except ValueError:
            messagebox.showerror('日期格式','实施日期请填写 YYYY-MM-DD');return
        paths={key:value.get().strip() for key,value in self.paths.items()}
        number=self.number.get().strip();effective=self.effective.get().strip()
        self.busy=True
        self.preview_button.state(['disabled']);self.generate_button.state(['disabled'])
        self.status.set('正在读取和校验…')
        def task():
            try:
                aliases=self.home/'data'/'字段别名.json'
                analysis=analyze_batch(paths['batch'],paths['master'],aliases if aliases.exists() else None)
                result=generate(analysis,paths['template'],paths['output'],Path(paths['batch']).name,number,effective) if write else None
                self.after(0,lambda:self.show_result(analysis,result))
            except Exception as exc:
                self.after(0,lambda e=str(exc):self.failure(e))
        threading.Thread(target=task,daemon=True).start()

    def put(self,key,value):
        text=self.texts[key];text.configure(state='normal')
        text.delete('1.0','end');text.insert('1.0',value);text.configure(state='disabled')

    def show_result(self,analysis,result):
        self.tree.delete(*self.tree.get_children())
        details=[]
        for row in analysis.rows:
            v=row.resolved
            self.tree.insert('', 'end',values=[row.standard_code,v.get('中文名称'),v.get('英文名称'),row.quantity,
                             v.get('包装模数'),v.get('单箱数量'),row.carton_count,v.get('工位号'),v.get('纸箱尺寸')])
            record=analysis.master.get(row.standard_code)
            def field_source(key):
                raw=row.chinese_name if key=='中文名称' else row.fields.get(key)
                if raw not in (None,''): return '本批 Excel'
                if key=='工位号': return 'Excel 未说明，按规则写 /'
                return '产品资料表'
            details.append(row.location+'\n'+'\n'.join(f'  {k}：{x}（{field_source(k)}）' for k,x in v.items()))
        req=analysis.requirements
        self.put('出口要求',f'合同号：{req.contract_no}\n车型：{req.model}\n国家：{req.country}\nLOGO类型：{req.logo_type}\n'
                f'交货地点：{req.delivery_place}\n交货时间：{req.delivery_time}\n包装模数：{req.modulus}\n标签网址：{req.label_url}\n\n原始要求：\n{req.source_text}')
        self.put('排除产品','\n'.join(f"{r['source_file']} / {r['source_sheet']} / 第{r['source_row']}行：{r['raw_code']} → {r['standard_code']}；{r['reason']}" for r in analysis.excluded) or '无')
        self.put('校验与字段来源','\n'.join(analysis.errors+analysis.warnings)+'\n\n'+'\n\n'.join(details))
        self.finish()
        if result:
            doc,report,errors=result
            if errors:
                self.status.set(f'已阻断通知生成，报告保存至：{report}')
                messagebox.showerror('校验未通过','\n'.join(errors[:12])+f'\n\n报告：{report}')
            else:
                self.status.set(f'已生成，待人工审核：{doc}')
                messagebox.showinfo('生成完成',f'Word：{doc}\n报告：{report}\n\n请审核后下发。')
        else:
            self.status.set(f'已读取：负责产品 {len(analysis.rows)}，排除 {len(analysis.excluded)}，错误 {len(analysis.errors)}。')

    def failure(self,error):
        self.finish();self.status.set('读取失败，请检查所选文件')
        messagebox.showerror('读取失败',error)

    def finish(self):
        self.busy=False
        self.preview_button.state(['!disabled']);self.generate_button.state(['!disabled'])


def main():
    App().mainloop()
