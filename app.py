"""双击桌面入口；--self-check 用于打包后的依赖和业务校验。"""
import json
import sys
from pathlib import Path

def main():
    if len(sys.argv)>1 and sys.argv[1]=='--self-check':
        from export_notice.gui import App, app_home
        from export_notice.service import analyze_batch, generate
        import tempfile
        home=app_home()
        app=App()
        app.withdraw()
        app.update()
        app.destroy()
        result=analyze_batch(home/'examples'/'样例批次',home/'data'/'产品资料表.xlsx')
        with tempfile.TemporaryDirectory() as td:
            doc,report,errors=generate(result,home/'data'/'自动化模板.docx',td,'样例','自检','2026-10-12')
            outcome={'ok':not errors and doc is not None and report.exists(),
                     'errors':errors,'owned':len(result.rows),'excluded':len(result.excluded),
                     'cartons':{r.standard_code:r.carton_count for r in result.rows}}
        Path(sys.argv[2]).write_text(json.dumps(outcome,ensure_ascii=False,indent=2),encoding='utf-8')
        return
    from export_notice.gui import main as gui_main
    gui_main()

if __name__=='__main__':
    main()
