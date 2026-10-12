import argparse
from pathlib import Path
from .service import analyze_batch, generate

def main():
    parser=argparse.ArgumentParser(description='出口技术通知批量生成')
    parser.add_argument('folder')
    parser.add_argument('--master',required=True)
    parser.add_argument('--template',required=True)
    parser.add_argument('--output',default='outputs')
    parser.add_argument('--aliases')
    parser.add_argument('--number',default='')
    parser.add_argument('--date',default='')
    args=parser.parse_args()
    analysis=analyze_batch(args.folder,args.master,args.aliases)
    doc,report,errors=generate(analysis,args.template,args.output,Path(args.folder).name,args.number,args.date)
    print(f'负责产品：{len(analysis.rows)}；排除：{len(analysis.excluded)}')
    print(f'报告：{report}')
    if errors:
        for error in errors: print(error)
        raise SystemExit(1)
    print(f'通知：{doc}')

if __name__=='__main__':main()
