"""
xls_convert.py — 独立进程的 xls→xlsx 转换

由 engine.py 通过 subprocess 调用，带超时保护。
Excel COM 挂死时父进程可 kill 此进程 + 杀 Excel.exe

用法: python -m tax_audit_engine.xls_convert <输入.xls> <输出.xlsx>
"""

import sys, os


def main():
    import win32com.client as win32

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(sys.argv[1]))
        wb.SaveAs(os.path.abspath(sys.argv[2]), FileFormat=51)
        wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"XLS_CONVERT_ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
