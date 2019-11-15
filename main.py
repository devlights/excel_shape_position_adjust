#################################################################
# 指定されたフォルダ配下のExcelを開いていき、画像が指定位置に貼付けされていないファイルを出力or調整します。
#
# 実行には、以下のライブラリが必要です.
#   - win32com
#     - $ python -m pip install pywin32
#
# [参考にした情報]
#   - https://www.sejuku.net/blog/23647
#################################################################
import argparse


# noinspection SpellCheckingInspection
def go(target_dir: str, base_position: float, report_only: bool):
    import pathlib

    import pywintypes
    import win32com.client

    excel_dir = pathlib.Path(target_dir)
    if not excel_dir.exists():
        print(f'target directory not found [{target_dir}]')
        return

    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True

        for f in excel_dir.glob('**/*.xlsx'):
            abs_path = str(f)
            try:
                wb = excel.Workbooks.Open(abs_path)
                wb.Activate()
            except pywintypes.com_error as err:
                print(err)
                continue

            try:
                sheets_count = wb.Sheets.Count
                for sheet_index in range(0, sheets_count):
                    ws = wb.Worksheets(sheet_index + 1)
                    ws.Activate()
                    for sh in ws.Shapes:
                        if base_position <= sh.Left:
                            if report_only:
                                print(f'{abs_path}-{ws.Name}')
                            else:
                                sh.Left = base_position
                if not report_only:
                    wb.Save()
                wb.Saved = True
            finally:
                wb.Close()
    finally:
        excel.Quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        usage='python main.py -d /path/to/excel/dir -p base-left-position(e.g. 100.0) [-r]',
        description='指定されたフォルダ配下のExcelを開いていき、画像が左端に貼付けされていないファイルを出力します。',
        add_help=True
    )

    parser.add_argument('-d', '--directory', help='対象ディレクトリ', required=True)
    parser.add_argument('-p', '--position', help='基準となるShape.Leftの値', type=float, default=100.0)
    parser.add_argument('-r', '--report', help='情報のみ出力して変更はしない', action='store_true')

    args = parser.parse_args()

    go(args.directory, args.position, args.report)
