using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace EffectiveEx
{
    partial class Form1
    {
        //カレントWorkbookをセット
        private void initCurrentBook()
        {
            currentWb = new XLWorkbook(currentWbPath);
        }

        //カレントWorkbookのカレントSheetをセット
        private void initCurrentWorksheet(int idx)
        {
            currentWs = currentWb.Worksheet(idx);
        }

        //カレントWorkbookのカレントSheetをセット（オーバーライド）
        private void initCurrentWorksheet(string sname)
        {
            currentWs = currentWb.Worksheets.Worksheet(sname);
        }

        //対象シートコンボのセットアップ
        private void initTargetWorksheetCombo()
        {
            initCurrentBook();
            sheetNameCombo.Items.Clear();

            var sheets = currentWb.Worksheets;
            foreach (IXLWorksheet st in sheets)
            {
                string sname = st.Name;
                sheetNameCombo.Items.Add(sname);
            }

            sheetNameCombo.Enabled = true;
        }

        //カレントSheetのデータ範囲開始行を取得
        private int getStartRow()
        {
            var first = currentWs.FirstCellUsed();
            return first.WorksheetRow().RowNumber();
        }

        //カレントSheetのデータ範囲最終行を取得
        private int getEndRow()
        {
            var last = currentWs.LastCellUsed();
            return last.WorksheetRow().RowNumber();
        }

        //カレントSheetのデータ範囲開始列を取得
        private int getStartCol()
        {
            var first = currentWs.FirstCellUsed();
            return first.WorksheetColumn().ColumnNumber();
        }

        //カレントSheetのデータ範囲最終列を取得
        private int getEndCol()
        {
            var last = currentWs.LastCellUsed();
            return last.WorksheetColumn().ColumnNumber();
        }

        //条件に一致する行判定
        private Boolean isMatchRow(List<string> data, string key)
        {
            Boolean flg = false;
            foreach(string row in data)
            {
                if (row == key) flg = true;
            }
            return flg;
        }

        //条件に一致する行数取得（スキップ行数は含まない）
        private int getHitsRowCount()
        {
            //カウンタ
            int cnt = 0;

            List<string> vals = getSearchValues();
            int icx = Int32.Parse(columnValues.Text);
            int skip_nm = Int32.Parse(skipRowNumbers.Text);
            int r = getStartRow();
            int rx = getEndRow();
            int start_nm = r + skip_nm;

            for (int i = start_nm; i <= rx; i++)
            {
                //条件判定のためのデータ取り出し処理
                var chk_val = currentWs.Cell(i, icx).Value;
                Type chk_t = currentWs.Cell(i, icx).Value.GetType();
                if (chk_t.Equals(typeof(double)) || chk_t.Equals(typeof(int)) || chk_t.Equals(typeof(float)))
                {
                    chk_val = chk_val.ToString();
                }
                else if (chk_t.Equals(typeof(DateTime)))
                {
                    chk_val = chk_val.ToString();
                }
                else if (chk_t.Equals(typeof(ClosedXML.Excel.XLHyperlink)))
                {
                    chk_val = (ClosedXML.Excel.XLHyperlink)chk_val;
                    chk_val = chk_val.ToString();
                }
                else
                {
                    chk_val = (string)chk_val;
                }

                //条件に一致したらカウンタをインクリメント
                if (isMatchRow(vals, (string)chk_val))
                {
                    cnt++;
                }
            }
            return cnt;
        }

        //行削除
        private void deleteRow()
        {
            List<string> vals = getSearchValues();
            int icx = Int32.Parse(columnValues.Text);
            int skip_nm = Int32.Parse(skipRowNumbers.Text);
            int start_nm = 1 + skip_nm;

            for(int i=start_nm; i<=getEndRow(); i++)
            {
                //条件判定のためのデータ取り出し処理
                var chk_val = currentWs.Cell(i, icx).Value;
                Type chk_t = currentWs.Cell(i, icx).Value.GetType();
                if (chk_t.Equals(typeof(double)) || chk_t.Equals(typeof(int)) || chk_t.Equals(typeof(float)))
                {
                    chk_val = chk_val.ToString();
                }
                else if (chk_t.Equals(typeof(DateTime)))
                {
                    chk_val = chk_val.ToString();
                }
                else if (chk_t.Equals(typeof(ClosedXML.Excel.XLHyperlink)))
                {
                    chk_val = (ClosedXML.Excel.XLHyperlink)chk_val;
                    chk_val = chk_val.ToString();
                }
                else
                {
                    chk_val = (string)chk_val;
                }

                //条件に一致しない時に削除実行し関数を抜ける（削除の度に全体データ範囲行数が減ってしまうため）
                if (!isMatchRow(vals, (string)chk_val))
                {
                    writeLog("行を削除しました");
                    currentWs.Row(i).Delete();
                    break;
                }
            }
        }

        //行削除のラッパー関数
        private void deleteRowByCondition()
        {
            if (sheetNameCombo.Text == "" || columnValues.Text == "" || skipRowNumbers.Text == "" || searchValues.Text == "")
            {
                MessageBox.Show("シート名あるいは列番号あるいはスキップ行数あるいは検索条件が入力されていません");
                return;
            }

            //コンボ選択値をアクティブシートにする
            initCurrentWorksheet(sheetNameCombo.Text);

            int skip_nm = Int32.Parse(skipRowNumbers.Text);

            //残す行数＋スキップ行数を取得
            int during = getHitsRowCount() + skip_nm;

            //データ範囲行数が残す行数＋スキップ行数に達するまでループ
            while(getEndRow() != during)
            {
                //ひたすら行を削除する
                deleteRow();
            }

            //保存の段取り
            string new_filename = Path.GetFileNameWithoutExtension(currentWbPath);
            string ext = Path.GetExtension(currentWbPath);
            string new_savepath = saveDirPath + new_filename + "_out_" + fetch_filename_logtime() + ext;

            //別名保存
            currentWb.SaveAs(new_savepath);
            writeLog("処理が完了しました。出力ファイル：" + new_savepath);
        }
    }
}
