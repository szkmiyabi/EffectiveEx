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
        Dictionary<string, int> dict = null;

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

        //キー存在判定
        private Boolean isExistsKey(string key)
        {
            Boolean flg = false;
            foreach(string str in dict.Keys)
            {
                if (str == key) flg = true;
            }
            return flg;
        }

        //辞書を生成
        private void adjustDictionary(string key)
        {
            if(dict.Count == 0)
            {
                dict.Add(key, 1);
            }
            else if (!isExistsKey(key))
            {
                dict.Add(key, 1);
            }
            else if (isExistsKey(key))
            {
                int val = (int)dict[key];
                val++;
                dict[key] = val;
            }
        }

        //値検索集計
        private void searchValsResult()
        {
            if (sheetNameCombo.Text == "" || columnValues.Text == "" || skipRowNumbers.Text == "")
            {
                MessageBox.Show("シート名あるいは列番号あるいはスキップ行数あるいは検索条件が入力されていません");
                return;
            }

            //コンボ選択値をアクティブシートにする
            initCurrentWorksheet(sheetNameCombo.Text);

            dict = new Dictionary<string, int>();
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

                if ((searchValResutWithoutCheck.Checked == true))
                {
                    int without_cond_col = Int32.Parse(withoutConditionColNum.Text);
                    var inchk_val = currentWs.Cell(i, without_cond_col).Value;
                    Type inchk_t = currentWs.Cell(i, without_cond_col).Value.GetType();
                    if (inchk_t.Equals(typeof(double)) || inchk_t.Equals(typeof(int)) || inchk_t.Equals(typeof(float)))
                    {
                        inchk_val = inchk_val.ToString();
                    }
                    else if (inchk_t.Equals(typeof(DateTime)))
                    {
                        inchk_val = inchk_val.ToString();
                    }
                    else if (inchk_t.Equals(typeof(ClosedXML.Excel.XLHyperlink)))
                    {
                        inchk_val = (ClosedXML.Excel.XLHyperlink)chk_val;
                        inchk_val = chk_val.ToString();
                    }
                    else
                    {
                        inchk_val = (string)inchk_val;
                    }
                    if (!isMatchRow(vals, (string)inchk_val)) continue;
                }

                adjustDictionary((string)chk_val);
            }


            foreach(string str in dict.Keys)
            {
                reportText.AppendText(str + "\t" + dict[str].ToString() + "\r\n");
            }
            
        }

        //罫線自動描画
        private void tableBordered()
        {

            //コンボ選択値をアクティブシートにする
            initCurrentWorksheet(sheetNameCombo.Text);

            writeLog("罫線描画を開始します....");

            int r = getStartRow();
            int rx = getEndRow();
            int cx = getEndCol();

            for(int i=r; i<=rx; i++)
            {
                writeLog(i + "行目の処理....");
                for(int j=1; j<=cx; j++)
                {
                    currentWs.Cell(i, j).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    currentWs.Cell(i, j).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    currentWs.Cell(i, j).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    currentWs.Cell(i ,j).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    currentWs.Cell(i, j).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);

                }
            }

            writeLog("罫線描画の処理が完了しました....");

        }

        //LibraPlusの検査結果一覧表を書式設定する（LRPフォーマット）
        private void lpReportFormat()
        {
            //罫線描画
            tableBordered();

            //コンボ選択値をアクティブシートにする
            initCurrentWorksheet(sheetNameCombo.Text);

            writeLog("LPRフォーマット処理を開始します....");

            int r = getStartRow();
            int rx = getEndRow();
            int cx = getEndCol();

            int sv_index = 6;

            for (int i = r; i <= rx; i++)
            {

                writeLog(i + "行目の処理....");

                for (int j = 1; j <= cx; j++)
                {

                    //header cell
                    if (i == 1)
                    {
                        currentWs.Cell(i, j).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        currentWs.Cell(i, j).Style.Font.Bold = true;
                    }
                    //data cell
                    else
                    {

                        string sv_val = (string)currentWs.Cell(i, sv_index).Value;

                        if (sv_val == "はい")
                        {
                            currentWs.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(0x89FFFF);
                        }
                        else if (sv_val == "はい(注記)")
                        {
                            currentWs.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(0x99FF99);
                        }
                        else if (sv_val == "いいえ")
                        {
                            currentWs.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(0xFFB3B3);
                        }
                        else if (sv_val == "なし")
                        {
                            currentWs.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(0xDDDDDD);
                        }
                    }

                }
            }

            writeLog("LPRフォーマット処理が完了しました....");


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
