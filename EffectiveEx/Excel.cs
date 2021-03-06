﻿using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;
using System.Diagnostics;

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

            //デリゲートインスタンス
            _columnValuesVal __columnValuesVal = columnValuesVal;
            _skipRowNumbersVal __skipRowNumbersVal = skipRowNumbersVal;

            //カウンタ
            int cnt = 0;

            List<string> vals = getSearchValues();
            int icx = (int)this.Invoke(__columnValuesVal);
            int skip_nm = (int)this.Invoke(__skipRowNumbersVal);
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
        private async Task deleteRow()
        {

            await Task.Run(() =>
            {
                //デリゲートインスタンス
                _columnValuesVal __columnValuesVal = columnValuesVal;
                _skipRowNumbersVal __skipRowNumbersVal = skipRowNumbersVal;
                _writeLog __writeLog = writeLog;

                List<string> vals = getSearchValues();
                int icx = (int)this.Invoke(__columnValuesVal);
                int skip_nm = (int)this.Invoke(__skipRowNumbersVal);
                int start_nm = 1 + skip_nm;

                for (int i = start_nm; i <= getEndRow(); i++)
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
                        this.Invoke(__writeLog, "行を削除しました");
                        currentWs.Row(i).Delete();
                        break;
                    }
                }
            });

        }

        //行削除のラッパー関数
        private async Task deleteRowByCondition()
        {

            await Task.Run(() =>
            {
                //デリゲートインスタンス
                _sheetNameComboVal __sheetNameComboVal = sheetNameComboVal;
                _skipRowNumbersVal __skipRowNumbersVal = skipRowNumbersVal;
                _searchValuesVal __searchValuesVal = searchValuesVal;
                _writeLog __writeLog = writeLog;

                if ((string)this.Invoke(__sheetNameComboVal) == "" || (string)this.Invoke(__searchValuesVal) == "")
                {
                    MessageBox.Show("シート名あるいは検索条件が入力されていません");
                    return;
                }

                //コンボ選択値をアクティブシートにする
                initCurrentWorksheet((string)this.Invoke(__sheetNameComboVal));

                int skip_nm = (int)this.Invoke(__skipRowNumbersVal);

                //残す行数＋スキップ行数を取得
                int during = getHitsRowCount() + skip_nm;

                //データ範囲行数が残す行数＋スキップ行数に達するまでループ
                while (getEndRow() != during)
                {
                    //ひたすら行を削除する
                    deleteRow().Wait();
                }

                //保存の段取り
                string new_filename = Path.GetFileNameWithoutExtension(currentWbPath);
                string ext = Path.GetExtension(currentWbPath);
                string new_savepath = getCurrentFileWorkPath() + new_filename + "_行削除_" + fetch_filename_logtime() + ext;

                //別名保存
                currentWb.SaveAs(new_savepath);
                this.Invoke(__writeLog, "処理が完了しました。出力ファイル：" + new_savepath);
            });

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
        private async Task searchValsResult()
        {

            await Task.Run(() =>
            {
                //デリゲートインスタンス
                _sheetNameComboVal __sheetNameComboVal = sheetNameComboVal;
                _skipRowNumbersVal __skipRowNumbersVal = skipRowNumbersVal;
                _searchValuesVal __searchValuesVal = searchValuesVal;
                _columnValuesVal __columnValuesVal = columnValuesVal;
                _withoutConditionColNumVal __withoutConditionColNumVal = withoutConditionColNumVal;
                _searchResultWithoutCheckIsChecked __searchResultWithoutCheckIsChecked = searchResultWithoutCheckIsChecked;
                _reportTextAppendTextThis __reportTextAppendTextThis = reportTextAppendTextThis;
                _writeLog __writeLog = writeLog;

                if ((string)this.Invoke(__sheetNameComboVal) == "")
                {
                    MessageBox.Show("シート名が入力されていません");
                    return;
                }

                //コンボ選択値をアクティブシートにする
                initCurrentWorksheet((string)this.Invoke(__sheetNameComboVal));

                dict = new Dictionary<string, int>();
                List<string> vals = getSearchValues();
                int icx = (int)this.Invoke(__columnValuesVal);
                int skip_nm = (int)this.Invoke(__skipRowNumbersVal);
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

                    if ((Boolean)this.Invoke(__searchResultWithoutCheckIsChecked) == true)
                    {
                        int without_cond_col = (int)this.Invoke(__withoutConditionColNumVal);
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

                foreach (string str in dict.Keys)
                {
                    this.Invoke(__reportTextAppendTextThis, str, dict[str]);
                }

            });
            
        }

        //罫線自動描画
        private async Task tableBorderedAsync()
        {
            await Task.Run(() =>
            {
                //共通デリゲートインスタンス
                _sheetNameComboVal __sheetNameComboVal = sheetNameComboVal;
                _writeLog __writeLog = writeLog;

                //コンボ選択値をアクティブシートにする
                initCurrentWorksheet((string)this.Invoke(__sheetNameComboVal));

                this.Invoke(__writeLog, "罫線描画を開始します....");

                int r = getStartRow();
                int rx = getEndRow();
                int cx = getEndCol();

                for (int i = r; i <= rx; i++)
                {
                    this.Invoke(__writeLog, i + "行目の処理....");
                    for (int j = 1; j <= cx; j++)
                    {
                        currentWs.Cell(i, j).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        currentWs.Cell(i, j).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        currentWs.Cell(i, j).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                        currentWs.Cell(i, j).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                        currentWs.Cell(i, j).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);

                    }
                }

                this.Invoke(__writeLog, "罫線描画の処理が完了しました....");
            });

        }

        //LibraPlusの検査結果一覧表を書式設定する（LRPフォーマット）
        private async Task lpReportFormat()
        {
            await Task.Run(() =>
            {
                //共通デリゲートインスタンス
                _sheetNameComboVal __sheetNameComboVal = sheetNameComboVal;
                _writeLog __writeLog = writeLog;


                //罫線描画（処理完了まで待機）
                tableBorderedAsync().Wait();

                //コンボ選択値をアクティブシートにする
                initCurrentWorksheet((string)this.Invoke(__sheetNameComboVal));

                this.Invoke(__writeLog, "LPRフォーマット処理を開始します....");

                int r = getStartRow();
                int rx = getEndRow();
                int cx = getEndCol();

                int sv_index = 6;

                for (int i = r; i <= rx; i++)
                {

                    this.Invoke(__writeLog, i + "行目の処理....");

                    for (int j = 1; j <= cx; j++)
                    {

                        //header cell
                        if (i == 1)
                        {
                            currentWs.Cell(i, j).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            currentWs.Cell(i, j).Style.Font.Bold = true;

                            //列幅指定
                            double[] colWidthArr = { 8.4, 8.4, 13.9, 45.6, 24, 8.4, 6.1, 45.2, 45.2, 45.2, 8.4, 8.4 };
                            int wcnt = 1;
                            foreach(double wc in colWidthArr)
                            {
                                currentWs.Column(wcnt).Width = wc;
                                wcnt++;
                            }
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

                this.Invoke(__writeLog, "LPRフォーマット処理が完了しました....");

                //保存の段取り
                string new_filename = Path.GetFileNameWithoutExtension(currentWbPath);
                string ext = Path.GetExtension(currentWbPath);
                string new_savepath = getCurrentFileWorkPath() + new_filename + "_LPR書式整形_" + fetch_filename_logtime() + ext;

                //別名保存
                currentWb.SaveAs(new_savepath);
                this.Invoke(__writeLog, "処理が完了しました。出力ファイル：" + new_savepath);

            });

        }

        //Excelワークシートをプレビュー
        private async Task previewWorksheetWrap()
        {
            await Task.Run(() =>
            {
                _previewWorksheet __previewWorksheet = previewWorksheet;
                this.Invoke(__previewWorksheet);
            });
        }
        //DataGridFormのためのデリゲート
        private delegate void _previewWorksheet();
        private void previewWorksheet()
        {
            //共通デリゲートインスタンス
            _sheetNameComboVal __sheetNameComboVal = sheetNameComboVal;
            _writeLog __writeLog = writeLog;

            if ((string)this.Invoke(__sheetNameComboVal) == "")
            {
                MessageBox.Show("シート名が入力されていません");
                return;
            }

            //データテーブル（DataGridFormへの引数となる）
            DataTable tbl = new DataTable("worksheetPreviewTable");

            //コンボ選択値をアクティブシートにする
            initCurrentWorksheet((string)this.Invoke(__sheetNameComboVal));

            int r = getStartRow();
            int rx = getEndRow();
            int cx = getEndCol();

            //DataGridViewのヘッダーを設定
            for(int cnt = 1; cnt<=cx; cnt++)
            {
                tbl.Columns.Add("列" + cnt.ToString());
            }

            //Worksheetを走査してDataTable組み立て
            for(int i=r; i<=rx; i++)
            {
                //列データ配列（Listでなく配列でないとだめなので）
                string[] cols = new string[cx];

                for (int j=1; j<=cx; j++)
                {
                    var cell_val = currentWs.Cell(i, j).Value;
                    Type chk_t = currentWs.Cell(i, j).Value.GetType();
                    if (chk_t.Equals(typeof(double)) || chk_t.Equals(typeof(int)) || chk_t.Equals(typeof(float)))
                    {
                        cell_val = cell_val.ToString();
                    }
                    else if (chk_t.Equals(typeof(DateTime)))
                    {
                        cell_val = cell_val.ToString();
                    }
                    else if (chk_t.Equals(typeof(ClosedXML.Excel.XLHyperlink)))
                    {
                        cell_val = (ClosedXML.Excel.XLHyperlink)cell_val;
                        cell_val = cell_val.ToString();
                    }
                    else
                    {
                        cell_val = (string)cell_val;
                    }
                    cols[j-1] = (string)cell_val;
                }

                //組み上がった1行分のデータをデータテーブルに追加
                tbl.Rows.Add(cols);
            }

            //DataGridFormを表示
            data_grid_form.init(tbl, "プレビュー");
            data_grid_form.Show();

        }

        //結合セルを強調する
        private async Task mergedCellAttention()
        {
            await Task.Run(() =>
            {
                //共通デリゲートインスタンス
                _sheetNameComboVal __sheetNameComboVal = sheetNameComboVal;
                _writeLog __writeLog = writeLog;

                //コンボ選択値をアクティブシートにする
                initCurrentWorksheet((string)this.Invoke(__sheetNameComboVal));

                this.Invoke(__writeLog, "結合セルの強調処理を開始します....");

                
                var mc = currentWs.MergedRanges;
                int cnt = 0;
                foreach (IXLRange rg in mc)
                {
                    int rc = rg.RowCount();
                    int cc = rg.ColumnCount();
                    string comm = "";

                    if(rc > 1 || cc > 1)
                    {
                        this.Invoke(__writeLog, "結合セル:" + rg.Cells().First().Address.ToString() + "を処理しています...");
                        //rg.Style.Fill.BackgroundColor = XLColor.FromArgb(0xFDCA5F);
                        if (rc > 1) comm += "縦に" + rc.ToString() + "行 ";
                        if (cc > 1) comm += "横に" + cc.ToString() + "列 ";
                        comm += "結合されています。";
                        rg.Cells().First().Comment.AddText(comm);
                        cnt++;
                        
                    }

                }

                if(cnt == 0)
                {
                    this.Invoke(__writeLog, "結合セルはありません。処理をキャンセルします。");
                    return;
                }

                this.Invoke(__writeLog, "結合セルの強調処理が完了しました....");

                //保存の段取り
                string new_filename = Path.GetFileNameWithoutExtension(currentWbPath);
                string ext = Path.GetExtension(currentWbPath);
                string new_savepath = getCurrentFileWorkPath() + new_filename + "_結合セル強調_" + fetch_filename_logtime() + ext;

                //別名保存
                currentWb.SaveAs(new_savepath);
                this.Invoke(__writeLog, "処理が完了しました。出力ファイル：" + new_savepath);
                
            });
        }

        //Excelファイルに出力
        private void saveNewBookAs(List<List<string>> data, string filename)
        {
            _writeLog __writeLog = writeLog;

            try
            {
                using (var wb = new ClosedXML.Excel.XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Sheet1");

                    //行のループ
                    for (int i = 0; i < data.Count; i++)
                    {
                        List<string> row = (List<string>)data[i];

                        //列のループ
                        for (int j = 0; j < row.Count; j++)
                        {
                            string col = (string)row[j];
                            ws.Cell(i + 1, j + 1).Value = col;
                            ws.Cell(i + 1, j + 1).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            ws.Cell(i + 1, j + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            ws.Cell(i + 1, j + 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            ws.Cell(i + 1, j + 1).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                            ws.Cell(i + 1, j + 1).Style.Font.FontName = "ＭＳ Ｐゴシック";
                            ws.Cell(i + 1, j + 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);

                        }

                    }

                    wb.SaveAs(filename);
                    main_form.Invoke(__writeLog, "保存に成功しました。（" + filename + "）");
                }
            }
            catch (Exception ex)
            {
                main_form.Invoke(__writeLog, "【エラー】" + ex.Message);
                return;
            }
        }

        //Bookのフッターをクリアする
        private void clearBookFootersWrap()
        {
            _writeLog __writeLog = writeLog;
            List<string> books = getBookList();
            foreach (string book in books)
            {
                main_form.Invoke(__writeLog, book + " を処理しています...");
                FooterClear(book);
            }
            var excels = Process.GetProcessesByName("EXCEL");
            foreach (var x in excels)
            {
                x.Kill();
            }
            this.Invoke(__writeLog, "処理が完了しました");
        }

        //BookをPDF保存する
        private void savePDFFromBook()
        {
            _writeLog __writeLog = writeLog;
            List<string> books = getBookList();
            foreach(string book in books)
            {
                main_form.Invoke(__writeLog, book + " を処理しています...");
                string fileName = Path.GetFileName(book);
                string fileNamePdf = fileName.Substring(0, fileName.IndexOf(".")) + ".pdf";
                string save_filename = currentFolderDir + @"\" + fileNamePdf;
                SaveAsPdf(book, save_filename);
            }
            var excels = Process.GetProcessesByName("EXCEL");
            foreach(var x in excels)
            {
                x.Kill();
            }
            this.Invoke(__writeLog, "処理が完了しました");
        }

    }
}
