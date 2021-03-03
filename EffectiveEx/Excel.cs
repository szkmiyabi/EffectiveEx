using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;

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

        //アンケートデータ自動抽出
        private async Task getAnkResultRecordWrap()
        {
            await Task.Run(() =>
            {
                _writeLog __writeLog = writeLog;
                _sheetNameComboVal __sheetNameComboVal = sheetNameComboVal;

                List<List<string>> data = new List<List<string>>();
                string str = "";
                List<string> head_row = new List<string>()
                {
                    "社名・団体名",
                    "Q1-1業種",
                    "Q1-2従業員数",
                    "Q2-1経営への影響",
                    "Q2-2具体的な影響の内容（複数回答）",
                    "Q3-1従業員への衛星管理（複）",
                    "Q3-2来客者への衛生管理（複）",
                    "Q3-3設備・屋内への消毒の実施（複）",
                    "Q3-4換気の実施（複数回答）",
                    "Q3-5その他",
                    "Q4-1ウエブ会議の活用",
                    "Q4-2テレワークの導入等",
                    "Q4-3従業員の配置等",
                    "Q4-4その他",
                    "Q5-1コスト削減（複数回答）",
                    "Q5-2外部への営業活動の強化",
                    "Q5-3生産の設備投資の実施",
                    "Q5-4ＩＴの設備投資の実施",
                    "Q5-5資金調達力の強化",
                    "Q5-6BCP（事業継続計画）の策定・見直し",
                    "Q5-7事業を改善・工夫したこと（自由回答）",
                    "Q6新規ビジネスの実施",
                    "Q8取材可否"
                };

                data.Add(head_row);

                foreach (var sh in currentWb.Worksheets)
                {
                    string shName = sh.Name;
                    if (shName == "入力方法") continue;
                    if (shName == "集計用") continue;

                    initCurrentWorksheet(shName);
                    this.Invoke(__writeLog, "シート [ " + shName + " ] からデータを抽出しています...");

                    List<string> row = getAnkResultRecord();
                    data.Add(row);
                }

                this.Invoke(__writeLog, "抽出完了。抽出データの保存を開始します....");

                //保存の段取り
                string new_filename = Path.GetFileNameWithoutExtension(currentWbPath);
                string ext = Path.GetExtension(currentWbPath);
                string new_savepath = getCurrentFileWorkPath() + "ANK抽出データ_【" + new_filename + "】_" + fetch_filename_logtime() + ext;

                //保存
                saveNewBookAs(data, new_savepath);

                this.Invoke(__writeLog, "処理が完了しました。");

                /*
                foreach (var r in data)
                {
                    foreach (string line in r)
                    {
                        str += line + "\t";
                    }
                    str += "\r\n";
                }

                this.Invoke(__writeLog, str);
                */
            });

        }

        //1件分のレコードデータ抽出
        private List<string> getAnkResultRecord()
        {
            List<string> data = new List<string>();

            string body = "";

            //誤字修正
            currentWs.Range("P2").Cells().ElementAt<IXLCell>(0).Value = "サービス業";
            currentWs.Range("H26").Cells().ElementAt<IXLCell>(0).Value = "コロナで導入";
            currentWs.Range("H28").Cells().ElementAt<IXLCell>(0).Value = "コロナで導入";
            currentWs.Range("G64").Cells().ElementAt<IXLCell>(0).Value = "取材を受けることは難しい";

            //社名・団体名
            string q0Addr = "$C$2";
            string a0 = querySingleResultBrowseVal(q0Addr);
            data.Add(a0);

            //Q1-1業種
            string q1x1Addr = "$F$2:$R$3";
            string a1x1 = queryMultiResult(q1x1Addr);
            data.Add(a1x1);

            //Q1-2従業員数
            string q1x2Addr = "$F$4:$L$5";
            string a1x2 = queryMultiResult(q1x2Addr);
            data.Add(a1x2);

            //Q2-1経営への影響
            string q2x1Addr = "$F$8:$K$9";
            string a2x1 = queryMultiResult(q2x1Addr);
            data.Add(a2x1);

            //Q2-2具体的な影響の内容（複数回答）
            string q2x2Addr = "$F$10:$O$11";
            string a2x2 = queryMultiResult(q2x2Addr);
            data.Add(a2x2);


            //Q3-1従業員への衛星管理（複）
            string q3x1Addr = "$F$14:$I$15";
            string a3x1 = queryMultiResult(q3x1Addr);
            data.Add(a3x1);

            //Q3-2来客者への衛生管理（複）
            string q3x2Addr = "$F$16:$J$17";
            string a3x2 = queryMultiResult(q3x2Addr);
            data.Add(a3x2);

            //Q3-3設備・屋内への消毒の実施（複）
            string q3x3Addr = "$F$18:$I$19";
            string a3x3 = queryMultiResult(q3x3Addr);
            data.Add(a3x3);

            //Q3-4換気の実施（複数回答）
            string q3x4Addr = "$F$20:$I$21";
            string a3x4 = queryMultiResult(q3x4Addr);
            data.Add(a3x4);

            //Q3-5その他  空欄＞記載なし
            string q3x5Addr = "$F$23";
            string a3x5 = querySingleResult(q3x5Addr);
            if(a3x5 == "記載あり")
            {
                string a3x5br = querySingleResultBrowseVal(q3x5Addr);
                data.Add(a3x5br);
            }
            else
            {
                data.Add(a3x5);
            }

            //Q4-1ウエブ会議の活用
            string q4x1Addr = "$F$26:$J$27";
            string a4x1 = queryMultiResult(q4x1Addr);
            data.Add(a4x1);

            //Q4-2テレワークの導入等
            string q4x2Addr = "$F$28:$J$29";
            string a4x2 = queryMultiResult(q4x2Addr);
            data.Add(a4x2);

            //Q4-3従業員の配置等
            string q4x3Addr = "$F$30:$J$31";
            string a4x3 = queryMultiResult(q4x3Addr);
            data.Add(a4x3);

            //Q4-4その他  空欄＞記載なし
            string q4x4Addr = "$F$33";
            string a4x4 = querySingleResult(q4x4Addr);
            if(a4x4 == "記載あり")
            {
                string a4x4br = querySingleResultBrowseVal(q4x4Addr);
                data.Add(a4x4br);
            }
            else
            {
                data.Add(a4x4);
            }

            //Q5-1コスト削減（複数回答）
            string q5x1Addr = "$E$36:$I$37";
            string a5x1 = queryMultiResult(q5x1Addr);
            data.Add(a5x1);

            //Q5-2外部への営業活動の強化
            string q5x2Addr = "$F$38:$I$39";
            string a5x2 = queryMultiResult(q5x2Addr);
            data.Add(a5x2);

            //Q5-3生産の設備投資の実施
            string q5x3Addr = "$F$40:$I$41";
            string a5x3 = queryMultiResult(q5x3Addr);
            data.Add(a5x3);

            //Q5-4ＩＴの設備投資の実施
            string q5x4Addr = "$F$42:$I$43";
            string a5x4 = queryMultiResult(q5x4Addr);
            data.Add(a5x4);

            //Q5-5資金調達力の強化
            string q5x5Addr = "$F$44:$I$45";
            string a5x5 = queryMultiResult(q5x5Addr);
            data.Add(a5x5);

            //Q5-6BCP（事業継続計画）の策定・見直し
            string q5x6Addr = "$F$46:$I$47";
            string a5x6 = queryMultiResult(q5x6Addr);
            data.Add(a5x6);

            //Q5-7事業を改善・工夫したこと（自由回答）
            string q5x7Addr = "$F$49";
            string a5x7 = querySingleResult(q5x7Addr);
            if(a5x7 == "記載あり")
            {
                string a5x7br = querySingleResultBrowseVal(q5x7Addr);
                data.Add(a5x7br);
            }
            else
            {
                data.Add(a5x7);
            }

            //Q6新規ビジネスの実施
            string q6Addr = "$F$52:$I$53";
            string a6 = queryMultiResult(q6Addr);
            data.Add(a6);

            //Q8取材可否
            string q8Addr = "$F$64:$H$65";
            string a8 = queryMultiResult(q8Addr);
            data.Add(a8);
            return data;

        }

        //単独セルの結果を取得
        private string querySingleResult(string addr)
        {
            string ret = "";
            IXLRange rg = currentWs.Range(addr);

            double val = 0;
            try
            {
                val = rg.Cells().ElementAt<IXLCell>(0).GetDouble();
            }
            catch(Exception ex) { }

            if (val == 1) ret = "記載あり";
            else ret = "記載なし";
            return ret;
        }
        private string querySingleResultBrowseVal(string addr)
        {
            string ret = "";
            string srcAddr = "";

            //社名・団体名
            string q0Addr = "$C$2";

            //Q3-5その他  空欄＞記載なし
            string q3x5Addr = "$F$23";

            //Q4-4その他  空欄＞記載なし
            string q4x4Addr = "$F$33";

            //Q5-7事業を改善・工夫したこと（自由回答）
            string q5x7Addr = "$F$49";

            //Q6新規ビジネスの実施
            string q6x2Addr = "E54";

            if (addr == q0Addr) srcAddr = q0Addr;
            else if (addr == q3x5Addr) srcAddr = "E22";
            else if (addr == q4x4Addr) srcAddr = "F32";
            else if (addr == q5x7Addr) srcAddr = "F48";
            else if (addr == q6x2Addr) srcAddr = "E54";

            IXLRange rg = currentWs.Range(srcAddr);
            try
            {
                ret = (string)rg.Cells().ElementAt<IXLCell>(0).Value;
            }
            catch(Exception ex) { }
            return ret;
        }

        //複数セルの複数結果を取得
        private string queryMultiResult(string addr)
        {
            string ret = "";
            IXLRange rg = currentWs.Range(addr);
            IXLCell lastCell = rg.LastCell();
            int rx = lastCell.WorksheetRow().RowNumber();
            int cx = lastCell.WorksheetColumn().ColumnNumber();

            for (int i=6; i<= cx; i++)
            {
                double val = 0;
                try
                {
                    val = currentWs.Cell(rx, i).GetDouble();
                }
                catch(Exception ex) { }
                
                if (val == 1)
                {
                    var inqval = currentWs.Cell(rx-1, i).Value;
                    ret += (string)inqval + ",";
                }
            }

            try
            {
                Regex pt = new Regex(@",$");
                ret = pt.Replace(ret, "");
            }
            catch(Exception ex) { }

            return ret;
        }

        //カンマ区切りのセルデータを点数付け、無回答、記載なしは0とする
        private async Task cvEncodePointWrap()
        {
            await Task.Run(() =>
            {
                //共通デリゲートインスタンス
                _sheetNameComboVal __sheetNameComboVal = sheetNameComboVal;
                _writeLog __writeLog = writeLog;
                _skipRowNumbersVal __skipRowNumbersVal = skipRowNumbersVal;
                _columnValuesVal __columnValuesVal = columnValuesVal;


                //コンボ選択値をアクティブシートにする
                initCurrentWorksheet((string)this.Invoke(__sheetNameComboVal));

                int r = getStartRow();
                int rx = getEndRow();
                int cx = getEndCol();

                int skip_nm = (int)this.Invoke(__skipRowNumbersVal);
                r = r + skip_nm;
                int icx = (int)this.Invoke(__columnValuesVal);

                for (int i = r; i <= rx; i++)
                {
                    this.Invoke(__writeLog, i + "行目の処理....");
                    for (int j = 1; j <= cx; j++)
                    {
                        if(j >= icx)
                        {
                            int cnt = cvEncodePoint(currentWs.Cell(i, j));
                            currentWs.Cell(i, j).Value = cnt;
                            if (cnt == 0) currentWs.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(0xFFB3B3);

                        }
                    }
                }


                //保存の段取り
                string new_filename = Path.GetFileNameWithoutExtension(currentWbPath);
                string ext = Path.GetExtension(currentWbPath);
                string new_savepath = getCurrentFileWorkPath() + new_filename + "_点数化_" + fetch_filename_logtime() + ext;

                //別名保存
                currentWb.SaveAs(new_savepath);
                this.Invoke(__writeLog, "処理が完了しました。出力ファイル：" + new_savepath);

            });
        }

        //
        private async Task cvColorWrap()
        {
            await Task.Run(() =>
            {
                //共通デリゲートインスタンス
                _sheetNameComboVal __sheetNameComboVal = sheetNameComboVal;
                _writeLog __writeLog = writeLog;
                _skipRowNumbersVal __skipRowNumbersVal = skipRowNumbersVal;

                //コンボ選択値をアクティブシートにする
                initCurrentWorksheet((string)this.Invoke(__sheetNameComboVal));

                int r = getStartRow();
                int rx = getEndRow();
                int cx = getEndCol();

                int skip_nm = (int)this.Invoke(__skipRowNumbersVal);
                r = r + skip_nm;

                for (int i = r; i <= rx; i++)
                {
                    this.Invoke(__writeLog, i + "行目の処理....");
                    for (int j = 1; j <= cx; j++)
                    {
                        if (j >= 5)
                        {
                            int cnt = cvEncodePoint(currentWs.Cell(i, j));
                            if(cnt == 0) currentWs.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(0xFFB3B3);
                            //currentWs.Cell(i, j).Value = cnt;
                        }
                    }
                }


                //保存の段取り
                //string new_filename = Path.GetFileNameWithoutExtension(currentWbPath);
                //string ext = Path.GetExtension(currentWbPath);
                //string new_savepath = getCurrentFileWorkPath() + new_filename + "_点数化_" + fetch_filename_logtime() + ext;

                //保存
                currentWb.Save();
                this.Invoke(__writeLog, "処理が完了しました。");

            });
        }

        private int cvEncodePoint(IXLCell cl)
        {
            int pt = 0;
            string vl = "";
            try
            {
                vl = (string)cl.Value;
                char[] delimit = { ',' };
                string[] cvl = vl.Split(delimit);
                pt = cvl.Length;
            }
            catch(Exception ex) { }

            if(vl == "無回答" || vl == "記載なし" || vl == "" || vl == "計画なし" || vl == "変更なし")
            {
                pt = 0;
            }
            return pt;
        }

        private List<string> effectArr = new List<string>()
        {
            "売上げ（販売や受注）の減少",
            "営業活動や商談の困難・延期",
            "仕入れの困難・遅延",
            "資金繰りに関する不安・逼迫化",
            "勤務時間や休日の変更等",
            "感染予防対策への手間・コスト増大",
            "ウエブ会議導入等",
            "採用活動の困難性",
            "その他",
            "無回答"
        };

        private Boolean inListVal(string str)
        {
            Boolean flg = false;
            foreach(string key in effectArr)
            {
                if(key == str)
                {
                    flg = true;
                    break;
                }
            }
            return flg;
        }

        private int cvCheckPoint(IXLCell cl)
        {
            int pt = 0;
            string vl = "";
            try
            {
                vl = (string)cl.Value;
                char[] delimit = { ',' };
                string[] cvl = vl.Split(delimit);
                for(int i=0; i<cvl.Length; i++)
                {
                    string row = cvl[i];
                    if (inListVal(row))
                    {
                        pt = 1;
                        return pt;
                    }
                }
            }
            catch (Exception ex) { }
            return pt;
        }

    }
}
