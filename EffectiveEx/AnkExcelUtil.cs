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
            if (a3x5 == "記載あり")
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
            if (a4x4 == "記載あり")
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
            if (a5x7 == "記載あり")
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
            catch (Exception ex) { }

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
            catch (Exception ex) { }
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

            for (int i = 6; i <= cx; i++)
            {
                double val = 0;
                try
                {
                    val = currentWs.Cell(rx, i).GetDouble();
                }
                catch (Exception ex) { }

                if (val == 1)
                {
                    var inqval = currentWs.Cell(rx - 1, i).Value;
                    ret += (string)inqval + ",";
                }
            }

            try
            {
                Regex pt = new Regex(@",$");
                ret = pt.Replace(ret, "");
            }
            catch (Exception ex) { }

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
                        if (j >= icx)
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
            catch (Exception ex) { }

            if (vl == "無回答" || vl == "記載なし" || vl == "" || vl == "計画なし" || vl == "変更なし")
            {
                pt = 0;
            }
            return pt;
        }

        //カンマ区切りデータセルの項目有無判定（0＝なし/1＝あり）を列展開する
        private async Task cvQueryOutputColumn()
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

                int targetCol = (int)this.Invoke(__columnValuesVal);

                int cz = cx + effectArr.Count;
                int yz = 0;

                for (int z = cx + 1; z <= cz; z++)
                {
                    currentWs.Cell(1, z).Value = (string)effectArr[yz];
                    yz++;
                }

                int skip_nm = (int)this.Invoke(__skipRowNumbersVal);
                r = r + skip_nm;
                int icx = (int)this.Invoke(__columnValuesVal);

                for (int i = r; i <= rx; i++)
                {
                    this.Invoke(__writeLog, i + "行目の処理....");
                    //string cv = (string)currentWs.Cell(i, 2).Value;

                    List<int> arr = createCvCheckCountArr(currentWs.Cell(i, targetCol));
                    int y = 0;

                    for (int j = cx + 1; j <= cz; j++)
                    {
                        currentWs.Cell(i, j).Value = arr[y];
                        if (arr[y] == 1) currentWs.Cell(i, j).Style.Fill.BackgroundColor = XLColor.FromArgb(0xFFB3B3);
                        y++;
                    }
                }

                //保存の段取り
                string new_filename = Path.GetFileNameWithoutExtension(currentWbPath);
                string ext = Path.GetExtension(currentWbPath);
                string new_savepath = getCurrentFileWorkPath() + new_filename + "_個別カウント_" + fetch_filename_logtime() + ext;

                //別名保存
                currentWb.SaveAs(new_savepath);
                this.Invoke(__writeLog, "処理が完了しました。出力ファイル：" + new_savepath);

            });

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

        private int getHitColumnNum(string str)
        {
            int n = 0;
            foreach (string key in effectArr)
            {
                if (key == str)
                {
                    return n;
                }
                n++;
            }
            return n;
        }

        private List<int> createCvCheckCountArr(IXLCell cl)
        {
            string vl = "";
            List<int> arr = new List<int>()
            {
                0,0,0,0,0,0,0,0,0,0
            };

            try
            {
                vl = (string)cl.Value;
                char[] delimit = { ',' };
                string[] cvl = vl.Split(delimit);
                for (int i = 0; i < cvl.Length; i++)
                {
                    string row = cvl[i];
                    int idx = getHitColumnNum(row);
                    arr[idx] = 1;
                }
            }
            catch (Exception ex) { }
            return arr;

        }
    }
}
