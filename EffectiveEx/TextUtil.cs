using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EffectiveEx
{
    partial class Form1
    {

        //検索値配列を取得
        private List<string> getSearchValues()
        {
            //デリゲートインスタンス
            _searchValuesVal __searchValuesVal = searchValuesVal;

            List<string> data = new List<string>();
            //非同期処理のためデリゲート経由で値を拾う
            var src = (string)this.Invoke(__searchValuesVal);
            string[] del = { "\r\n" };
            string[] lines = src.Split(del, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                data.Add(lines[i].ToString());
            }
            return data;
        }

        //現在の時刻からファイル名を生成
        public string fetch_filename_from_datetime(string extension)
        {
            DateTime ld = DateTime.Now;
            return ld.ToString("yyyyMMdd_HHmmss") + "." + extension;
        }

        //ファイル用時刻文字列を生成
        public string fetch_filename_logtime()
        {
            DateTime ld = DateTime.Now;
            return ld.ToString("yyyyMMdd_HHmmss");
        }

        //ログ出力の時刻文字列を生成
        public string get_logtime()
        {
            DateTime ld = DateTime.Now;
            return ld.ToString("yyyy-MM-dd HH:mm:ss");
        }

        //レポート出力
        public delegate void _writeLog(string str);
        public void writeLog(string str)
        {
            reportText.AppendText(get_logtime() + " " + str + "\r\n");
        }

    }
}
