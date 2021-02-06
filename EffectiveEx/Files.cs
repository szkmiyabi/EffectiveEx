using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EffectiveEx
{
    partial class Form1
    {
        //読み込みファイルパス取得
        private string getLoadPath(string extension)
        {
            string path = "";
            OpenFileDialog ofd = new OpenFileDialog();
            switch (extension)
            {
                case "xlsx":
                    ofd.Filter = "Excelワークブック(*.xlsx)|*.xlsx";
                    break;
                default:
                    ofd.Filter = "すべてのファイル(*.*)|*.*";
                    break;
            }
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                path = ofd.FileName;
            }
            return path;
        }

        //カレントファイルのパスを取得
        private string getCurrentFileWorkPath()
        {
            string retPath = "";
            if(currentWbPath == null)
            {
                retPath = getUserHomePath() + @"\Desktop\";
            }
            else
            {
                retPath = Path.GetDirectoryName(currentWbPath) + @"\";
            }
            return retPath;
        }

        //ユーザのホームフォルダパス取得
        private string getUserHomePath()
        {
            return System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }
    }
}
