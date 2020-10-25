using System;
using System.Collections.Generic;
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

        //保存ファイルパス取得
        private string getSavePath(string extension)
        {
            string path = "";
            SaveFileDialog sfd = new SaveFileDialog();
            switch (extension)
            {
                case "xlsx":
                    sfd.Filter = "Excelワークブック(*.xlsx)|*.xlsx";
                    break;
                default:
                    sfd.Filter = "すべてのファイル(*.*)|*.*";
                    break;
            }
            if (!saveDirPath.Equals("")) sfd.InitialDirectory = saveDirPath;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                path = sfd.FileName;
            }
            return path;
        }

        //出力フォルダパスを選択
        private string getSaveFolderName()
        {
            string dirname = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "保存先のフォルダを選択";
            fbd.SelectedPath = getUserHomePath() + @"\Desktop";
            fbd.ShowNewFolderButton = true;
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                dirname = fbd.SelectedPath + @"\";
            }
            return dirname;
        }

        //ユーザのホームフォルダパス取得
        private string getUserHomePath()
        {
            return System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }
    }
}
