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

        //フォルダパスを取得
        private string getFolderPath()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Bookの保存先フォルダ選択";
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            string path = "";
            if (fbd.ShowDialog(this) == DialogResult.OK)
                path = fbd.SelectedPath;
            return path;
        }

        //Bookファイル名の一覧リストを取得（指定ディレクトリから）
        private List<string> getBookList()
        {
            string path = getFolderPath();
            List<string> arr = new List<string>();
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(path);
            System.IO.FileInfo[] files = di.GetFiles("*.xlsx", System.IO.SearchOption.AllDirectories);
            foreach (System.IO.FileInfo f in files)
            {
                if((f.Attributes & System.IO.FileAttributes.Hidden)!= System.IO.FileAttributes.Hidden)
                    arr.Add(f.FullName);
            }
            return arr;
        }
    }
}
