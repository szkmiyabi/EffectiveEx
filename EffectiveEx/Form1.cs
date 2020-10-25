using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EffectiveEx
{
    public partial class Form1 : Form
    {
        string currentWbPath;
        string outputWbPath;
        string saveDirPath;
        XLWorkbook currentWb;
        XLWorkbook outputWb;
        IXLWorksheet currentWs;

        public Form1()
        {
            InitializeComponent();
            currentWbPath = "";
            outputWbPath = "";
            currentWb = null;
            outputWb = null;
            currentWs = null;

            sheetNameCombo.Enabled = false;

            saveDirPath = getUserHomePath() + @"\Desktop";
            outputDirectory.Text = saveDirPath;
        }

        //出力先参照ボタンクリック
        private void browseOutputDirectoryButton_Click(object sender, EventArgs e)
        {
            string dir = getSaveFolderName();
            if (dir.Equals("")) return;
            saveDirPath = dir;
            outputDirectory.Text = dir;
        }

        //開くボタンクリック
        private void openButton_Click(object sender, EventArgs e)
        {
            statusText.Text = "";
            currentWbPath = "";
            currentWb = null;
            string path = getLoadPath("xlsx");
            if (path.Equals("")) return;
            currentWbPath = path;
            statusText.Text = path;
            initTargetWorksheetCombo();
        }
    }
}
