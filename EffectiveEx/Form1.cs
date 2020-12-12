﻿using ClosedXML.Excel;
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
        string saveDirPath;
        XLWorkbook currentWb;
        IXLWorksheet currentWs;

        //コンストラクタ
        public Form1()
        {
            InitializeComponent();
            currentWbPath = "";
            currentWb = null;
            currentWs = null;

            sheetNameCombo.Enabled = false;
            columnValues.Text = "1";
            skipRowNumbers.Text = "1";
            withoutConditionColNum.Text = "1";
            saveDirPath = getUserHomePath() + @"\Desktop\";
            outputDirectory.Text = saveDirPath;
        }

        //出力先参照ボタンクリック
        private void browseOutputDirectoryButton_Click(object sender, EventArgs e)
        {
            setOutputPath();
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

        //行削除
        private void deleteRowButton_Click(object sender, EventArgs e)
        {
            deleteRowByCondition();
        }

        //値集計検索
        private void searchValsResultButton_Click(object sender, EventArgs e)
        {
            searchValsResult();
        }

        //クリア
        private void reportClearButton_Click(object sender, EventArgs e)
        {
            reportText.Clear();
        }

        //LPRフォーマット
        private async void lpReportFormatButton_Click(object sender, EventArgs e)
        {
            await lpReportFormat();
            //await tableBorderedAsync();
        }
    }
}