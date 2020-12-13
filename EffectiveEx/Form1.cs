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
        string saveDirPath;
        XLWorkbook currentWb;
        IXLWorksheet currentWs;

        //Form1インスタンス
        private static Form1 _main_form;

        //DataGridFormインスタンス
        private static DataGridForm _data_grid_form;

        //コンストラクタ
        public Form1()
        {
            InitializeComponent();
            currentWbPath = "";
            currentWb = null;
            currentWs = null;

            sheetNameCombo.Enabled = false;
            saveDirPath = getUserHomePath() + @"\Desktop\";
            outputDirectory.Text = saveDirPath;

            //静的プロパティに自身を代入
            _main_form = this;
        }

        //Form1のゲッターとセッター
        public static Form1 main_form
        {
            get { return _main_form; }
            set { _main_form = value; }
        }

        //DataGridFormのゲッター
        public static DataGridForm data_grid_form
        {
            get
            {
                if(_data_grid_form == null || _data_grid_form.IsDisposed)
                {
                    _data_grid_form = new DataGridForm();
                }
                return _data_grid_form;
            }
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
        private async void deleteRowButton_Click(object sender, EventArgs e)
        {
            await deleteRowByCondition();
        }

        //値集計検索
        private async void searchValsResultButton_Click(object sender, EventArgs e)
        {
            await searchValsResult();
        }

        //クリア
        private void reportClearButton_Click(object sender, EventArgs e)
        {
            reportText.Clear();
        }

        //LPR検査結果表整形
        private async void lpReportFormatButton_Click(object sender, EventArgs e)
        {
            await lpReportFormat();
        }

        //プレビュー
        private async void worksheetPreviewButton_Click(object sender, EventArgs e)
        {
            await previewWorksheetWrap();
        }
    }
}
