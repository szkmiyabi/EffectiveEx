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
    public partial class DataGridForm : Form
    {

        private Form1 main_form;

        //コンストラクタ
        public DataGridForm()
        {
            InitializeComponent();
            gridTableView.RowHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            gridTableView.AllowUserToAddRows = false;
            gridTableView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //行表示のイベントハンドラ登録
            gridTableView.RowPostPaint += new DataGridViewRowPostPaintEventHandler(RowPostPaint);
            //メインフォームと紐付け
            main_form = Form1.main_form;
        }

        //行番号表示
        private void RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //行ヘッダを描画領域にセットする（右端に4ドット隙間を空ける）
            Rectangle rect = new Rectangle(
                e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                gridTableView.RowHeadersWidth - 4,
                e.RowBounds.Height
            );
            //行番号を描画領域に表示する
            TextRenderer.DrawText(
                e.Graphics,
                (e.RowIndex + 1).ToString(),
                gridTableView.RowHeadersDefaultCellStyle.Font,
                rect,
                gridTableView.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right
            );
        }

        //データグリッドセットアップ
        public void init(DataTable tbl, string title)
        {
            this.Text = title;
            gridTableView.DataSource = tbl;
        }

        
    }
}
