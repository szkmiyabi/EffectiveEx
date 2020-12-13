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
        }

        //データグリッドセットアップ
        public void init(DataTable tbl, string title)
        {
            this.Text = title;
            gridTableView.DataSource = tbl;
        }
    }
}
