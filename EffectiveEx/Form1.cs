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

            saveDirPath = getUserHomePath() + @"\Desktop";
            outputDirectory.Text = saveDirPath;
        }
    }
}
