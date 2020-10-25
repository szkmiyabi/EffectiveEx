using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EffectiveEx
{
    partial class Form1
    {
        //カレントWorkbookをセット
        private void initCurrentBook()
        {
            currentWb = new XLWorkbook(currentWbPath);
        }

        //カレントWorkbookのカレントSheetをセット
        private void initCurrentWorksheet(int idx)
        {
            currentWs = currentWb.Worksheet(idx);
        }

        //カレントWorkbookのカレントSheetをセット（オーバーライド）
        private void initCurrentWorksheet(string sname)
        {
            currentWs = currentWb.Worksheets.Worksheet(sname);
        }

        //対象シートコンボのセットアップ
        private void initTargetWorksheetCombo()
        {
            initCurrentBook();
            sheetNameCombo.Items.Clear();

            var sheets = currentWb.Worksheets;
            foreach (IXLWorksheet st in sheets)
            {
                string sname = st.Name;
                sheetNameCombo.Items.Add(sname);
            }

            sheetNameCombo.Enabled = true;
        }
    }
}
