using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EffectiveEx
{
    partial class Form1
    {

        //sheetNameBombo.Textのデリゲート
        public delegate string _sheetNameComboVal();
        public string sheetNameComboVal()
        {
            return sheetNameCombo.Text;
        }

        //searchValues.Textのデリゲート
        public delegate string _searchValuesVal();
        public string searchValuesVal()
        {
            return searchValues.Text;
        }

        //columnValues.Valueのデリゲート
        public delegate int _columnValuesVal();
        public int columnValuesVal()
        {
            return (int)columnValues.Value;
        }

        //skipRowNumbers.Valueのデリゲート
        public delegate int _skipRowNumbersVal();
        public int skipRowNumbersVal()
        {
            return (int)skipRowNumbers.Value;
        }

        //withoutConditionColNum.Valueのデリゲート
        public delegate int _withoutConditionColNumVal();
        public int withoutConditionColNumVal()
        {
            return (int)withoutConditionColNum.Value;
        }

        //searchResultWithoutCheck.Checkedのデリゲート
        public delegate Boolean _searchResultWithoutCheckIsChecked();
        public Boolean searchResultWithoutCheckIsChecked()
        {
            return searchValResultWithoutCheck.Checked;
        }

        //reportText.AppendTextのデリゲート
        public delegate void _reportTextAppendTextThis(string key, int val);
        public void reportTextAppendTextThis(string key, int val)
        {
            reportText.AppendText(key + "\t" + val + "\r\n");
        }

    }
}
