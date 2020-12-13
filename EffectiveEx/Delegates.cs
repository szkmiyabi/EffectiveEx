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

    }
}
