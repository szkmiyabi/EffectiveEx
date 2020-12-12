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

    }
}
