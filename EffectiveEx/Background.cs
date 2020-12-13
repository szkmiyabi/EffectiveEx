using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EffectiveEx
{
    partial class Form1
    {

        //imageリソースを取得
        public Bitmap getImageFromResource(string imgname)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream stream = assembly.GetManifestResourceStream("EffectiveEx.resources." + imgname);
            Bitmap bmp = new Bitmap(stream);
            return bmp;
        }

        //iconリソースを取得
        public Icon getIconFromResource(string iconame)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream stream = assembly.GetManifestResourceStream("EffectiveEx.resources." + iconame);
            Icon ico = new Icon(stream);
            return ico;
        }

    }
}
