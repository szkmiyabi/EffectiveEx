﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace EffectiveEx
{
    public static class ComRelease
    {
        public static void FinalReleaseComObjects(params object[] objects)
        {
            foreach(object o in objects)
            {
                try
                {
                    if (o == null)
                        continue;
                    if (Marshal.IsComObject(o) == false)
                        continue;
                    Marshal.FinalReleaseComObject(o);
                }
                catch (Exception) { }
            }
        }
    }
}
