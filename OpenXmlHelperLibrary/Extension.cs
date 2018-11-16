using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlHelperLibrary
{
 public   static class Extension
    {
        public static JustificationValues GetJustificationByString(this string justificationVol, JustificationValues @default = JustificationValues.Both)
        {
            switch (justificationVol.ToLower())
            {
                case "center": return JustificationValues.Center;
                case "left": return JustificationValues.Left;
                case "right": return JustificationValues.Right;
            }
            return @default;
        }
    }
}
