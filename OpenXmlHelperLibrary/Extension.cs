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
                case "both": return JustificationValues.Both;
            }
            return @default;
        }

        public static UnderlineValues GetUnderlineByString(this string underlineVol, UnderlineValues @default = UnderlineValues.None)
        {
            switch (underlineVol.ToLower())
            {
                case "single": return UnderlineValues.Single;
                case "double": return UnderlineValues.Double;
                case "dash": return UnderlineValues.Dash;
                case "thick": return UnderlineValues.Thick;
                case "wave": return UnderlineValues.Wave;
                case "none": return UnderlineValues.None;
            }
            return @default;
        }
    }
}
