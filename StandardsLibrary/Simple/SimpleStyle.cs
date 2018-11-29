using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StandardsLibrary.Simple
{
 public class SimpleStyle
    {
        public string Font { get; set; }
        public int FontSize { get; set; }
        public float LineSpacing { get; set; }
        public float BeforeSpacing { get; set; }
        public float AfterSpacing { get; set; }
        public float FirstLineIndentation { get; set; }
        public float LeftIndentation { get; set; }
        public float RightIndentation { get; set; }
        public string Alignment { get; set; }
        public int MarginLeft { get; set; }
        public int MarginRight { get; set; }
        public int MarginTop { get; set; }
        public int MarginBottom { get; set; }
        public string Color { get; set; }
        public bool Bold { get; set; }
    }
}
