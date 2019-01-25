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
        public float MarginLeft { get; set; }
        public float MarginRight { get; set; }
        public float MarginTop { get; set; }
        public float MarginBottom { get; set; }
        public string Color { get; set; }
        public bool Bold { get; set; }
    }
}
