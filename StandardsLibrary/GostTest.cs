using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StandardsLibrary
{
    public class GostTest : Standards
    {

        public string Font { get; set; } = "Times New Roman";
        public int FontSize { get; set; } = 30;
        public string Alignment { get; set; } = "Both";
        public string Alignment_Image { get; set; } = "Center";

        public float LineSpacing { get; set; } = 1.5f;
        public int MarginLeft { get; set; } = 0;
        public int MarginRight { get; set; } = 0;
        public int MarginTop { get; set; } = 0;
        public int MarginBottom { get; set; } = 0;
        public string HeaderColor { get; set; } = "365F91";
        public float BeforeSpacing { get; set; } = 1;
        public float AfterSpacing { get; set; } = 1;
        public float FirstLineIndentation { get; set; } = 1.5f;
        public float LeftIndentation { get; set; } = 0;
        public float RightIndentation { get; set; } = 0;
      
        public bool Bold { get; set; } = true;

      

        public GostTest()
        {

        }
        public override string GetFont()
        {
            return Font;
        }

        public override int GetFontSize()
        {
            return FontSize;
        }

        public override float GetLineSpacing()
        {
            return LineSpacing;
        }

        public override string GetAlignment() => Alignment;
        public override string GetAlignment_Image()
        {
            return Alignment_Image;
        }

        public override float GetBeforeSpacing()
        {
            return BeforeSpacing;
        }

        public override float GetAfterSpacing()
        {
            return AfterSpacing;
        }

        public override float GetFirstLineIndentation()
        {
            return FirstLineIndentation;
        }

        public override float GetLeftIndentation()
        {
            return LeftIndentation;
        }

        public override float GetRightIndentation()
        {
            return RightIndentation;
        }

        public override int GetMarginLeft()
        {
            return MarginLeft;
        }

        public override int GetMarginRight()
        {
            return MarginRight;
        }

        public override int GetMarginTop()
        {
            return MarginTop;
        }

        public override int GetMarginBottom()
        {
            return MarginBottom;
        }

        public override bool isBold()
        {
            return Bold;
        }

        public override string GetHeaderColor()
        {
            return HeaderColor;
        }
    }
}
