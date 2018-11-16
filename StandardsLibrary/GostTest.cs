using System;

namespace StandardsLibrary
{
    public class GostTest : Standards
    {

        public string Font { get; set; } = "Times New Roman";
        public int FontSize { get; set; } = 30;
        public float LineSpacing { get; set; }
        public string Alignment { get; set; }
        public int MarginLeft { get; set; }
        public int MarginRight { get; set; }
        public int MarginTop { get; set; }
        public int MarginBottom { get; set; }
        public string HeaderColor { get; set; } = "365F91";
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
            throw new NotImplementedException();
        }

        public override string GetAlignment()
        {
            throw new NotImplementedException();
        }

        public override int GetMarginLeft()
        {
            throw new NotImplementedException();
        }

        public override int GetMarginRight()
        {
            throw new NotImplementedException();
        }

        public override int GetMarginTop()
        {
            throw new NotImplementedException();
        }

        public override int GetMarginBottom()
        {
            throw new NotImplementedException();
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
