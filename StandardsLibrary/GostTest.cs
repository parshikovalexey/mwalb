using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace StandardsLibrary
{
    public class GostTest:Standards
    {
        public override string GetFont()
        {
            return "Times New Roman";
        }

        public override int GetFontSize()
        {
            return 30;
        }

        public override float GetLineSpacing()
        {
            throw new NotImplementedException();
        }

        public override JustificationValues GetAlignment() => JustificationValues.Right;
        
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
            return true;
        }

        public override string GetHeaderColor()
        {
            return "365F91";
        }
    }
}
