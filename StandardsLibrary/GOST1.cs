using System;
using System.IO;
using Newtonsoft.Json;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace StandardsLibrary
{
    public class Gost1 : Standards
    {
        public string Font { get; set; }
        public int FontSize { get; set; }
        public float LineSpacing { get; set; }
        public JustificationValues Alignment { get; set; }
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
        public string HeaderColor { get; set; }
        public bool Bold { get; set; }
        public Gost1(string path)
        {
            Dictionary<string, string> gost =
                JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(path));
            Font = gost["Font"];
            FontSize = Int32.Parse(gost["FontSize"]);
            LineSpacing = Convert.ToSingle(gost["LineSpacing"], new CultureInfo("en-US"));
            BeforeSpacing = Convert.ToSingle(gost["BeforeSpacing"], new CultureInfo("en-US"));
            AfterSpacing = Convert.ToSingle(gost["AfterSpacing"], new CultureInfo("en-US"));
            FirstLineIndentation = Convert.ToSingle(gost["FirstLineIndentation"], new CultureInfo("en-US"));
            LeftIndentation = Convert.ToSingle(gost["LeftIndentation"], new CultureInfo("en-US"));
            RightIndentation = Convert.ToSingle(gost["RightIndentation"], new CultureInfo("en-US"));
            Alignment = gost["Alignment"];
            MarginLeft = Int32.Parse(gost["MarginLeft"]);
            MarginRight = Int32.Parse(gost["MarginRight"]);
            MarginTop = Int32.Parse(gost["MarginTop"]);
            MarginBottom = int.Parse(gost["MarginBottom"]);
            HeaderColor = gost["HeaderColor"];
            Bold = bool.Parse(gost["Bold"]);

        }
        


        public override string GetFont() => Font;

        public override int GetFontSize() => FontSize;

        public override float GetLineSpacing() => LineSpacing;

        public override float GetBeforeSpacing() => BeforeSpacing;

        public override float GetAfterSpacing() => AfterSpacing;

        public override float GetFirstLineIndentation() => FirstLineIndentation;

        public override float GetLeftIndentation() => LeftIndentation;

        public override float GetRightIndentation() => RightIndentation;

        public override JustificationValues GetAlignment() => Alignment;
        
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

        public override bool isBold() => Bold;

        public override string GetHeaderColor() => HeaderColor;
    }
}

