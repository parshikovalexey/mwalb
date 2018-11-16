using System;
using System.IO;
using Newtonsoft.Json;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StandardsLibrary
{
    public class Gost1 : Standards
    {
        public string Font { get; set; }
        public int FontSize { get; set; }
        public float LineSpacing { get; set; }
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

        public override bool isBold() => Bold;

        public override string GetHeaderColor() => HeaderColor;
    }
}
