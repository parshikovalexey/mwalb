using System;
using StandardsLibrary.Simple;

namespace StandardsLibrary
{
    public class GostModel
    {
        public string Header { get; set; }  = String.Empty;
        public SimpleStyle GlobalText { get; set; } = new SimpleStyle();
        public SimpleStyle Image { get; set; } = new SimpleStyle();
        public SimpleImageCaption ImageCaption { get; set; } = new SimpleImageCaption();
        public SimpleStyle Headline { get; set; } = new SimpleStyle();
        public SimpleStyle HeaderPart { get; set; } = new SimpleStyle();
        public SimpleStyle FooterPart { get; set; } = new SimpleStyle();
        public SimpleNumbering Numbering { get; set;} = new SimpleNumbering();
        public SimpleNumbering Bullet { get; set; } = new SimpleNumbering();
        public SimpleTOC TOC { get; set; } = new SimpleTOC();
}
}
