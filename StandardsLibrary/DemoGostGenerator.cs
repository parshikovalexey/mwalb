using StandardsLibrary.Simple;
using System;

namespace StandardsLibrary
{
    public static class DemoGostGenerator
    {
        public static GostModel Demo = new GostModel()
        {
            Header = $"Тестовый гост {DateTime.Now.ToString("dd-MM-yyyy")}",
            GlobalText = new SimpleStyle()
            {
                Font = "Times New Roman",
                FontSize = 15,
                Alignment = "Both",
                Bold = false,
                AfterSpacing = 1,
                BeforeSpacing = 1,
                FirstLineIndentation = 1.5f,
                LeftIndentation = 0,
                LineSpacing = 1.5f,
                MarginBottom = 0,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                RightIndentation = 0,
                Color = "Black",
                Italic = false,
                Underline = ""
            },
            Headline = new SimpleStyle()
            {
                Font = "Times New Roman",
                FontSize = 30,
                Alignment = "Center",
                Bold = true,
                AfterSpacing = 1,
                BeforeSpacing = 1,
                FirstLineIndentation = 1.5f,
                LeftIndentation = 0,
                LineSpacing = 1.5f,
                MarginBottom = 0,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                RightIndentation = 0,
                Color = "365F91",
                Italic = false,
                Underline = ""
            },
            Image = new SimpleStyle()
            {
                Font = "Times New Roman",
                FontSize = 30,
                Alignment = "Center",
                Bold = true,
                AfterSpacing = 1,
                BeforeSpacing = 1,
                FirstLineIndentation = 1.5f,
                LeftIndentation = 0,
                LineSpacing = 1.5f,
                MarginBottom = 0,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                RightIndentation = 0,
            },
            FooterPart = new SimpleStyle()
            {
                Font = "Times New Roman",
                FontSize = 15,
                Alignment = "Both",
                Bold = false,
                AfterSpacing = 1,
                BeforeSpacing = 1,
                FirstLineIndentation = 1.5f,
                LeftIndentation = 0,
                LineSpacing = 1.5f,
                MarginBottom = 0,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                RightIndentation = 0,
                Color = "Black",
                Italic = false,
                Underline = ""
            },
            HeaderPart = new SimpleStyle()
            {
                Font = "Times New Roman",
                FontSize = 15,
                Alignment = "Both",
                Bold = false,
                AfterSpacing = 1,
                BeforeSpacing = 1,
                FirstLineIndentation = 1.5f,
                LeftIndentation = 0,
                LineSpacing = 1.5f,
                MarginBottom = 0,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                RightIndentation = 0,
                Color = "Black",
                Italic = false,
                Underline = ""
            },
            ImageCaption = new SimpleStyle()
            {
                Font = "Times New Roman",
                FontSize = 15,
                Alignment = "Both",
                Bold = false,
                AfterSpacing = 1,
                BeforeSpacing = 1,
                FirstLineIndentation = 1.5f,
                LeftIndentation = 0,
                LineSpacing = 1.5f,
                MarginBottom = 0,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                RightIndentation = 0,
                Color = "Black",
                Italic = false,
                Underline = ""
            },
            Numbering = new SimpleNumbering()
            {
                LeftIndentation = 5,
                LeftNextIndentation = 5,
                Hanging = 2,
                Level1 = new SimpleNumberingLevel() { NumberingFormat = SimpleNumberingLevel.NumberingEnum.Decimal.ToString(), LevelText = "%1" },
                Level2 = new SimpleNumberingLevel() { NumberingFormat = SimpleNumberingLevel.NumberingEnum.Decimal.ToString(), LevelText = "%1.%2." },
                Level3 = new SimpleNumberingLevel() { NumberingFormat = SimpleNumberingLevel.NumberingEnum.Decimal.ToString(), LevelText = "%1.%2.%3." },
                Level4 = new SimpleNumberingLevel() { NumberingFormat = SimpleNumberingLevel.NumberingEnum.Decimal.ToString(), LevelText = "%1.%2.%3.%4." },
                Level5 = new SimpleNumberingLevel() { NumberingFormat = SimpleNumberingLevel.NumberingEnum.Decimal.ToString(), LevelText = "%1.%2.%3.%4.%5." },
                Level6 = new SimpleNumberingLevel() { NumberingFormat = SimpleNumberingLevel.NumberingEnum.Decimal.ToString(), LevelText = "%1.%2.%3.%4.%5.%6." },
                Level7 = new SimpleNumberingLevel() { NumberingFormat = SimpleNumberingLevel.NumberingEnum.Decimal.ToString(), LevelText = "%1.%2.%3.%4.%5.%6.%7." },
                Level8 = new SimpleNumberingLevel() { NumberingFormat = SimpleNumberingLevel.NumberingEnum.Decimal.ToString(), LevelText = "%1.%2.%3.%4.%5.%6.%7.%8." },
                Level9 = new SimpleNumberingLevel() { NumberingFormat = SimpleNumberingLevel.NumberingEnum.Decimal.ToString(), LevelText = "%1.%2.%3.%4.%5.%6.%7.%8.%9." }

            }
        };
    }
}
