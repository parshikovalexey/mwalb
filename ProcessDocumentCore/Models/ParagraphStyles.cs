using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using StandardsLibrary;

namespace ProcessDocumentCore.Models {
    public class ParagraphStyles {
        public FontSize FontSize { get; set; }
        public Color Color { get; set; }
        public Bold Bold { get; set; }
        public RunFonts RunFonts { get; set; }

        public ParagraphStyles(Standards standards) {
            FontSize = new FontSize { Val = (standards.GetFontSize() * 2).ToString() };
            Color = new Color { Val = standards.GetHeaderColor() };
            Bold = new Bold { Val = standards.isBold() };
            RunFonts = new RunFonts { Ascii = standards.GetFont(), HighAnsi = standards.GetFont() };
        }

        /// <summary>
        /// Устанавливает в RunProperties значения стилей 
        /// </summary>
        /// <param name="runProperties">Элемент параграфа</param>
        /// <returns>Измененный объект RunProperties</returns>
        public RunProperties SetParagraphStylesForRunProperties(RunProperties runProperties) {
            runProperties.FontSize = FontSize;
            runProperties.Color = Color;
            runProperties.Bold = Bold;
            runProperties.RunFonts = RunFonts;
            return runProperties;
        }

        /// <summary>
        /// Возращает коллекцию элементов стилей
        /// </summary>
        /// <returns></returns>
        public IEnumerable<OpenXmlElement> GetCollectionOfElements() {
            return new List<OpenXmlElement> { FontSize, Color, Bold, RunFonts };
        }

        /// <summary>
        /// Удаляет из переданного параграфа ChildElements набора стилей
        /// </summary>
        /// <param name="paragraph">Объект параграфа openXml</param>
        public void RemoveChildElementFromParagraph(Paragraph paragraph) {
            var properties = GetType().GetProperties();
            foreach (var property in properties) {
                Type type = Type.GetType(property.Name);
                paragraph.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x => x.GetType() == type)?.Remove();
            }
        }
    }
}
