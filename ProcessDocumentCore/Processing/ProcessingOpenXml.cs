using CommonLibrary;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ProcessDocumentCore.Interface;
using StandardsLibrary;
using System;
using System.IO;
using System.Linq;

namespace ProcessDocumentCore.Processing
{
    public class ProcessingOpenXml : IDocumentProcessing
    {
        private Standards _designStandard;
        private const string ExtensionDoc = ".docx";
        public ResultExecute Processing(Standards designStandard, string filePath)
        {
            _designStandard = designStandard;
            PathHelper.ClearTmpDirectory();
            Stream stream = null;
            WordprocessingDocument wordDoc = null;
            var pathToSaveObj = GetPathToSaveObj();
            try
            {
                File.Copy(filePath, pathToSaveObj);
                stream = File.Open(pathToSaveObj, FileMode.Open);

                using (wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    string docText;
                    using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd();
                    }

                    var body = wordDoc.MainDocumentPart.Document.Body;

                    foreach (var para in body.Elements<Paragraph>())
                    {
                        bool isNeedChangeStyleForParagraph = false;
                        if (para.Elements<BookmarkStart>().Any(p => p.Name != "_GoBack") && para.ToList().Any(p => p is BookmarkEnd))
                        {
                            foreach (var openXmlElement1 in para.ToList().Where(x => x is Run).ToList())
                            {
                                var openXmlElement = (Run)openXmlElement1;
                                //Определяем есть ли нумерованный список для задания стиля всему параграфу, т.к. на него распотсраняется отдельный стиль
                                var hasNumberingProperties = para.ParagraphProperties.FirstOrDefault(o =>
                                    o.GetType() == typeof(NumberingProperties));
                                if (hasNumberingProperties != null) isNeedChangeStyleForParagraph = true;

                                RunProperties runProperties = openXmlElement.RunProperties;
                                if (runProperties == null) continue;
                                runProperties.FontSize = new FontSize() { Val = (_designStandard.GetFontSize() * 2).ToString() };
                                runProperties.Color = new Color() { Val = _designStandard.GetHeaderColor() };
                                runProperties.Bold = new Bold() { Val = _designStandard.isBold() };
                                runProperties.RunFonts = new RunFonts() { Ascii = _designStandard.GetFont(), HighAnsi = _designStandard.GetFont() };
                            }
                        }
                        if (isNeedChangeStyleForParagraph) SetParagraphStyle(para);

                    }

                    using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }

                    wordDoc.Close();
                    stream.Close();
                }
                return new ResultExecute() { Callbacks = pathToSaveObj };
            }
            catch (Exception ex)
            {
                wordDoc?.Close();
                stream?.Close();
                return new ResultExecute().OnError(ex.Message);
            }
        }

        private void SetParagraphStyle(Paragraph para)
        {
            try
            {
                //задаем стили для параграфа, т.к. если в параграфе имеется нумерация, то стиль берется общий для параграфа
                if (para.ParagraphProperties?.ParagraphMarkRunProperties != null)
                {
                    var fontSize = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x =>
                        x.GetType() == typeof(FontSize));
                    var color = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x =>
                        x.GetType() == typeof(Color));

                    var bold = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x =>
                        x.GetType() == typeof(Bold));

                    var runFonts = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x =>
                        x.GetType() == typeof(RunFonts));

                    if (fontSize != null)
                    {
                        var el = (FontSize)fontSize;
                        el.Val = (_designStandard.GetFontSize() * 2).ToString();
                    }
                    else
                    {
                        var _size = new FontSize { Val = (_designStandard.GetFontSize() * 2).ToString() };
                        para.ParagraphProperties.ParagraphMarkRunProperties.Append(_size);
                    }

                    if (bold != null)
                    {
                        var el = (Bold)bold;
                        el.Val = _designStandard.isBold();
                    }
                    else
                    {
                        var _bold = new Bold { Val = _designStandard.isBold() };
                        para.ParagraphProperties.ParagraphMarkRunProperties.Append(_bold);
                    }

                    runFonts?.Remove();

                    var _runFonts = new RunFonts() { Ascii = _designStandard.GetFont(), HighAnsi = _designStandard.GetFont() };
                    para.ParagraphProperties.ParagraphMarkRunProperties.Append(_runFonts);

                    if (color != null)
                    {
                        var el = (Color)color;
                        el.Val = "365F91";
                    }
                    else
                    {
                        var _color = new Color { Val = _designStandard.GetHeaderColor() };
                        para.ParagraphProperties.ParagraphMarkRunProperties.Append(_color);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

        }

        private string GetPathToSaveObj()
        {
            return Path.GetFullPath(Path.Combine(PathHelper.TmpDirectory(), $"{Guid.NewGuid().ToString()}{ExtensionDoc}"));
        }
    }
}
