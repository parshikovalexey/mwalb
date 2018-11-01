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
        private const string ExtensionDoc = ".docx";
        public ResultExecute Processing(Standards designStandard, string filePath)
        {
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

                    var isHeader = false;

                    foreach (var para in body.Elements<Paragraph>())
                    {
                        bool isNeedChangeStyleForParagraph = false;
                        foreach (var itemPara in para)
                        {
                            if (itemPara is BookmarkStart) isHeader = true;

                            if (itemPara is BookmarkEnd) isHeader = false;

                            if (itemPara is Run run && isHeader)
                            {
                                //Определяем есть ли нумерованный список для задания стиля всему параграфу, т.к. на него распотсраняется отдельный стиль
                                var hasNumberingProperties = para.ParagraphProperties.FirstOrDefault(o =>
                                    o.GetType() == typeof(NumberingProperties));
                                if (hasNumberingProperties != null) isNeedChangeStyleForParagraph = true;

                                RunProperties runProperties = run.RunProperties;
                                runProperties.FontSize = new FontSize() { Val = "30" };
                                runProperties.Color = new Color() { Val = "365F91" };
                                runProperties.Bold = new Bold() { Val = true };
                                runProperties.RunFonts = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
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
            return new ResultExecute().OnError("Что то не так");
        }

        private void SetParagraphStyle(Paragraph para)
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
                    el.Val = "30";
                }
                else
                {
                    var _size = new FontSize { Val = "30" };
                    para.ParagraphProperties.ParagraphMarkRunProperties.Append(_size);
                }

                if (bold != null)
                {
                    var el = (Bold)bold;
                    el.Val = true;
                }
                else
                {
                    var _bold = new Bold { Val = true };
                    para.ParagraphProperties.ParagraphMarkRunProperties.Append(_bold);
                }

                runFonts?.Remove();

                var _runFonts = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                para.ParagraphProperties.ParagraphMarkRunProperties.Append(_runFonts);


                if (color != null)
                {
                    var el = (Color)color;
                    el.Val = "365F91";
                }
                else
                {
                    var _color = new Color { Val = "365F91" };
                    para.ParagraphProperties.ParagraphMarkRunProperties.Append(_color);
                }
            }
        }

        private string GetPathToSaveObj()
        {
            return Path.GetFullPath(Path.Combine(PathHelper.TmpDirectory(), $"{Guid.NewGuid().ToString()}{ExtensionDoc}"));
        }
    }
}
