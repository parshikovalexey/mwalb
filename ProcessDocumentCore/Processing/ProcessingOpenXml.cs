using CommonLibrary;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HelperLibrary;
using ProcessDocumentCore.Interface;
using StandardsLibrary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OpenXmlHelperLibrary;

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


                    CorrectImage(body);
                    var isHeader = false;
                   
                    foreach (var para in body.Elements<Paragraph>())
                    {
                        
                        bool isNeedChangeStyleForParagraph = false;
                        foreach (var itemPara in para)
                        {

                            if (itemPara is BookmarkStart)  isHeader = true;

                            if (itemPara is BookmarkEnd) isHeader = false;

                            if (itemPara is Run run && isHeader)
                            {
                                //Определяем есть ли нумерованный список для задания стиля всему параграфу, т.к. на него распотсраняется отдельный стиль
                                var hasNumberingProperties = para.ParagraphProperties.FirstOrDefault(o =>
                                    o.GetType() == typeof(NumberingProperties));
                                if (hasNumberingProperties != null) isNeedChangeStyleForParagraph = true;

                                RunProperties runProperties = run.RunProperties;
                                runProperties.FontSize = new FontSize() { Val = (_designStandard.GetFontSize() * 2).ToString() };
                                runProperties.Color = new Color() { Val = _designStandard.GetHeaderColor() };
                                runProperties.Bold = new Bold() { Val = _designStandard.isBold() };
                                runProperties.RunFonts = new RunFonts() { Ascii = _designStandard.GetFont(), HighAnsi = _designStandard.GetFont() };

                            }

                            //Форматирование абзацев
                            if (!para.ToList().Any(p => p is BookmarkStart) && !para.ToList().Any(p => p is BookmarkEnd) || (para.Elements<BookmarkStart>().All(p => p.Name == "_GoBack")))
                            {
                                foreach (var runs in para.Elements<Run>()) //Форматирование шрифта и его размера для каждого run'а
                                {
                                    if (runs.RunProperties == null) runs.RunProperties = new RunProperties();
                                    runs.RunProperties.RunFonts = new RunFonts() { Ascii = _designStandard.GetFont(), HighAnsi = _designStandard.GetFont() };
                                    runs.RunProperties.FontSize = new FontSize() { Val = (_designStandard.GetFontSize() * 2).ToString() };
                                }

                                if (para.ParagraphProperties == null) para.ParagraphProperties = new ParagraphProperties();

                                para.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines() //Интервалы между строками и абзацами
                                {
                                    Line = (_designStandard.GetLineSpacing() * 240).ToString(),
                                    Before = (_designStandard.GetBeforeSpacing() * 20).ToString(),
                                    After = (_designStandard.GetAfterSpacing() * 20).ToString(),
                                    BeforeAutoSpacing = new DocumentFormat.OpenXml.OnOffValue(false),
                                    AfterAutoSpacing = new DocumentFormat.OpenXml.OnOffValue(false)
                                };

                                if (para.ParagraphProperties.ParagraphMarkRunProperties != null) //Тут форматируется шрифт знака абзаца
                                {
                                    var runfont = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(p => p is RunFonts);
                                    var fontsize = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(p => p is FontSize);
                                    runfont?.Remove(); 
                                    runfont = new RunFonts() { Ascii = _designStandard.GetFont(), HighAnsi = _designStandard.GetFont() };
                                    para.ParagraphProperties.ParagraphMarkRunProperties.Append(runfont);
                                    fontsize?.Remove();
                                    fontsize = new FontSize() { Val = (_designStandard.GetFontSize() * 2).ToString() };
                                    para.ParagraphProperties.ParagraphMarkRunProperties.Append(fontsize);
                                }
                                else
                                {
                                    para.ParagraphProperties.ParagraphMarkRunProperties = new ParagraphMarkRunProperties();
                                    var runfont = new RunFonts() { Ascii = _designStandard.GetFont(), HighAnsi = _designStandard.GetFont() };
                                    var fontsize = new FontSize() { Val = (_designStandard.GetFontSize() * 2).ToString() };
                                    para.ParagraphProperties.ParagraphMarkRunProperties.Append(runfont);
                                    para.ParagraphProperties.ParagraphMarkRunProperties.Append(fontsize);
                                }

                                para.ParagraphProperties.Indentation = new Indentation() //Отступы
                                {
                                    FirstLine = ((int)(_designStandard.GetFirstLineIndentation() * 567)).ToString(),
                                    Left = ((int)(_designStandard.GetLeftIndentation() * 567)).ToString(), 
                                    Right = ((int)(_designStandard.GetRightIndentation() * 567)).ToString(),
                                };
                            }

                            if (isHeader == false)
                            {
                                // Выранивание для текста
                                var textAlignBody = para.ParagraphProperties.FirstOrDefault(x =>
                                    x.GetType() == typeof(Justification));
                                if (textAlignBody != null)
                                {
                                    var _el = (Justification)textAlignBody;
                                    _el.Val = (Extension.GetJustificationByString(_designStandard.GetAlignment()));
                                }
                                else
                                {
                                    var _el = new Justification() { Val = Extension.GetJustificationByString(_designStandard.GetAlignment()) };
                                    para.ParagraphProperties.Append(_el);
                                }
                            }
                            else
                            {
                                // Определяем положение выравнивания для заголовков
                                var textAlignHead = para.ParagraphProperties.FirstOrDefault(x =>
                                x.GetType() == typeof(Justification));
                                if (textAlignHead != null)
                                {
                                    var _el = (Justification)textAlignHead;
                                    _el.Val = (Extension.GetJustificationByString(_designStandard.GetAlignment()));
                                }
                                else
                                {
                                    var _el = new Justification() { Val = Extension.GetJustificationByString(_designStandard.GetAlignment()) };
                                    para.ParagraphProperties.Append(_el);
                                }
                            }
                        }

                        if (isNeedChangeStyleForParagraph)  SetParagraphStyle(para); 
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

        private void CorrectImage(Body body)
        {
            if (body == null) return;

            var isNextRunIsHeaderImg = false;
            IDictionary<object, string> style = new Dictionary<object, string>();
            style.Add(typeof(FontSize), (_designStandard.GetFontSize() * 2).ToString());
            style.Add(typeof(Bold), _designStandard.isBold().ToString());
            style.Add(typeof(RunFonts), _designStandard.GetFont());
            style.Add(typeof(Justification), JustificationValues.Center.ToString());
            foreach (var item in body.Elements<Paragraph>())
            {
                if (isNextRunIsHeaderImg)
                {
                    foreach (var itemRun in item)
                    {
                        if (itemRun is Run run) SetRunStyle(run, style);
                    }
                    SetParagraphStyle(item, style);
                    isNextRunIsHeaderImg = false;

                }

                var findDrawing = item.Any(f => f.ToList().Any(e => e is Drawing));
                if (findDrawing)
                {
                    isNextRunIsHeaderImg = true;
                    SetParagraphStyle(item, style);
                }
            }
        }


        private void SetParagraphStyle(Paragraph para)//todo переправить на новую реализацию
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

        private void SetParagraphStyle(Paragraph para, IDictionary<object, string> styleDictionary)
        {
            if (para == null) return;
            if (styleDictionary == null) return;

            var p = new OpenXmlGenericRepository<Paragraph>(para);
            p.ClearAll();

            //para?.ParagraphProperties?.Remove();
            //para?.ParagraphProperties?.ParagraphMarkRunProperties?.Remove();
            //if (para.ParagraphProperties == null) para.ParagraphProperties = new ParagraphProperties();
            //para.ParagraphProperties.ParagraphMarkRunProperties = new ParagraphMarkRunProperties();

            if (styleDictionary.ContainsKey(typeof(FontSize)))
            {
                p.FontSize(styleDictionary.GetVolStyle(typeof(FontSize)).SafeToInt(0));
                //var newStyle = new FontSize { Val = styleDictionary.GetVolStyle(typeof(FontSize)) };
                //para.ParagraphProperties?.ParagraphMarkRunProperties.Append(newStyle);
            }

            if (styleDictionary.ContainsKey(typeof(Color)))
            {
                p.Color(styleDictionary.GetVolStyle(typeof(Color)));
                //var newStyle = new Color { Val = styleDictionary.GetVolStyle(typeof(Color)) };
                //para.ParagraphProperties?.ParagraphMarkRunProperties.Append(newStyle);
            }

            if (styleDictionary.ContainsKey(typeof(Bold)))
            {
                p.Bold(styleDictionary.GetVolStyle(typeof(Bold)).SafeToBoolean());
                //var newStyle = new Bold { Val = styleDictionary.GetVolStyle(typeof(Bold)).SafeToBoolean() };
                //para.ParagraphProperties?.ParagraphMarkRunProperties.Append(newStyle);
            }

            if (styleDictionary.ContainsKey(typeof(RunFonts)))
            {
                p.RunFonts(styleDictionary.GetVolStyle(typeof(RunFonts)), styleDictionary.GetVolStyle(typeof(RunFonts)));
                //var newStyle = new RunFonts { Ascii = styleDictionary.GetVolStyle(typeof(RunFonts)), HighAnsi = styleDictionary.GetVolStyle(typeof(RunFonts)) };
                //para.ParagraphProperties?.ParagraphMarkRunProperties.Append(newStyle);
            }

            if (styleDictionary.ContainsKey(typeof(Justification)))
            {
                p.Justification(styleDictionary.GetVolStyle(typeof(Justification)));
                //var t = para.ParagraphProperties;
                //if (t == null) return;
                //var newStyle = new Justification { Val = styleDictionary.GetVolStyle(typeof(Justification)).GetJustificationByString() };
                //t.Append(newStyle);
            }
        }
        private void SetRunStyle(Run run, IDictionary<object, string> styleDictionary)
        {
            if (run == null) return;
            if (styleDictionary == null) return;
            run?.RunProperties?.Remove();
            run.RunProperties = new RunProperties();
            RunProperties runProperties = run.RunProperties ?? new RunProperties();
            if (styleDictionary.ContainsKey(typeof(FontSize)))
            {
                var size = new FontSize { Val = styleDictionary.GetVolStyle(typeof(FontSize)) };
                runProperties.Append(size);
            }

            if (styleDictionary.ContainsKey(typeof(Color)))
            {
                var color = new Color { Val = styleDictionary.GetVolStyle(typeof(Color)) };
                runProperties.Append(color);
            }

            if (styleDictionary.ContainsKey(typeof(Bold)))
            {
                var bold = new Bold { Val = styleDictionary.GetVolStyle(typeof(Bold)).SafeToBoolean() };
                runProperties.Append(bold);
            }

            if (styleDictionary.ContainsKey(typeof(RunFonts)))
            {
                var fonts = new RunFonts
                {
                    Ascii = styleDictionary.GetVolStyle(typeof(RunFonts)),
                    HighAnsi = styleDictionary.GetVolStyle(typeof(RunFonts))
                };
                runProperties.Append(fonts);
            }
        }



        private string GetPathToSaveObj()
        {
            return Path.GetFullPath(Path.Combine(PathHelper.TmpDirectory(), $"{Guid.NewGuid().ToString()}{ExtensionDoc}"));
        }
    }

    public static class ExtensionsGost
    {
        /// <summary>
        /// Безопасное возвращение значения по ключу из словаря стилей
        /// </summary>
        /// <param name="styleDictionary">словарь стилей</param>
        /// <param name="key">typeof Класса</param>
        /// <param name="default">Значение если в словаре не окажется запрошенного ключа</param>
        /// <returns></returns>
        public static string GetVolStyle(this IDictionary<object, string> styleDictionary, object key, string @default = "")
        {
            if (@default == null) @default = string.Empty;
            styleDictionary.TryGetValue(key, out @default);
            return @default;
        }

       
    }
}
