using CommonLibrary;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HelperLibrary;
using OpenXmlHelperLibrary;
using ProcessDocumentCore.Interface;
using StandardsLibrary;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ProcessDocumentCore.Processing
{
    public class ProcessingOpenXml : IDocumentProcessing
    {
        private const string ExtensionDoc = ".docx";
        public ResultExecute Processing(GostModel gostModel, string filePath)
        {

            _gostRepository = new GostGenericRepository<GostModel>(gostModel);
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
                        if ((para.Elements<BookmarkStart>().Any(p => p.Name != "_GoBack") && para.ToList().Any(p => p is BookmarkEnd))|| para.Elements<BookmarkStart>().Any(p => p.Name != "_GoBack"))
                        {
                            foreach (var openXmlElement1 in para.ToList().Where(x => x is Run).ToList())
                            {
                                var openXmlElement = (Run)openXmlElement1;
                                Debug.WriteLine(openXmlElement.InnerText);
                                ////Определяем есть ли нумерованный список для задания стиля всему параграфу, т.к. на него распотсраняется отдельный стиль
                                //var hasNumberingProperties = para.ParagraphProperties.FirstOrDefault(o =>
                                //    o.GetType() == typeof(NumberingProperties));
                                //if (hasNumberingProperties != null)
                                    isNeedChangeStyleForParagraph = true;

                                SetRunStyle(openXmlElement, CommonGost.StyleTypeEnum.Headline);
                            }
                            //    if (isHeader == false)
                            //    {
                            //        //var p = new OpenXmlGenericRepositoryParagraph<Paragraph>(para);
                            //        //p.ClearAll();
                            //        //p.Justification(_designStandard.GetAlignment());

                            //        //// Выранивание для текста
                            //        //var textAlignBody = para.ParagraphProperties.FirstOrDefault(x =>
                            //        //    x.GetType() == typeof(Justification));
                            //        //if (textAlignBody != null)
                            //        //{
                            //        //    var _el = (Justification)textAlignBody;
                            //        //    _el.Val = (Extension.GetJustificationByString(_designStandard.GetAlignment()));
                            //        //}
                            //        //else
                            //        //{
                            //        //    var _el = new Justification() { Val = Extension.GetJustificationByString(_designStandard.GetAlignment()) };
                            //        //    para.ParagraphProperties.Append(_el);
                            //        //}
                            //    }
                            //    else
                            //    {
                            //        var p = new OpenXmlGenericRepositoryParagraph<Paragraph>(para);
                            //        p.ClearAll();
                            //        p.Justification(_gostRepository.GetAlignment(CommonGost.StyleTypeEnum.GlobalText));
                            //        //// Определяем положение выравнивания для заголовков
                            //        //var textAlignHead = para.ParagraphProperties.FirstOrDefault(x =>
                            //        //x.GetType() == typeof(Justification));
                            //        //if (textAlignHead != null)
                            //        //{
                            //        //    var _el = (Justification)textAlignHead;
                            //        //    _el.Val = (Extension.GetJustificationByString(_designStandard.GetAlignment()));
                            //        //}
                            //        //else
                            //        //{
                            //        //    var _el = new Justification() { Val = Extension.GetJustificationByString(_designStandard.GetAlignment()) };
                            //        //    para.ParagraphProperties.Append(_el);
                            //        //}
                            //    }
                        }

                       

                        if (isNeedChangeStyleForParagraph)
                        {
                            SetParagraphStyle(para, CommonGost.StyleTypeEnum.Headline);
                            //SetParagraphStyle(para);
                        }
                        else
                        {
                            //var p = new OpenXmlGenericRepositoryParagraph<Paragraph>(para);
                            ////p.ClearAll();
                            //p.Justification(_gostRepository.GetAlignment(CommonGost.StyleTypeEnum.GlobalText));

                            //Форматирование абзацев
                            if (para.Elements<BookmarkStart>().All(p => p.Name == "_GoBack") || !para.ToList().Any(p => p is BookmarkStart))
                            {


                                foreach (var runs in para.Elements<Run>()) //Форматирование шрифта и его размера для каждого run'а
                                {
                                    SetRunStyle(runs, CommonGost.StyleTypeEnum.GlobalText);
                                }

                                SetParagraphStyle(para, CommonGost.StyleTypeEnum.GlobalText);
                            }
                        }
                    }

                    CorrectImage(body);

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

        private void SetRunStyle(Run openXmlElement, CommonGost.StyleTypeEnum typeStyle)
        {
            if (openXmlElement == null) return;

            var p = new OpenXmlGenericRepositoryRun<Run>(openXmlElement);
            p.ClearAll();
            if (_gostRepository.GetFontSize(typeStyle) != null) p.FontSize(_gostRepository.GetFontSize(typeStyle).SafeToInt(-1));
            if (_gostRepository.GetColor(typeStyle) != null) p.Color(_gostRepository.GetColor(typeStyle));
            if (_gostRepository.GetBold(typeStyle) != null) p.Bold(_gostRepository.GetBold(typeStyle).nvl());
            if (_gostRepository.GetFont(typeStyle) != null) p.RunFonts(_gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle));
        }


        private GostGenericRepository<GostModel> _gostRepository;


        //private void SetParagraphStyle(Paragraph para)//todo переправить на новую реализацию
        //{
        //    //задаем стили для параграфа, т.к. если в параграфе имеется нумерация, то стиль берется общий для параграфа
        //    if (para.ParagraphProperties?.ParagraphMarkRunProperties != null)
        //    {
        //        var fontSize = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x =>
        //            x.GetType() == typeof(FontSize));
        //        var color = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x =>
        //            x.GetType() == typeof(Color));

        //        var bold = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x =>
        //            x.GetType() == typeof(Bold));

        //        var runFonts = para.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x =>
        //            x.GetType() == typeof(RunFonts));

        //        if (fontSize != null)
        //        {
        //            var el = (FontSize)fontSize;
        //            el.Val = (_designStandard.GetFontSize() * 2).ToString();
        //        }
        //        else
        //        {
        //            var _size = new FontSize { Val = (_designStandard.GetFontSize() * 2).ToString() };
        //            para.ParagraphProperties.ParagraphMarkRunProperties.Append(_size);
        //        }

        //        if (bold != null)
        //        {
        //            var el = (Bold)bold;
        //            el.Val = _designStandard.isBold();
        //        }
        //        else
        //        {
        //            var _bold = new Bold { Val = _designStandard.isBold() };
        //            para.ParagraphProperties.ParagraphMarkRunProperties.Append(_bold);
        //        }

        //        runFonts?.Remove();

        //        var _runFonts = new RunFonts() { Ascii = _designStandard.GetFont(), HighAnsi = _designStandard.GetFont() };
        //        para.ParagraphProperties.ParagraphMarkRunProperties.Append(_runFonts);

        //        if (color != null)
        //        {
        //            var el = (Color)color;
        //            el.Val = _designStandard.GetHeaderColor();
        //        }
        //        else
        //        {
        //            var _color = new Color { Val = _designStandard.GetHeaderColor() };
        //            para.ParagraphProperties.ParagraphMarkRunProperties.Append(_color);
        //        }
        //    }
        //}



        private void CorrectImage(Body body)
        {
            if (body == null) return;

            var isNextRunIsHeaderImg = false;

            foreach (var item in body.Elements<Paragraph>().ToList())
            {
                if (isNextRunIsHeaderImg)
                {
                    foreach (var itemRun in item)
                    {
                        if (itemRun is Run run) SetRunStyle(run, CommonGost.StyleTypeEnum.Image);
                    }
                    SetParagraphStyle(item, CommonGost.StyleTypeEnum.Image);
                    isNextRunIsHeaderImg = false;

                    if (!item.Any(r => r.GetType() == typeof(Run)))
                    {
                        item.Remove();
                        isNextRunIsHeaderImg = true;
                        continue;
                    }
                }

                var findDrawing = item.Any(f => f.ToList().Any(e => e is Drawing));
                if (findDrawing)
                {
                    isNextRunIsHeaderImg = true;
                    SetParagraphStyle(item, CommonGost.StyleTypeEnum.Image);
                }
            }
        }

        private void SetParagraphStyle(Paragraph para, IDictionary<object, string> styleDictionary)
        {
            if (para == null) return;
            if (styleDictionary == null) return;

            var p = new OpenXmlGenericRepositoryParagraph<Paragraph>(para);
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

            if (styleDictionary.ContainsKey("lineSpacing"))
            {
                p.LineSpacing(styleDictionary.GetVolStyle("lineSpacing"));
            }
            if (styleDictionary.ContainsKey("beforeSpacing"))
            {
                p.BeforeSpacing(styleDictionary.GetVolStyle("beforeSpacing"));
            }
            if (styleDictionary.ContainsKey("afterSpacing"))
            {
                p.AfterSpacing(styleDictionary.GetVolStyle("afterSpacing"));
            }

            if (styleDictionary.ContainsKey("firstlineIndent"))
            {
                p.FirstLineIndent(styleDictionary.GetVolStyle("firstlineIndent"));
            }
            if (styleDictionary.ContainsKey("leftIndent"))
            {
                p.LeftIndent(styleDictionary.GetVolStyle("leftIndent"));
            }
            if (styleDictionary.ContainsKey("rightIndent"))
            {
                p.RightIndent(styleDictionary.GetVolStyle("rightIndent"));
            }
        }

        private void SetParagraphStyle(Paragraph para, CommonGost.StyleTypeEnum typeStyle)
        {

            if (para == null) return;

            var p = new OpenXmlGenericRepositoryParagraph<Paragraph>(para);
            p.ClearAll();
            if (_gostRepository.GetFontSize(typeStyle) != null) p.FontSize(_gostRepository.GetFontSize(typeStyle).SafeToInt(-1));
            if (_gostRepository.GetColor(typeStyle) != null) p.Color(_gostRepository.GetColor(typeStyle));
            if (_gostRepository.GetBold(typeStyle) != null) p.Bold(_gostRepository.GetBold(typeStyle).nvl());
            if (_gostRepository.GetFont(typeStyle) != null) p.RunFonts(_gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle));
            if (_gostRepository.GetAlignment(typeStyle) != null) p.Justification(_gostRepository.GetAlignment(typeStyle));
            p.SpacingBetweenLines(_gostRepository.GetLineSpacing(typeStyle).nvl(), _gostRepository.GetBeforeSpacing(typeStyle).nvl(), _gostRepository.GetAfterSpacing(typeStyle).nvl());
            p.Indentation(_gostRepository.GetFirstLineIndentation(typeStyle).nvl(), _gostRepository.GetLeftIndentation(typeStyle).nvl(), _gostRepository.GetRightIndentation(typeStyle).nvl());
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
