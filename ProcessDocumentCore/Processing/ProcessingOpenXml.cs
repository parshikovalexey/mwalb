using CommonLibrary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HelperLibrary;
using OpenXmlHelperLibrary;
using ProcessDocumentCore.Interface;
using StandardsLibrary;
using System;
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

                    //устанавливаем отступы для всего документа
                    SetPageMargin(body);


                    var isHeader = false;

                    foreach (var para in body.Elements<Paragraph>())
                    {

                        bool isNeedChangeStyleForParagraph = false;
                        if ((para.Elements<BookmarkStart>().Any(p => p.Name != "_GoBack") && para.ToList().Any(p => p is BookmarkEnd)) || para.Elements<BookmarkStart>().Any(p => p.Name != "_GoBack"))
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
                            SetParagraphStyle(para, CommonGost.StyleTypeEnum.GlobalText);
                            foreach (var runs in para.Elements<Run>()) //Форматирование шрифта и его размера для каждого run'а
                                SetRunStyle(runs, CommonGost.StyleTypeEnum.GlobalText);
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

        private void SetPageMargin(Body body)
        {
            if (body == null) {
                LoggerLibrary.Logger.Write().Error("Объект body null");
                return;
            }

            try {
                PageMargin pgMar = body.Descendants<PageMargin>().FirstOrDefault();
                if (pgMar != null) {
                    pgMar.Top = _gostRepository.GetMarginTop(CommonGost.StyleTypeEnum.GlobalText);
                    pgMar.Bottom = _gostRepository.GetMarginBottom(CommonGost.StyleTypeEnum.GlobalText);
                    pgMar.Left = new UInt32Value(_gostRepository.GetMarginLeft(CommonGost.StyleTypeEnum.GlobalText).SafeToUint());
                    pgMar.Right = new UInt32Value(_gostRepository.GetMarginRight(CommonGost.StyleTypeEnum.GlobalText).SafeToUint());
                }
            }
            catch (Exception e) {
                LoggerLibrary.Logger.Write().Error(e);
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
            if (_gostRepository.GetFont(typeStyle) != null) p.RunFonts(_gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle));
        }


        private GostGenericRepository<GostModel> _gostRepository;

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
                        if (itemRun is Run run) SetRunStyle(run, CommonGost.StyleTypeEnum.ImageCaption);
                    }
                    SetParagraphStyle(item, CommonGost.StyleTypeEnum.ImageCaption);
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


        private string GetPathToSaveObj()
        {
            return Path.GetFullPath(Path.Combine(PathHelper.TmpDirectory(), $"{Guid.NewGuid().ToString()}{ExtensionDoc}"));
        }
    }

}
