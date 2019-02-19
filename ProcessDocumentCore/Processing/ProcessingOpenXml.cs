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
                        var isNeedClearProperty = true;
                        var isNeedChangeStyleForParagraph = false;
                        var isNumberingParagraph = false;
                        if (para.ParagraphProperties != null && para.ParagraphProperties.Elements<NumberingProperties>().Any())
                        {
                            isNumberingParagraph = true;
                            isNeedClearProperty = false;
                        }
                        else
                        {

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
                            }


                        }


                        if (isNeedChangeStyleForParagraph)
                        {
                            SetParagraphStyle(para, CommonGost.StyleTypeEnum.Headline, isNeedClearProperty);
                        }
                        else
                        {
                            
                            SetParagraphStyle(para, CommonGost.StyleTypeEnum.GlobalText, isNeedClearProperty);
                            foreach (var runs in para.Elements<Run>()) //Форматирование шрифта и его размера для каждого run'а
                                SetRunStyle(runs, CommonGost.StyleTypeEnum.GlobalText);
                                  
                                
                        }
                        //Делаем форматирование списков после того, как основной текст отформатирован
                        if (isNumberingParagraph)
                            SetNumberingProperties(para, wordDoc);
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

        private void SetNumberingProperties(Paragraph para, WordprocessingDocument wordDoc)
        {
            //Все стили для списка хранятся в AbstractNum основного документа и связанны цепочкой ParagraphProperties.NumberingId -> MainDocumentPart.NumberingInstance.Val -> AbstractNum.AbstractNumberId

            var numId = para.ParagraphProperties.FirstOrDefault(p => p.GetType() == typeof(NumberingProperties)).ToList()
                .FirstOrDefault(p => p.GetType() == typeof(NumberingId));

            //удаляем Indentation из параграфа, т.к. будут использоваться отступы из стилей нумерованных списков
            var pg = new OpenXmlGenericRepositoryParagraph<Paragraph>(para);
            pg.ClearSingleStyleFromProperties(typeof(Indentation));
            pg.ClearSingleStyleFromMarkRunProperties(typeof(Indentation));

            if (numId is NumberingId num)
            {
                int v = -1;

                var instances = wordDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.ToList()
                    .Where(p => p is NumberingInstance);

                //Ищем описания списка в основном документе по ID списка из параграфа
                foreach (var openXmlElement in instances)
                {
                    var instanceItem = (NumberingInstance)openXmlElement;
                    if (instanceItem.NumberID.ToString() == num.Val.ToString())
                    {
                        v = instanceItem.AbstractNumId.Val;
                        break;
                    }
                }


                var abs = wordDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.ToList()
                    .Where(p => p is AbstractNum);

                foreach (var item in abs)
                {
                    if (item is AbstractNum an)
                    {
                        if (an.AbstractNumberId == v)
                        {
                            var levels = an.Where(e => e is Level);
                            foreach (var itemLevel in levels)
                            {
                                if (itemLevel is Level level)
                                {

                                    var numberingFormat = _gostRepository.GetNumberingFormat(level.LevelIndex);
                                    var levelText = _gostRepository.GetNumberingLevelText(level.LevelIndex);

                                    SetlevelIndentation(level);
                                    SetLevelJustification(level);

                                    level.NumberingFormat = new NumberingFormat() { Val = numberingFormat };
                                    level.LevelText = new LevelText() { Val = levelText };

                                    var prop = level?.NumberingSymbolRunProperties;
                                    if (prop != null)
                                        level?.NumberingSymbolRunProperties.Remove();


                                    if (numberingFormat == NumberFormatValues.Bullet)
                                    {

                                        if (prop == null)
                                        {
                                            prop = new NumberingSymbolRunProperties();
                                            level.Append(prop);
                                        }
                                        else
                                        {
                                            prop = new NumberingSymbolRunProperties();
                                        }

                                        RunFonts runFonts1 = new RunFonts()
                                        { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

                                        prop.Append(runFonts1);
                                    }

                                }
                            }
                        }
                    }
                }
            }
        }

        private void SetlevelIndentation(Level level)
        {
            var paragraphProperties = level?.PreviousParagraphProperties;
            var indentation = paragraphProperties?.FirstOrDefault(p =>
                p.GetType() == typeof(Indentation));
            int multiplier = 567;
            if (indentation != null && indentation is Indentation levelIndentation)
            {
                levelIndentation.Left =
                    ((int)(_gostRepository.GetNumberingIndentationLeft(level.LevelIndex) * multiplier)).ToString();
                levelIndentation.Hanging = ((int)(_gostRepository.GetNumberingHanging() * multiplier)).ToString();
            }
        }

        private void SetLevelJustification(Level level)
        {
            var justification = level.LevelJustification;
            if (justification != null)
            {
                justification.Val = _gostRepository.GetNumberingJustification(level.LevelIndex);
            }
        }

        private void SetPageMargin(Body body)
        {
          
            if (body == null)
            {
                LoggerLibrary.Logger.Write().Error("Объект body null");
                return;
            }

          
            try
            {
                PageMargin pgMar = body.Descendants<PageMargin>().FirstOrDefault();
              
                if (pgMar != null)
                {
                    pgMar.Top = _gostRepository.GetMarginTop(CommonGost.StyleTypeEnum.GlobalText);
                    pgMar.Bottom = _gostRepository.GetMarginBottom(CommonGost.StyleTypeEnum.GlobalText);
                    pgMar.Left = new UInt32Value(_gostRepository.GetMarginLeft(CommonGost.StyleTypeEnum.GlobalText).SafeToUint());
                    pgMar.Right = new UInt32Value(_gostRepository.GetMarginRight(CommonGost.StyleTypeEnum.GlobalText).SafeToUint());
                }
            }
         
            catch (Exception e)
            {
                LoggerLibrary.Logger.Write().Error(e);
            }

        }

        private void SetRunStyle(Run openXmlElement, CommonGost.StyleTypeEnum typeStyle)
        {
            if (openXmlElement == null) return;

            var p = new OpenXmlGenericRepositoryRun<Run>(openXmlElement);
            foreach ( var run in openXmlElement.Elements<RunProperties>())
            {
                if (run.Bold != null && (run.Bold.Val == null || run.Bold.Val == true))
                {
                    p.ClearAll();
                    p.Bold(true);
                    if (_gostRepository.GetFontSize(typeStyle) != null) p.FontSize(_gostRepository.GetFontSize(typeStyle).SafeToInt(-1));
                    if (_gostRepository.GetColor(typeStyle) != null) p.Color(_gostRepository.GetColor(typeStyle));
                    //if (_gostRepository.GetBold(typeStyle) != null) p.Bold(_gostRepository.GetBold(typeStyle).nvl());
                    if (_gostRepository.GetFont(typeStyle) != null) p.RunFonts(_gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle));
                }
                else if(run.Italic != null &&(run.Italic.Val == null || run.Italic.Val == true))
                {
                    p.ClearAll();
                    p.Italic(true);
                    if (_gostRepository.GetFontSize(typeStyle) != null) p.FontSize(_gostRepository.GetFontSize(typeStyle).SafeToInt(-1));
                    if (_gostRepository.GetColor(typeStyle) != null) p.Color(_gostRepository.GetColor(typeStyle));
                    if (_gostRepository.GetBold(typeStyle) != null) p.Bold(_gostRepository.GetBold(typeStyle).nvl());
                    if (_gostRepository.GetFont(typeStyle) != null) p.RunFonts(_gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle)); 
                }
                else if(run.Underline != null)
                {
                    string uVal = run.Underline.Val;
                    p.ClearAll();
                    p.Underline(uVal);
                    if (_gostRepository.GetFontSize(typeStyle) != null) p.FontSize(_gostRepository.GetFontSize(typeStyle).SafeToInt(-1));
                    if (_gostRepository.GetColor(typeStyle) != null) p.Color(_gostRepository.GetColor(typeStyle));
                    if (_gostRepository.GetBold(typeStyle) != null) p.Bold(_gostRepository.GetBold(typeStyle).nvl());
                    if (_gostRepository.GetFont(typeStyle) != null) p.RunFonts(_gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle));
                }
                else
                {
                    p.ClearAll();
                    if (_gostRepository.GetFontSize(typeStyle) != null) p.FontSize(_gostRepository.GetFontSize(typeStyle).SafeToInt(-1));
                    if (_gostRepository.GetColor(typeStyle) != null) p.Color(_gostRepository.GetColor(typeStyle));
                    if (_gostRepository.GetBold(typeStyle) != null) p.Bold(_gostRepository.GetBold(typeStyle).nvl());
                    if (_gostRepository.GetFont(typeStyle) != null) p.RunFonts(_gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle));

                }
            }
            if (_gostRepository.GetItalic(typeStyle) != null) p.Italic(_gostRepository.GetItalic(typeStyle).nvl());
            if (_gostRepository.GetUnderline(typeStyle) != null) p.Underline(_gostRepository.GetUnderline(typeStyle));
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
                        if (itemRun is Run run) SetRunStyle(run, CommonGost.StyleTypeEnum.Image);
                    }
                 
                    SetParagraphStyle(item, CommonGost.StyleTypeEnum.Image, true);
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
                 
                    SetParagraphStyle(item, CommonGost.StyleTypeEnum.Image, true);
                }
            }
        }


        private void SetParagraphStyle(Paragraph para, CommonGost.StyleTypeEnum typeStyle, bool clearAllProperties)
        {

            if (para == null) return;

            var p = new OpenXmlGenericRepositoryParagraph<Paragraph>(para);
            if (clearAllProperties) p.ClearAll();

            if (_gostRepository.GetFontSize(typeStyle) != null) p.FontSize(_gostRepository.GetFontSize(typeStyle).SafeToInt(-1));
            if (_gostRepository.GetColor(typeStyle) != null) p.Color(_gostRepository.GetColor(typeStyle));
            if (_gostRepository.GetBold(typeStyle) != null) p.Bold(_gostRepository.GetBold(typeStyle).nvl());
            if (_gostRepository.GetFont(typeStyle) != null) p.RunFonts(_gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle));
            if (_gostRepository.GetAlignment(typeStyle) != null) p.Justification(_gostRepository.GetAlignment(typeStyle));
            p.SpacingBetweenLines(_gostRepository.GetLineSpacing(typeStyle).nvl(), _gostRepository.GetBeforeSpacing(typeStyle).nvl(), _gostRepository.GetAfterSpacing(typeStyle).nvl());
        }


        private string GetPathToSaveObj()
        {
            return Path.GetFullPath(Path.Combine(PathHelper.TmpDirectory(), $"{Guid.NewGuid().ToString()}{ExtensionDoc}"));
        }
    }

}
