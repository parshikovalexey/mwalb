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
using System.Collections.Generic;

namespace ProcessDocumentCore.Processing
{
    public class ProcessingOpenXml : IDocumentProcessing
    {
        private const string ExtensionDoc = ".docx";
        private GostGenericRepository<GostModel> _gostRepository;

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
                    var styles = wordDoc.MainDocumentPart.StyleDefinitionsPart.Styles;

                    //устанавливаем отступы для всего документа
                    SetPageMargin(body);
                    List<int> formattedNumID = new List<int>(); //Список с id отформатированных списков

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
                        //else
                        //{
                            //Определение заголовков
                            if (IsParagraphHeader(para, styles))
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
                        //}


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
                            SetNumberingProperties(para, wordDoc, formattedNumID);
                    }

                    CorrectImage(body, styles, _gostRepository.GetImageCaptionFormat(), _gostRepository.GetImageCaptionRule());

                    SetHeaderPartStyle(wordDoc);
                    SetFooterPartStyle(wordDoc);

                    SetTOCStyle(body, styles);

                    DocumentSettingsPart settingsPart = wordDoc.MainDocumentPart.DocumentSettingsPart;
                    UpdateFieldsOnOpen update = new UpdateFieldsOnOpen() { Val = true };
                    settingsPart.Settings.PrependChild(update);

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

        /// <summary>
        /// Возвращает true, если абзац является заголовком
        /// </summary>
        /// <param name="paragraph">Абзац для проверки</param>
        /// <param name="styles">Стили документа</param>
        private bool IsParagraphHeader(Paragraph paragraph, Styles styles)
        {
            if (paragraph?.ParagraphProperties?.OutlineLevel != null) return true;

            var style = GetParagraphStyle(paragraph, styles);
            string stylename = style != null ? style.StyleName.Val.ToString().ToLower() : null;

            if (stylename != null && stylename.Contains("heading")) return true;

            return false;
        }

        private void SetNumberingProperties(Paragraph para, WordprocessingDocument wordDoc, IList<int> ids)
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

                if (ids.Any(i => i == num.Val)) return; // Если список уже форматировался выходим
                else ids.Add(num.Val); //Иначе добавляем список в отформатированные

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

                foreach (var item in abs.ToList())
                {
                    if (item is AbstractNum an)
                    {
                        if (an.AbstractNumberId == v)
                        {
                            var levels = an.Where(e => e is Level);
                            bool isBulletList = false;
                            foreach (var itemLevel in levels)
                            {
                                if (itemLevel is Level level)
                                {
                                    if (level.LevelIndex == 0)
                                    {
                                        if (level.NumberingFormat.Val == NumberFormatValues.Bullet) isBulletList = true;
                                    }

                                    var numberingFormat = _gostRepository.GetNumberingFormat(level.LevelIndex, isBulletList);
                                    var levelText = _gostRepository.GetNumberingLevelText(level.LevelIndex, isBulletList);

                                    SetlevelIndentation(level, isBulletList);
                                    SetLevelJustification(level, isBulletList);

                                    if (numberingFormat != NumberFormatValues.None)
                                    {
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
                                            level.Append(prop);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void SetlevelIndentation(Level level, bool isBullet)
        {
            var paragraphProperties = level?.PreviousParagraphProperties;
            var indentation = paragraphProperties?.FirstOrDefault(p =>
                p.GetType() == typeof(Indentation));
            int multiplier = 567;
            if (indentation != null && indentation is Indentation levelIndentation)
            {
                levelIndentation.Left =
                    ((int)(_gostRepository.GetNumberingIndentationLeft(level.LevelIndex, isBullet) * multiplier)).ToString();
                levelIndentation.Hanging = ((int)(_gostRepository.GetNumberingHanging(isBullet) * multiplier)).ToString();
            }
        }

        private void SetLevelJustification(Level level, bool isBullet)
        {
            var justification = level.LevelJustification;
            if (justification != null)
            {
                justification.Val = _gostRepository.GetNumberingJustification(level.LevelIndex, isBullet);
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
            foreach (var run in openXmlElement.Elements<RunProperties>())
            {
                bool bold, italic;
                UnderlineValues underline;
                bold = (run.Bold != null && (run.Bold.Val == null || run.Bold.Val == true)) ? true : false;
                italic = (run.Italic != null && (run.Italic.Val == null || run.Italic.Val == true)) ? true : false;
                underline = (run.Underline != null && run.Underline.Val != null) ? run.Underline.Val.Value : UnderlineValues.None;

                p.ClearAll();
                if (typeStyle == CommonGost.StyleTypeEnum.GlobalText)
                {
                    p.Bold(bold);
                    p.Italic(italic);
                    p.Underline(underline.ToString());
                }
                else
                {
                    if (_gostRepository.GetBold(typeStyle) != null) p.Bold(_gostRepository.GetBold(typeStyle).nvl());
                    if (_gostRepository.GetItalic(typeStyle) != null) p.Italic(_gostRepository.GetItalic(typeStyle).nvl());
                    if (_gostRepository.GetUnderline(typeStyle) != null) p.Underline(_gostRepository.GetUnderline(typeStyle));
                }

                if (_gostRepository.GetFontSize(typeStyle) != null) p.FontSize(_gostRepository.GetFontSize(typeStyle).SafeToInt(-1));
                if (_gostRepository.GetColor(typeStyle) != null) p.Color(_gostRepository.GetColor(typeStyle));
                if (_gostRepository.GetFont(typeStyle) != null) p.RunFonts(_gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle), _gostRepository.GetFont(typeStyle));
            }

        }

        //private AbstractNum GetAbstractNumByNumId(int numid, Numbering numbering)
        //{
        //    int abstractid = -1;
        //    foreach (var numInst in numbering.Descendants<NumberingInstance>().ToList())
        //    {
        //        if (numInst.NumberID == numid) abstractid = numInst.AbstractNumId.Val;
        //    }

        //    if (abstractid != -1)
        //    {
        //        foreach (var abNum in numbering.Descendants<AbstractNum>().ToList())
        //        {
        //            if (abNum.AbstractNumberId == abstractid) return abNum;
        //        }
        //    }
        //    return null;
        //}

        private void CorrectImage(Body body, Styles styles, string format, int rule)
        {
            if (body == null) return;

            var isNextRunIsHeaderImg = false;

            int image_number = 1;

            foreach (var item in body.Elements<Paragraph>().ToList())
            {
                if (isNextRunIsHeaderImg)
                {
                    isNextRunIsHeaderImg = false;

                    if (!item.Any(r => r.GetType() == typeof(Run)))
                    {
                        item.Remove();
                        isNextRunIsHeaderImg = true;
                        continue;
                    }

                    Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                    Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };

                    //Переписываем весь текст из параграфа в одну переменную
                    foreach (var t in item.Descendants<Text>().ToList())
                    {
                        text1.Text += t.Text;
                    }

                    //Находим в переменной с текстом позицию номера и количество цифр
                    int pos = -1, count = -1;
                    for (int i = 0; i < text1.Text.Length; i++)
                    {
                        if (Char.IsNumber(text1.Text[i]))
                        {
                            if (pos == -1)
                            {
                                pos = i;
                                count = 1;
                            }
                            else count++;
                        }
                        if (text1.Text.Substring(0, i).ToLower().Contains("рисунок ") && pos == -1) pos = i;
                    }

                    //Удаляем номер
                    if (pos != -1)
                    {
                        if (count != -1)
                            text1.Text = text1.Text.Remove(pos, count);
                        text2.Text = text1.Text.Substring(pos);
                        text1.Text = text1.Text.Remove(pos);
                    }

                    string stylename = GetHeadlineStyleName(item, body, styles);
                    Run run1 = new Run(text1);
                    Run run2 = new Run(text2);

                    //Записываем в параграф
                    item.RemoveAllChildren<Run>();
                    item.Append(run1);

                    //
                    InsertImageNumber(item, stylename, format, rule);

                    item.Append(run2);

                    foreach (var r in item.Descendants<Run>().ToList())
                    {
                        SetRunStyle(r, CommonGost.StyleTypeEnum.ImageCaption);
                    }

                    SetParagraphStyle(item, CommonGost.StyleTypeEnum.Image, true);
                }

                var findDrawing = item.Any(f => f.ToList().Any(e => e is Drawing));
                if (findDrawing)
                {
                    isNextRunIsHeaderImg = true;

                    SetParagraphStyle(item, CommonGost.StyleTypeEnum.Image, true);
                    SetRunStyle(item.GetFirstChild<Run>(), CommonGost.StyleTypeEnum.Image);
                }
            }
        }

        private void InsertImageNumber(Paragraph item, string stylename, string format, int rule)
        {
            if (format.ToLower() != "section" && format.ToLower() != "simple") return;

            Run runbegin = new Run();
            FieldChar fieldCharBegin = new FieldChar() { FieldCharType = FieldCharValues.Begin, Dirty = true };
            runbegin.Append(fieldCharBegin);

            Run runseparate = new Run();
            FieldChar fieldCharSeparate = new FieldChar() { FieldCharType = FieldCharValues.Separate };
            runseparate.Append(fieldCharSeparate);

            Run runend = new Run();
            FieldChar fieldCharEnd = new FieldChar() { FieldCharType = FieldCharValues.End };
            runend.Append(fieldCharEnd);

            Run emptyrun = new Run();

            Run run1 = new Run();
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = $" STYLEREF \"{stylename}\" \\n ";
            run1.Append(fieldCode1);

            Run run2 = new Run();
            FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode2.Text = $" SEQ Рисунок \\* ARABIC \\s {rule} ";
            run2.Append(fieldCode2);
            if (format == "Section")
            {
                item.Append(runbegin.CloneNode(true));
                item.Append(run1.CloneNode(true));
                item.Append(runseparate.CloneNode(true));
                item.Append(emptyrun.CloneNode(true));
                item.Append(runend.CloneNode(true));

                item.Append(new Run(new Text(".")));
            }

            item.Append(runbegin.CloneNode(true));
            item.Append(run2.CloneNode(true));
            item.Append(runseparate.CloneNode(true));
            item.Append(emptyrun.CloneNode(true));
            item.Append(runend.CloneNode(true));
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
            p.Indentation(_gostRepository.GetFirstLineIndentation(typeStyle).nvl(), _gostRepository.GetLeftIndentation(typeStyle).nvl(), _gostRepository.GetRightIndentation(typeStyle).nvl());
        }

        private void SetHeaderPartStyle(WordprocessingDocument wDoc)
        {
            if (!wDoc.MainDocumentPart.HeaderParts.Any()) return;
            var paragrpahs = wDoc.MainDocumentPart.HeaderParts.FirstOrDefault().Header.Descendants<Paragraph>().ToList();
            foreach (var p in paragrpahs)
            {
                SetParagraphStyle(p, CommonGost.StyleTypeEnum.HeaderPart, true);
                foreach (var r in p.Descendants<Run>().ToList())
                {
                    SetRunStyle(r, CommonGost.StyleTypeEnum.HeaderPart);
                }
            }
        }

        private void SetFooterPartStyle(WordprocessingDocument wDoc)
        {
            if (!wDoc.MainDocumentPart.FooterParts.Any()) return;
            var paragrpahs = wDoc.MainDocumentPart.FooterParts.FirstOrDefault().Footer.Descendants<Paragraph>().ToList();
            foreach (var p in paragrpahs)
            {
                SetParagraphStyle(p, CommonGost.StyleTypeEnum.FooterPart, true);
                foreach (var r in p.Descendants<Run>().ToList())
                {
                    SetRunStyle(r, CommonGost.StyleTypeEnum.FooterPart);
                }
            }
        }

        /// <summary>
        /// Возвращает блок оглавления из тела документа
        /// </summary>
        /// <param name="body">Тело документа</param>
        /// <returns>Блок оглавления</returns>
        private SdtBlock GetTOC(Body body)
        {
            var blocks = body.Descendants<SdtBlock>();
            foreach (var block in blocks)
            {
                SdtContentDocPartObject docpartobj = (SdtContentDocPartObject)block.SdtProperties.FirstOrDefault(e => e.GetType() == typeof(SdtContentDocPartObject));
                if (docpartobj.DocPartGallery.Val == "Table of Contents") return block;
            }
            return null;
        }

        /// <summary>
        /// Возвращает стиль абзаца из списка стилей документа или null
        /// </summary>
        /// <param name="para">Абзац</param>
        /// <param name="styles">Стили документа</param>
        /// <returns>Стиль абзаца</returns>
        private Style GetParagraphStyle(Paragraph para, Styles styles)
        {
            if (para?.ParagraphProperties?.ParagraphStyleId == null) return null;

            var styleid = para?.ParagraphProperties?.ParagraphStyleId.Val;
            var style = styles.Descendants<Style>().FirstOrDefault(s => s.StyleId.Value == styleid.Value);

            return style ?? null;
        }

        private string GetHeadlineStyleName(Paragraph paragraph, Body body, Styles styles)
        {
            string stylename = "0";
            foreach (var p in body.Descendants<Paragraph>().ToList())
            {
                var style = GetParagraphStyle(p, styles);
                if (style != null && IsParagraphHeader(p, styles))
                    stylename = style.StyleName.Val;
                if (stylename.ToLower().Contains("heading")) stylename = stylename.ToLower().Replace("heading", "Заголовок");
                if (stylename == "List Paragraph") stylename = "Абзац списка";
                if (p == paragraph) return stylename;
            }
            return "0";
        }

        /// <summary>
        /// Возвращает уровень пункта оглавления в документе по ссылке
        /// </summary>
        /// <param name="body">Тело документа</param>
        /// <param name="styles">Стили документа</param>
        /// <param name="anchor">Ссылка на абзац с заголовком</param>
        /// <returns>Уровень пункта оглавления</returns>
        private int GetHeaderLevelByAnchor(Body body, Styles styles, string anchor)
        {
            int level = 1;
            if (anchor == null) return level - 1;
            foreach (var bookmark in body.Descendants<BookmarkStart>())
            {
                if (bookmark.Name == anchor)
                {
                    var paragraph = (Paragraph)bookmark.Parent;
                    var styleId = paragraph.ParagraphProperties.ParagraphStyleId.Val.ToString();

                    var name = GetParagraphStyle(paragraph, styles).StyleName.Val.ToString();
                    if (name.Contains("heading")) int.TryParse(name.Substring(name.Length - 1), out level);

                }
            }
            return level - 1;
        }

        /// <summary>
        /// Задает форматирование оглавления (Table of Contents)
        /// </summary>
        private void SetTOCStyle(Body body, Styles styles)
        {
            var TOC = GetTOC(body);

            if (TOC == null) return;

            foreach (var para in TOC.Descendants<Paragraph>())
            {
                if (para.ParagraphProperties == null) para.ParagraphProperties = new ParagraphProperties();
                var style = GetParagraphStyle(para, styles);

                //Ищем ссылку на заголовок
                string anchor = null;
                if (para.Descendants<Hyperlink>().Any())
                {
                    var hyperlink = para.Descendants<Hyperlink>().FirstOrDefault();
                    anchor = hyperlink.Anchor;
                }
                else if (para.Descendants<FieldCode>().Any(f => f.Text.ToLower().Contains("hyperlink")))
                {
                    var field = para.Descendants<FieldCode>().FirstOrDefault(f => f.Text.ToLower().Contains("hyperlink"));
                    anchor = field.Text.Substring(field.Text.IndexOf('\"') + 1, field.Text.Length - field.Text.IndexOf('\"') - 2);
                }

                //Устанавливаем расположение табов
                if (para.ParagraphProperties.Tabs != null || style?.StyleParagraphProperties?.Tabs != null)
                {
                    Tabs tabs = para.ParagraphProperties.Tabs ?? style.StyleParagraphProperties.Tabs;
                    tabs.Descendants<TabStop>().Last().Remove();

                    if (tabs.Count() > 0)
                    {
                        foreach (var tab in tabs.Descendants<TabStop>().ToList())
                        {
                            int ind = 0;
                            if (para.ParagraphProperties.Indentation?.Left != null)
                            {
                                ind = Int32.Parse(para.ParagraphProperties.Indentation.Left);
                            }
                            else if (style.StyleParagraphProperties?.Indentation?.Left != null)
                            {
                                ind = Int32.Parse(style.StyleParagraphProperties?.Indentation?.Left);
                            }

                            int position = tab.Position - ind;
                            tab.Position = _gostRepository.GetTOCIndentationLeft(GetHeaderLevelByAnchor(body, styles, anchor)) + position;
                        }
                    }

                    TabStop pagenum = new TabStop()
                    {
                        Leader = _gostRepository.GetTOCTabLeader(),
                        Position = _gostRepository.GetTOCTabPosition(),
                        Val = TabStopValues.Right
                    };
                    tabs.Append(pagenum);
                }
                else
                {
                    Tabs tabs = new Tabs();
                    tabs.Append(new TabStop()
                    {
                        Leader = _gostRepository.GetTOCTabLeader(),
                        Position = _gostRepository.GetTOCTabPosition(),
                        Val = TabStopValues.Right
                    });
                    para.ParagraphProperties.Append(tabs);
                }

                //Устанавливаем отступы
                para.ParagraphProperties.Indentation = new Indentation()
                {
                    Left = _gostRepository.GetTOCIndentationLeft(GetHeaderLevelByAnchor(body, styles, anchor)).ToString(),
                    FirstLine = _gostRepository.GetTOCFirstIndentation().ToString()
                };
            }
        }

        private string GetPathToSaveObj()
        {
            return Path.GetFullPath(Path.Combine(PathHelper.TmpDirectory(), $"{Guid.NewGuid().ToString()}{ExtensionDoc}"));
        }
    }

}
