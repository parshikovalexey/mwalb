using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXmlHelperLibrary

{
    public class OpenXmlGenericRepositoryParagraph<T> where T : Paragraph
    {
        private readonly Paragraph _paragraph;
        public OpenXmlGenericRepositoryParagraph(Paragraph paragraph)
        {
            _paragraph = paragraph;
            if (paragraph == null) return;
            if (_paragraph.ParagraphProperties == null) _paragraph.ParagraphProperties = new ParagraphProperties();
            if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null) _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new ParagraphMarkRunProperties();
        }
        public void ClearAll()
        {
            _paragraph?.ParagraphProperties?.Remove();
            _paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.Remove();
            if (_paragraph != null)
            {
                _paragraph.ParagraphProperties = new ParagraphProperties
                {
                    ParagraphMarkRunProperties = new ParagraphMarkRunProperties()
                };
            }
        }

        public void FontSize(int size)
        {
            ClearSingleStyleFromMarkRunProperties(typeof(FontSize));
            var newStyle = new FontSize { Val = (size * 2).ToString() };
            AddStyleToMarkRunProperties(newStyle);
        }

        public void Color(string color)
        {

            if (string.IsNullOrEmpty(color)) return;
          
            ClearSingleStyleFromMarkRunProperties(typeof(Color));
            var _color = new Color { Val = color };
            AddStyleToMarkRunProperties(_color);
            //_paragraph.ParagraphProperties.ParagraphMarkRunProperties.Append(_color);
            //var t = _paragraph.ParagraphProperties.ParagraphMarkRunProperties;
            //if (t == null) return;
            //var newStyle = new Color { Val = color };
            ////AddStyleToProperties(newStyle);
            //t.Append(newStyle);


            //if (string.IsNullOrEmpty(color)) return;
            //ClearSingleStyleFromMarkRunProperties(typeof(Color));
            //var newStyle = new Color { Val = color };
            //
        }

        public void Bold(bool bold)
        {
            ClearSingleStyleFromMarkRunProperties(typeof(Bold));
            var newStyle = new Bold { Val = bold };
            AddStyleToMarkRunProperties(newStyle);
        }
        public void RunFonts(string ascii, string highAnsi)
        {
            ClearSingleStyleFromMarkRunProperties(typeof(RunFonts));
            var newStyle = new RunFonts();
            if (!string.IsNullOrEmpty(ascii)) newStyle.Ascii = ascii;
            if (!string.IsNullOrEmpty(highAnsi)) newStyle.Ascii = highAnsi;
            AddStyleToMarkRunProperties(newStyle);
        }

        public void SpacingBetweenLines(float lineSpacing, float beforeSpacing, float afterSpacing)
        {
            ClearSingleStyleFromMarkRunProperties(typeof(SpacingBetweenLines));
            var newStyle = new SpacingBetweenLines() //Интервалы между строками и абзацами
            {
                Line = (lineSpacing * 240).ToString(),
                Before = (beforeSpacing * 20).ToString(),
                After = (afterSpacing * 20).ToString(),
                BeforeAutoSpacing = new OnOffValue(false),
                AfterAutoSpacing = new OnOffValue(false)
            };
            AddStyleToProperties(newStyle);
        }

        public void Indentation(float firstLineIndentation, float leftIndentation, float rightIndentation)
        {
            ClearSingleStyleFromMarkRunProperties(typeof(Indentation));
            var newStyle = new Indentation() //Отступы
            {
                FirstLine = ((int)(firstLineIndentation * 567 - leftIndentation * 567)).ToString(),
                Left = ((int)(leftIndentation * 567)).ToString(),
                Right = ((int)(rightIndentation * 567)).ToString(),
            };
            AddStyleToProperties(newStyle);
        }

        public void Justification(string justification) //todo Неизведанная магия! при отправке в метод AddStyleToProperties выравнивание не добавляется к стили
        {
            if (string.IsNullOrEmpty(justification)) return;
            ClearSingleStyleFromProperties(typeof(Justification));

            //var t = _paragraph.ParagraphProperties;
            //if (t == null) return;
            var newStyle = new Justification { Val = justification.GetJustificationByString() };
            AddStyleToProperties(newStyle);
            //AddStyleToProperties(newStyle);
            //t.Append(newStyle);


            //var newStyle = new Justification() { Val = justification.GetJustificationByString() };
            //AddStyleToProperties(newStyle);
        }

        public void LineSpacing(string line)
        {
            var spacing = new SpacingBetweenLines()
            {
                Line = line
            };

            //AddStyleToProperties(spacing);
            if (_paragraph.ParagraphProperties.SpacingBetweenLines == null)
                _paragraph.ParagraphProperties.Append(spacing);
            else _paragraph.ParagraphProperties.SpacingBetweenLines = spacing;
        }

        public void BeforeSpacing(string before)
        {
            var spacing = new SpacingBetweenLines()
            {
                Before = before
            };

            //AddStyleToProperties(spacing);
            if (_paragraph.ParagraphProperties.SpacingBetweenLines == null)
                _paragraph.ParagraphProperties.Append(spacing);
            else _paragraph.ParagraphProperties.SpacingBetweenLines = spacing;
        }

        public void AfterSpacing(string after)
        {
            var spacing = new SpacingBetweenLines()
            {
                After = after
            };

            //AddStyleToProperties(spacing);
            if (_paragraph.ParagraphProperties.SpacingBetweenLines == null)
                _paragraph.ParagraphProperties.Append(spacing);
            else _paragraph.ParagraphProperties.SpacingBetweenLines = spacing;
        }

        public void FirstLineIndent(string firstline)
        {
            var indent = new Indentation()
            {
                FirstLine = firstline
            };

            //AddStyleToProperties(spacing);
            if (_paragraph.ParagraphProperties.Indentation == null)
                _paragraph.ParagraphProperties.Append(indent);
            else _paragraph.ParagraphProperties.Indentation = indent;
        }
        public void LeftIndent(string left)
        {
            var indent = new Indentation()
            {
                Left = left
            };

            //AddStyleToProperties(spacing);
            if (_paragraph.ParagraphProperties.Indentation == null)
                _paragraph.ParagraphProperties.Append(indent);
            else _paragraph.ParagraphProperties.Indentation = indent;
        }
        public void RightIndent(string right)
        {
            var indent = new Indentation()
            {
                Right = right
            };

            //AddStyleToProperties(spacing);
            if (_paragraph.ParagraphProperties.Indentation == null)
                _paragraph.ParagraphProperties.Append(indent);
            else _paragraph.ParagraphProperties.Indentation = indent;
        }

        private void AddStyleToMarkRunProperties(OpenXmlElement obj)
        {
            try
            {
                _paragraph.ParagraphProperties?.ParagraphMarkRunProperties.Append(obj);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

        }
        private void AddStyleToProperties(OpenXmlElement obj)
        {
            try
            {
                //var t = _paragraph.ParagraphProperties;
                //if (t == null) return;
                //Justification newStyle = obj;
                _paragraph.ParagraphProperties.Append(obj);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

        }
        private void ClearSingleStyleFromMarkRunProperties(object obj)
        {
            try
            {
                _paragraph.ParagraphProperties.ParagraphMarkRunProperties.ChildElements.FirstOrDefault(x =>
                    obj != null && x.GetType() == obj)?.Remove();

                //_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.FirstOrDefault(element =>
                //    element.GetType() == obj.GetType())?.Remove();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

        }

        private void ClearSingleStyleFromProperties(object obj)
        {
            try
            {
                _paragraph.ParagraphProperties.ChildElements.FirstOrDefault(x =>
                    obj != null && x.GetType() == obj)?.Remove();
                //_paragraph?.ParagraphProperties?.FirstOrDefault(element =>
                //    element.GetType() == obj.GetType())?.Remove();

                //_paragraph?.ParagraphProperties?.FirstOrDefault(element =>element)?.Remove();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

        }
    }
}
