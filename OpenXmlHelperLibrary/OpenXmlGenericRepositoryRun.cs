using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXmlHelperLibrary

{
    public class OpenXmlGenericRepositoryRun<T> where T : Run
    {
        private readonly Run _run;

        public OpenXmlGenericRepositoryRun(Run run)
        {
            _run = run;
            if (run == null) return;
            if (_run.RunProperties == null) _run.RunProperties = new RunProperties();
        }
        public void ClearAll()
        {
            _run?.RunProperties?.Remove();
            if (_run != null)
            {
                _run.RunProperties = new RunProperties() { };
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
            var newColor = new Color { Val = color };
            _run?.RunProperties.Append(newColor);

            //ClearSingleStyleFromMarkRunProperties(typeof(Color));
            //var newStyle = new Color { Val = color };
            //AddStyleToMarkRunProperties(newStyle);
        }

        public void Bold(bool bold)
        {
            if (bold) {
                ClearSingleStyleFromMarkRunProperties(typeof(Bold));
                var newStyle = new Bold { Val = bold };
                AddStyleToMarkRunProperties(newStyle);
            }
            else
            {
                ClearSingleStyleFromMarkRunProperties(typeof(Bold));
            }
            
        }

        public void Italic(bool bold)
        {
            ClearSingleStyleFromMarkRunProperties(typeof(Italic));
            var newStyle = new Italic { Val = bold };
            AddStyleToMarkRunProperties(newStyle);
        }

        public void Underline(string uVal)
        {
            ClearSingleStyleFromMarkRunProperties(typeof(Underline));
            var newStyle = new Underline { Val = uVal.GetUnderlineByString() };
            AddStyleToMarkRunProperties(newStyle);
        }

        /// <summary>
        /// Задаем шрифт для текста
        /// </summary>
        /// <param name="ascii">для символов в диапазоне Unicode (U + 0000-U + 007F)</param>
        /// <param name="highAnsi">для символов в сложном наборе Unicode, например. Арабский текст.</param>
        /// <param name="complexScript">для символов в диапазоне Unicode, который не относится к одной из других категорий.</param>
        public void RunFonts(string ascii, string highAnsi, string complexScript)
        {
            ClearSingleStyleFromMarkRunProperties(typeof(RunFonts));
            var newStyle = new RunFonts();
            if (!string.IsNullOrEmpty(ascii)) newStyle.Ascii = ascii;
            if (!string.IsNullOrEmpty(highAnsi)) newStyle.HighAnsi = highAnsi;
            if (!string.IsNullOrEmpty(complexScript)) newStyle.ComplexScript = complexScript;
            AddStyleToMarkRunProperties(newStyle);
        }

        private void AddStyleToMarkRunProperties(OpenXmlElement obj)
        {
            try
            {
                _run.RunProperties?.Append(obj);
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
                _run.RunProperties.ChildElements.FirstOrDefault(x => obj != null && x.GetType() == obj)?.Remove();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

        }
    }
}
