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
            ClearSingleStyleFromMarkRunProperties(typeof(Color));
            var newStyle = new Color { Val = color };
            AddStyleToMarkRunProperties(newStyle);
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

        private void AddStyleToMarkRunProperties(IEnumerable<OpenXmlElement> obj)
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
