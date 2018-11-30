using CommonLibrary;
using CommonLibrary.Interfaces;
using Newtonsoft.Json;
using StandardsLibrary.Simple;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HelperLibrary;

namespace StandardsLibrary
{
    public class GostService : IErrorInterface
    {
        List<SimpleGost> Gosts = new List<SimpleGost>();
        private string _folderGostFile = Path.GetFullPath(Path.Combine(PathHelper.GetBaseDirectory, "gost"));

        public void SetFolderGostFile(string filePath)
        {
            _folderGostFile = filePath;
        }
        public List<SimpleHeaderGost> GetGostList()
        {
            List<SimpleHeaderGost> result = Gosts.Select(x => new SimpleHeaderGost()
                {Header = x.Header.nvl(), GuidGost = x.GuidGost}).ToList();
            return result;
        }

        public GostModel GetGostModel(Guid guid)
        {
            var g = Gosts.FirstOrDefault(gost => gost.GuidGost == guid)?.Gost?? new GostModel();
            return g;
        }

        public void LoadGostFromFile()
        {
            Gosts = new List<SimpleGost>();
            if (PathHelper.DirectoryExists(_folderGostFile, false))
            {
                ICollection<string> files = PathHelper.GetAllFiles(_folderGostFile, "json");
                foreach (var item in files)
                {
                    try
                    {
                        var gostTmp = JsonConvert.DeserializeObject<GostModel>(File.ReadAllText(item));
                        Gosts.Add(new SimpleGost() { Header = gostTmp.Header, FilePath = item, Gost = gostTmp });
                    }
                    catch (Exception e)
                    {
                        OnErrors(new ResultExecute().OnError(e.Message));
                    }
                }
            }
            else
            {
                OnErrors(new ResultExecute().OnError($"Путь {_folderGostFile} недоступен"));
            }
        }

        public void SaveDemoGost()
        {
            var g = new GostModel()
            {
                Header = "Тестовый гост",
                GlobalText = new SimpleStyle()
                {
                    Font = "Times New Roman",
                    FontSize = 15,
                    Alignment = "Both",
                    Bold = false,
                    AfterSpacing = 1,
                    BeforeSpacing = 1,
                    FirstLineIndentation = 1.5f,
                    LeftIndentation = 0,
                    LineSpacing = 1.5f,
                    MarginBottom = 0,
                    MarginLeft = 0,
                    MarginRight = 0,
                    MarginTop = 0,
                    RightIndentation = 0
                },
                Headline = new SimpleStyle()
                {
                    Font = "Times New Roman",
                    FontSize = 30,
                    Alignment = "Center",
                    Bold = true,
                    AfterSpacing = 1,
                    BeforeSpacing = 1,
                    FirstLineIndentation = 1.5f,
                    LeftIndentation = 0,
                    LineSpacing = 1.5f,
                    MarginBottom = 0,
                    MarginLeft = 0,
                    MarginRight = 0,
                    MarginTop = 0,
                    RightIndentation = 0,
                    Color = "365F91"
                },
                Image = new SimpleStyle()
                {
                    Font = "Times New Roman",
                    FontSize = 30,
                    Alignment = "Center",
                    Bold = true,
                    AfterSpacing = 1,
                    BeforeSpacing = 1,
                    FirstLineIndentation = 1.5f,
                    LeftIndentation = 0,
                    LineSpacing = 1.5f,
                    MarginBottom = 0,
                    MarginLeft = 0,
                    MarginRight = 0,
                    MarginTop = 0,
                    RightIndentation = 0,
                },
            };
            string json = JsonConvert.SerializeObject(g);
            PathHelper.DirectoryExists(_folderGostFile, true);
            //write string to file
            System.IO.File.WriteAllText(Path.Combine(_folderGostFile, "demo.json"), json);
        }

        public event Action<ResultExecute> Errors;

        protected virtual void OnErrors(ResultExecute obj)
        {
            Errors?.Invoke(obj);
        }
    }
}
