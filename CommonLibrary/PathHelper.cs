using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CommonLibrary
{
    public static class PathHelper
    {
        public static bool DirectoryExists(string sPath, bool create)
        {
            var exist = Directory.Exists(sPath);
            if (create && !exist)
            {
                Directory.CreateDirectory(sPath);
                return true;
            }
            return exist;
        }

        public static ICollection<string> GetAllFiles(string rootDirectory, string fileExtension)
        {
            DirectoryInfo directory = new DirectoryInfo(rootDirectory);//Assuming Test is your Folder
            var files = directory.GetFiles($"*.{fileExtension}"); //Getting Text files
            return files.Select(f => f.Name).ToList();
        }

        public static void FileDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
            }

        }

        public static bool FileExists(string sPath, bool create)
        {
            var exist = File.Exists(sPath);
            if (create && !exist)
            {
                File.Create(sPath);
                return true;
            }
            return exist;

        }

        public static void ClearTmpDirectory()
        {
            try
            {
                string[] dirs = Directory.GetFiles(TmpDirectory());
                foreach (var item in dirs)
                {
                    FileDelete(item);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }


        }

        public static string GetBaseDirectory => AppDomain.CurrentDomain.BaseDirectory;

        public static string TmpDirectory()
        {
            var tmp = Path.GetFullPath(Path.Combine(GetBaseDirectory, "tmp"));
            return DirectoryExists(tmp, true) ? tmp : String.Empty;
        }

    }
}
