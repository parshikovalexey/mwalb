using System;
using System.IO;

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



        public static string TmpDirectory()
        {
            var tmp = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp"));
            return DirectoryExists(tmp, true) ? tmp : String.Empty;
        }

    }
}
