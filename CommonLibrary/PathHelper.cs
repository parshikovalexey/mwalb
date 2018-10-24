using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary
{
  public  static class PathHelper
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

    }
}
