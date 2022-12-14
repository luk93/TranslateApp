using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace TranslateApp.Extensions
{
    public static class FileInfoExt
    {
        public static bool IsFileLocked(this FileInfo fileInfo, string filePath)
        {
            try
            {
                var stream = File.OpenRead(filePath);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }
    }
}
