using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlFind.Utils
{
    public class FilesFinder
    {
        private string InitDir { get; set; }
        private FilesFinder(string InitDir)
        {
            this.InitDir = InitDir;
        }

        public static FilesFinder FindFrom(string initDir)
        {
            return new FilesFinder(initDir);
        }

        public IEnumerable<string> WithExtension(string ext)
        {
            if (!ext.StartsWith(".")) { ext = "." + ext; }
            return FilesFrom(InitDir).Where(filename => filename.EndsWith(ext));
        }

        public IEnumerable<string> AllFiles()
        {
            return FilesFrom(InitDir);
        }

        private static IEnumerable<string> FilesFrom(string sDir)
        {
            List<string> files = new List<string>();
            try
            {
                foreach (string f in Directory.GetFiles(sDir))
                {
                    files.Add(f);
                }
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    files = files.Concat(FilesFrom(d)).ToList();
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
            return files;
        }

    }
}
