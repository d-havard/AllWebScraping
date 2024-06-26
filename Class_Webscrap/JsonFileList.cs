using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Class_Webscrap
{
    public class JsonFileList
    {
        List<string> Pathfiles = new List<string>();

        /// <summary>
        /// Add a file path to the List "Pathfiles"
        /// </summary>
        /// <param name="path"></param>
        public void AddFileToList(string path)
        {
            Pathfiles.Add(path);
        }
        /// <summary>
        /// Return the List of file path
        /// </summary>
        /// <returns>A List<string> of files path</returns>
        public List<string> GetPathFile()
        {
            return Pathfiles;
        }
    }
}
