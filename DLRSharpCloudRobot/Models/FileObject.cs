using System.IO;

namespace DLRSharpCloudRobot.Models
{
    public class FileObject
    {
        public string FullPath { get; set; }

        public string FileName
        {
            get { return Path.GetFileName(FullPath); }
        }

        public string Name
        {
            get { return Path.GetFileNameWithoutExtension(FullPath); }
        }

        public bool IsSelected { get; set; }
        
        public FileObject(string path)
        {
            FullPath = path;
        }
    }
}
