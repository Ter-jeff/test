using System.Collections.Generic;

namespace FileDiffLib
{
    public class FileComparer
    {
        public List<string> AddItems = [];
        public List<FileDiff> DiffItems = [];
        public List<string> MissingItems = [];
        public string Output { get; internal set; } = string.Empty;
        public string Expected { get; internal set; } = string.Empty;
    }
}
