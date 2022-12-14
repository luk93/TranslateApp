using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TranslateApp.Data
{
    public class TextToTranslate
    {
        public string Id { get; set; }
        public string SourceText { get; set; }
        public string TargetText { get; set; }
        public int Row { get; set; }

        public TextToTranslate(string id, string srcText)
        {
            Id = id;
            SourceText = srcText;
            TargetText = string.Empty;
            Row = 0;
        }
    }
}
