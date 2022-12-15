using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using TranslateApp.Data;
using TranslateApp.Tools;

namespace TranslateApp.Extensions
{
    public static class TextlistExt
    {
        public static List<TextToTranslate> GetListWithoutDuplicatedSource(this List<TextToTranslate> textList) 
        {
            return textList.GroupBy(x => x.SourceText).Select(y => y.First()).ToList();
        }
        private static void Translate(this List<TextToTranslate> textList, string srcLangCode, string trgLangCode)
        {
            foreach(var textToTranslate in textList)
            {
                textToTranslate.TargetText = Translators.TranslateTextWithoutApi(textToTranslate.SourceText, srcLangCode, trgLangCode);
            }
        }
        public static Task TranslateAsync(this List<TextToTranslate> textList, string srcLangCode, string trgLangCode)
        {
            return Task.Run(() => Translate(textList, srcLangCode, trgLangCode));
        }
        public static void FillListWithTranslationsList(this List<TextToTranslate> wholeList, List<TextToTranslate> shortList)
        {
            wholeList.ForEach(x =>
            {

                x.TargetText = shortList.Find(y => y.SourceText == x.SourceText)?.TargetText ?? string.Empty;
            });
        }
    }
}
