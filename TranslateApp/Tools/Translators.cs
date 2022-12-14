using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TranslateApp.Tools
{
    internal class Translators
    {
        public static string TranslateTextWithoutApi(string inputText, string langSource, string langTarget)
        {
            string result = "";
            var url = $"https://translate.google.com/m?hl={langSource}&sl={langSource}&tl={langTarget}&hl=en&ei=UTF-8&q={System.Web.HttpUtility.UrlEncode(inputText)}";
            HtmlWeb web = new()
            {
                OverrideEncoding = Encoding.UTF8
            };
            var htmlDoc = web.Load(url);
            var node = htmlDoc.DocumentNode.SelectSingleNode("//div");
            foreach (HtmlNode childNode in node.ChildNodes)
            {
                if (string.Equals(childNode.Attributes["class"].Value, "result-container"))
                {
                    //invisible chars
                    string correctedResult = System.Text.RegularExpressions.Regex.Replace(childNode.InnerText, @"[^\u0000-\u007F]", string.Empty);
                    //additional chars correction
                    correctedResult = correctedResult.Replace("&quot", "\"");
                    result = correctedResult.Replace("&gt", ">");
                    return result;
                }
            }
            return result;
        }
    }
}
