using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using TranslateApp.Data;

namespace TranslateApp.Extensions
{
    public static class DataTableExt
    {
        public static List<TextToTranslate> GetTextList(this DataTable dataTable, WSData wSData)
        {
            List<TextToTranslate> listResult = new();
            //Get only non empty texts
            foreach (DataRow dr in dataTable.Rows)
            {
                if (string.IsNullOrEmpty(dr[wSData.SrcColumn].ToString())) continue;
                string? srcText = dr[wSData.SrcColumn].ToString();
                string? id = dr[wSData.IdColumn].ToString();
                int index = dataTable.Rows.IndexOf(dr);
                if (srcText == null || id == null) continue;
                TextToTranslate obj = new(id, srcText, index);
                listResult.Add(obj);
            }
            return listResult;
        }
        public static void UpdateWithTextList(this DataTable dataTable, List<TextToTranslate> textList, WSData wSData)
        {
            foreach (DataRow dr in dataTable.Rows)
            {
                if (!string.IsNullOrEmpty(dr[wSData.SrcColumn].ToString()))
                {
                    dr[wSData.TrgColumn] = textList.Find(x => x.Row == dataTable.Rows.IndexOf(dr))?.TargetText ?? string.Empty;
                }
            }
        }

        public static void CheckHeaders(this DataTable dataTable, WSData wSData)
        {
            foreach (DataColumn column in dataTable.Columns)
            {
                if (column.ColumnName.Contains("ID")) wSData.IdColumn = column.Ordinal;
                if (column.ColumnName.Contains(wSData.SrcLangCode)) wSData.SrcColumn = column.Ordinal;
                if (column.ColumnName.Contains(wSData.TrgLangCode)) wSData.TrgColumn = column.Ordinal;
            }
        }
    }
}
