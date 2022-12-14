using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using TranslateApp.Data;

namespace TranslateApp.Tools
{
    public static class DataTableOperations
    {
        public static List<TextToTranslate> GetTextList(this DataTable dataTable, WSData wSData) 
        {
            List<TextToTranslate> listResult = new();
            //Get only non empty texts
            foreach (DataRow dr in dataTable.Rows)
            {
                if (!string.IsNullOrEmpty(dr[wSData.SrcColumn].ToString()))
                {
                    string? srcText = dr[wSData.SrcColumn].ToString();
                    string? id = dr[wSData.IDColumn].ToString();
                    int index = dataTable.Rows.IndexOf(dr);
                    if(srcText != null && id != null)
                    {
                        TextToTranslate obj = new(id, srcText, index);
                        listResult.Add(obj);
                    }
                }
            }
            return listResult;
        }
        public static void CheckHeaders(this DataTable dataTable, WSData wSData)
        {
            foreach (DataColumn column in dataTable.Columns)
            {
                if(column.ColumnName.Contains("ID")) wSData.IDColumn = column.Ordinal;
                if(column.ColumnName.Contains(wSData.SrcLangCode)) wSData.SrcColumn = column.Ordinal;
                if(column.ColumnName.Contains(wSData.TrgLangCode)) wSData.TrgColumn = column.Ordinal;
            }
        }
    }
}
