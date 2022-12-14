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
        public static List<TextToTranslate> GetTextList(this DataTable dataTable) 
        {
            List<TextToTranslate> listResult = new();
            foreach (DataRow dr in dataTable.Rows) { }
            return listResult;
        }
        public static string CheckHeaders(this DataTable dataTable, WSData wSData)
        {
            string strResult = string.Empty;
            foreach (DataColumn column in dataTable.Columns)
            {
                //ID:
                if (column.ColumnName.Contains("ID")) wSData.IDColumn = column.Ordinal;
                if (column.ColumnName.Contains("ID")) wSData.IDColumn = column.Ordinal;

            }
            return strResult;
        }
    }
}
