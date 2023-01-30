using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TranslateApp.Extensions;

namespace TranslateApp.Tools
{
    public static class ExcelOperations
    {
        public static async Task<DataTable> FromExcelFileToDataTable(FileInfo file, bool hasHeaderRow = true)
        {
            var package = new ExcelPackage(file);
            await package.LoadAsync(file);
            var ws = package.Workbook.Worksheets[0];
            var dtResult = new DataTable();
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                dtResult.Columns.Add(hasHeaderRow ?
                    firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            var startRow = hasHeaderRow ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = dtResult.NewRow();
                foreach (var cell in wsRow) row[cell.Start.Column - 1] = cell.Text;
                dtResult.Rows.Add(row);
            }
            return dtResult;
        }
        public static ExcelPackage? CreateExcelFile(string path)
        {
            var file = new FileInfo(path);
            if (file.Exists)
            {
                try
                {
                    file.Delete();
                }
                catch
                {
                    return null;
                }
            }
            return new ExcelPackage(file);
        }
        public static ExcelPackage? DuplicateExcelFile(FileInfo file, string exportFolderPath, string nameExtenstion)
        {
            string newPath = Path.Combine(exportFolderPath, file.Name.Replace(".", $"{nameExtenstion}."));
            File.Copy(file.FullName, newPath, true);
            FileInfo newFile = new FileInfo(newPath);
            return new ExcelPackage(newFile);
        }
        public static async Task SaveExcelFile(ExcelPackage excelPackage)
        {
            await excelPackage.SaveAsync();
            excelPackage.Dispose();
        }
    }

}
