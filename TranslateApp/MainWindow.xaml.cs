using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Data;
using TranslateApp.Tools;
using TranslateApp.Extensions;
using TranslateApp.Data;
using System.Drawing;
using Brushes = System.Windows.Media.Brushes;

namespace TranslateApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string expFolderPath = string.Empty;
        public static FileInfo? textFile_g = null!;
        public static List<TextToTranslate>? textToTranslateList;
        public static DataTable textDataTable_g = null!;
        public static WSData wsData_g = null!;
        public static List<string> langCodeList_g = null!;
        public MainWindow()
        {
            wsData_g = new();
            langCodeList_g = LangCodes.CreateLanguageCodes();
            InitializeComponent();
            textToTranslateList = new();
            textDataTable_g = new();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        #region UI Handlers
        private void B_SelectExpFolder_Click(object sender, RoutedEventArgs e)
        {
            DisableButtonAndChangeCursor(sender);
            var openFolderDialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            {
                openFolderDialog.Description = "Select export Directory:";
            };
            var result = openFolderDialog.ShowDialog();
            if (result == true)
            {
                expFolderPath = openFolderDialog.SelectedPath + @"\";
                TB_ExpFolderPath.Text = expFolderPath;
                B_OpenExpFolder.IsEnabled = true;
            }
            EnableButtonAndChangeCursor(sender);
        }
        private void B_OpenExpFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("explorer.exe", @expFolderPath);
            }
            catch (Exception ex)
            {
                TB_Status.AddLine($"\n{ex.Message}");
                TB_Status.AddLine(ex.StackTrace != null ? $"\n{ex.StackTrace}" : "");
            }
        }
        private async void B_SelectTextsXLSX_ClickAsync(object sender, RoutedEventArgs e)
        {
            DisableButtonAndChangeCursor(sender);
            textFile_g = SelectXlsxFileAndTryToUse("Select Text File (.xlsx)");
            if (textFile_g != null)
            {
                L_TextfilePath.Text = textFile_g.FullName;
                TB_Status.AddLine($"\nSelected: {textFile_g.FullName}");
                try
                {
                    if (textDataTable_g.Rows.Count >= 0)
                        textDataTable_g.Clear();
                    textDataTable_g = await ExcelOperations.FromExcelFileToDataTable(textFile_g);
                    TB_Status.AddLine($"\nAcquired {textDataTable_g.Rows.Count} texts");
                    textDataTable_g.CheckHeaders(wsData_g);
                    UpdateUIDataWithWSData(wsData_g);
                    UpdateUIDataWithWSData(wsData_g);
                    CheckWSData();
                }
                catch (Exception ex)
                {
                    TB_Status.AddLine($"\n{ex.Message}");
                    TB_Status.AddLine($"\n{ex.StackTrace}");
                }
            }
            EnableButtonAndChangeCursor(sender);
        }
        private async void B_Translate_ClickAsync(object sender, RoutedEventArgs e)
        {
            DisableButtonAndChangeCursor(sender);
            string nameExtension = "_translated";
            if (textFile_g == null)
            {
                TB_Status.AddLine("\nFailed to duplicate (.xlsx) file! Selected file is null!");
                return;
            }
            var excelPackage = ExcelOperations.DuplicateExcelFile(textFile_g, nameExtension);
            if (excelPackage == null)
            {
                TB_Status.AddLine("\nFailed to duplicate (.xlsx) file! Selected file not correct!");
                return;
            }
            textToTranslateList = textDataTable_g.GetTextList(wsData_g);
            TB_Status.AddLine($"\nAcquired {textToTranslateList.Count()} non empty texts");
            //var ws = excelPackage.Workbook.Worksheets[0];
            //var range = ws.Cells["A1"].LoadFromDataTable(textDataTable_g, true);
            //range.AutoFitColumns();
            //var newName = excelPackage.File.FullName;
            //await ExcelOperations.SaveExcelFile(excelPackage);
            //TB_Status.AddLine($"\nCreated file : {newName}");
            EnableButtonAndChangeCursor(sender);
        }
        #endregion
        #region UI Extensions
        public void DisableButtonAndChangeCursor(object sender)
        {
            Cursor = Cursors.Wait;
            Button button = (Button)sender;
            button.IsEnabled = false;
        }
        public void EnableButtonAndChangeCursor(object sender)
        {
            Cursor = Cursors.Arrow;
            Button button = (Button)sender;
            button.IsEnabled = true;
        }
        public void CheckWSData()
        {
            if (textDataTable_g.Rows.Count >= 0 && wsData_g.CheckData()) B_Translate.IsEnabled = true;
        }
        public FileInfo? SelectXlsxFileAndTryToUse(string title)
        {
            OpenFileDialog openFileDialog1 = new()
            {
                InitialDirectory = @"c:\Users\localadm\Desktop",
                Title = title,
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xlsx",
                Filter = "Excel file (*.xlsx)|*.xlsx",
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true,
            };
            if (openFileDialog1.ShowDialog() == true)
            {
                FileInfo xmlFile = new(openFileDialog1.FileName);
                if (xmlFile.Exists && !xmlFile.IsFileLocked(xmlFile.FullName))
                {
                    return xmlFile;
                }
                TB_Status.AddLine("\nFile not exist or in use!");
                return null;
            }
            else
            {
                TB_Status.AddLine("\nFile not selected!");
                return null;
            }
        }
        public void UpdateUIDataWithWSData(WSData wSData)
        {
            if (TB_colId.Text != wSData.IDColumn.ToString())
            {
                TB_Status.AddLine($"\nColumn id has been updated {TB_colId.Text} -> {wSData.IDColumn}");
                TB_colId.Text = wSData.IDColumn.ToString();
            }
            if (TB_colSrc.Text != wSData.SrcColumn.ToString())
            {
                TB_Status.AddLine($"\nColumn source has been updated {TB_colSrc.Text} -> {wSData.SrcColumn}");
                TB_colSrc.Text = wSData.SrcColumn.ToString();
            }
            if (TB_colTrg.Text != wSData.TrgColumn.ToString())
            {
                TB_Status.AddLine($"\nColumn target has been updated {TB_colTrg.Text} -> {wSData.TrgColumn}");
                TB_colTrg.Text = wSData.TrgColumn.ToString();
            }
            if (TB_srcLang.Text != wSData.SrcLangCode)
            {
                TB_Status.AddLine($"\nSource language has been updated {TB_srcLang.Text} -> {wSData.SrcLangCode}");
                TB_srcLang.Text = wSData.SrcLangCode;
            }
            if (TB_trgLang.Text != wSData.TrgLangCode)
            {
                TB_Status.AddLine($"\nTarget language has been updated {TB_trgLang.Text} -> {wSData.TrgLangCode}");
                TB_trgLang.Text = wSData.TrgLangCode;
            }
        }

        #endregion
        #region UI Input Validation
        private void TB_colId_textChanged(object sender, TextChangedEventArgs e)
        {
            int colId = 0;
            TextBox textBox = (TextBox)sender;
            if (int.TryParse(textBox.Text, out colId))
            {
                wsData_g.IDColumn = colId;
                textBox.Background = Brushes.LightGreen;
                wsData_g.valOk[0] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                wsData_g.valOk[0] = false;
            }
        }
        private void TB_colSrc_textChanged(object sender, TextChangedEventArgs e)
        {
            int colSrc = 0;
            TextBox textBox = (TextBox)sender;
            if (int.TryParse(textBox.Text, out colSrc))
            {
                wsData_g.SrcColumn = colSrc;
                textBox.Background = Brushes.LightGreen;
                wsData_g.valOk[1] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                wsData_g.valOk[1] = false;
            }
        }
        private void TB_srcLang_textChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string inputValue = textBox.Text;
            if (inputValue == "auto")
            {
                wsData_g.SrcLangCode = inputValue;
                textBox.Background = Brushes.LightGreen;
                wsData_g.valOk[2] = true;
            }
            else if (langCodeList_g.Contains(inputValue))
            {
                wsData_g.SrcLangCode = inputValue;
                textBox.Background = Brushes.LightGreen;
                wsData_g.valOk[2] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                wsData_g.valOk[2] = false;
            }
        }
        private void TB_colTrg_textChanged(object sender, TextChangedEventArgs e)
        {
            int colTrg = 0;
            TextBox textBox = (TextBox)sender;
            if (int.TryParse(textBox.Text, out colTrg))
            {
                wsData_g.TrgColumn = colTrg;
                textBox.Background = Brushes.LightGreen;
                wsData_g.valOk[3] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                wsData_g.valOk[3] = false;
            }
        }
        private void TB_trgLang_textChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string inputValue = textBox.Text;
            if (langCodeList_g.Contains(inputValue))
            {
                wsData_g.TrgLangCode = inputValue;
                textBox.Background = Brushes.LightGreen;
                wsData_g.valOk[4] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                wsData_g.valOk[4] = false;
            }
        }
        #endregion


    }
}
