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
using Path = System.IO.Path;

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
        public static Stopwatch stopWatch_g = null!;
        public static Progress<int> progress = null!;

        public MainWindow()
        {
            wsData_g = new();
            langCodeList_g = LangCodes.CreateLanguageCodes();
            InitializeComponent();
            textToTranslateList = new();
            textDataTable_g = new();
            stopWatch_g = new();
            progress = new Progress<int>(val => PB_Status.Value = val);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            TB_StatusBar.Text = "Select File (.xlsx) to translate";
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
                TB_Status.AddLine($"{ex.Message}");
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
                TB_Status.AddLine($"Selected: {textFile_g.FullName}");
                CheckPathAndFillTextBlock(expFolderPath, textFile_g.FullName, TB_ExpFolderPath);
                try
                {
                    if (textDataTable_g.Rows.Count >= 0)
                        textDataTable_g.Clear();
                    TB_StatusBar.Text = "Loading DataTable from Excelfile...";
                    textDataTable_g = await ExcelOperations.FromExcelFileToDataTable(textFile_g);
                    TB_Status.AddLine($"Acquired {textDataTable_g.Rows.Count} texts");
                    TB_StatusBar.Text = "Checking headers...";
                    textDataTable_g.CheckHeaders(wsData_g);
                    UpdateUIDataWithWSData(wsData_g);
                    CheckWSData();
                }
                catch (Exception ex)
                {
                    TB_Status.AddLine($"{ex.Message}");
                    TB_Status.AddLine($"{ex.StackTrace}");
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
                TB_Status.AddLine("Failed to duplicate (.xlsx) file! Selected file is null!");
                return;
            }
            if(textFile_g.IsFileLocked())
            {
                TB_Status.AddLine("Failed to duplicate (.xlsx) file! File not exist or is being used!");
                return;
            }
            var excelPackage = ExcelOperations.DuplicateExcelFile(textFile_g, expFolderPath, nameExtension);
            if (excelPackage == null)
            {
                TB_Status.AddLine("Failed to duplicate (.xlsx) file! Selected file not correct!");
                return;
            }
            TB_StatusBar.Text = "Creating Textlist...";
            textToTranslateList = textDataTable_g.GetTextList(wsData_g);
            TB_Status.AddLine($"Acquired {textToTranslateList.Count()} non empty texts");
            TB_StatusBar.Text = "Removing duplicates in Textlist...";
            var shortVerTextList = textToTranslateList.GetListWithoutDuplicatedSource();
            TB_Status.AddLine($"Acquired {shortVerTextList.Count()} non empty UNIQUE texts");
            stopWatch_g.Reset();
            stopWatch_g.Start();
            TB_StatusBar.Text = "Translating Textlist...";
            PB_Status.Maximum = shortVerTextList.Count;
            await shortVerTextList.TranslateAsync(wsData_g.SrcLangCode,wsData_g.TrgLangCode, progress);
            stopWatch_g.Stop();
            TB_Status.AddLine($"Translated in {stopWatch_g.Elapsed}.");
            TB_StatusBar.Text = "Filling Textlist with translations...";
            textToTranslateList.FillListWithTranslationsList(shortVerTextList);
            TB_StatusBar.Text = "Updating DataTable with Textlist...";
            textDataTable_g.UpdateWithTextList(textToTranslateList, wsData_g);
            var ws = excelPackage.Workbook.Worksheets[0];
            TB_StatusBar.Text = "Loading DataTable to Excelfile...";
            var range = ws.Cells["A1"].LoadFromDataTable(textDataTable_g, true);
            range.AutoFitColumns();
            var newName = excelPackage.File.FullName;
            TB_StatusBar.Text = "Saving Excelfile...";
            await ExcelOperations.SaveExcelFile(excelPackage);
            TB_StatusBar.Text = "Translations made!";
            TB_Status.AddLine($"Created file : {newName}");
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
        private string? CheckPathAndFillTextBlock(string path, string path2, TextBlock textBlock)
        {
            if (path != string.Empty || path == null)
                return path;
            var dir = Path.GetDirectoryName(path2);
            textBlock.Text = dir;
            B_OpenExpFolder.IsEnabled = true;
            return dir;
        }
        private void CheckWSData()
        {
            if (textDataTable_g.Rows.Count >= 0 && wsData_g.CheckData())
            {
                B_Translate.IsEnabled = true;
                TB_StatusBar.Text = "Configuration OK! Click MakeTranslations Button";

            }
            else
            {
                B_Translate.IsEnabled = false;
                TB_StatusBar.Text = "Configuration not OK";
            }
        }
        private FileInfo? SelectXlsxFileAndTryToUse(string title)
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
                if (xmlFile.Exists && !xmlFile.IsFileLocked())
                {
                    return xmlFile;
                }
                TB_Status.AddLine("File not exist or in use!");
                return null;
            }
            else
            {
                TB_Status.AddLine("File not selected!");
                return null;
            }
        }
        private void UpdateUIDataWithWSData(WSData wSData)
        {
            if (TB_colId.Text != wSData.IDColumn.ToString())
            {
                TB_Status.AddLine($"Column id has been updated {TB_colId.Text} -> {wSData.IDColumn}");
                TB_colId.Text = wSData.IDColumn.ToString();
            }
            if (TB_colSrc.Text != wSData.SrcColumn.ToString())
            {
                TB_Status.AddLine($"Column source has been updated {TB_colSrc.Text} -> {wSData.SrcColumn}");
                TB_colSrc.Text = wSData.SrcColumn.ToString();
            }
            if (TB_colTrg.Text != wSData.TrgColumn.ToString())
            {
                TB_Status.AddLine($"Column target has been updated {TB_colTrg.Text} -> {wSData.TrgColumn}");
                TB_colTrg.Text = wSData.TrgColumn.ToString();
            }
            if (TB_srcLang.Text != wSData.SrcLangCode)
            {
                TB_Status.AddLine($"Source language has been updated {TB_srcLang.Text} -> {wSData.SrcLangCode}");
                TB_srcLang.Text = wSData.SrcLangCode;
            }
            if (TB_trgLang.Text != wSData.TrgLangCode)
            {
                TB_Status.AddLine($"Target language has been updated {TB_trgLang.Text} -> {wSData.TrgLangCode}");
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
