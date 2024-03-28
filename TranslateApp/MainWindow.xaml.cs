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
using OfficeOpenXml.FormulaParsing.Logging;
using Microsoft.Extensions.Logging;
using Serilog;

namespace TranslateApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string _expFolderPath = string.Empty;
        private static FileInfo? _textFileG;
        private static List<TextToTranslate>? _textToTranslateList;
        private static DataTable _textDataTableG = null!;
        private static WSData _wsDataG = null!;
        private static List<string> _langCodeListG = null!;
        private static Stopwatch _stopWatchG = null!;
        private static Progress<int> _progress = null!;
        public MainWindow()
        {
            _wsDataG = new();
            _langCodeListG = LangCodes.CreateLanguageCodes();
            InitializeComponent();

            //Logger Configuration
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Verbose()
                .WriteTo.File("logs.txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();
                Log.Information("Logging started");

            _textToTranslateList = new();
            _textDataTableG = new();
            _stopWatchG = new();
            _progress = new Progress<int>(val => PB_Status.Value = val);
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
                TB_StatusBar.Text = "Operation failed. Check status text or log file";
                _expFolderPath = openFolderDialog.SelectedPath + @"\";
                TB_ExpFolderPath.Text = _expFolderPath;
                B_OpenExpFolder.IsEnabled = true;
            }
            EnableButtonAndChangeCursor(sender);
        }
        private void B_OpenExpFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("explorer.exe", _expFolderPath);
            }
            catch (Exception ex)
            {
                CatchExceptionBehaviour(ex);
            }
        }
        private async void B_SelectTextsXLSX_ClickAsync(object sender, RoutedEventArgs e)
        {
            DisableButtonAndChangeCursor(sender);
            _textFileG = SelectXlsxFileAndTryToUse("Select Text File (.xlsx)");
            if (_textFileG != null)
            {
                L_TextfilePath.Text = _textFileG.FullName;
                TB_Status.AddLine($"Selected: {_textFileG.FullName}");
                CheckPathAndFillTextBlock(_expFolderPath, _textFileG.FullName, TB_ExpFolderPath);
                try
                {
                    if (_textDataTableG.Rows.Count >= 0)
                        _textDataTableG.Clear();
                    TB_StatusBar.Text = "Loading DataTable from Excelfile...";
                    _textDataTableG = await ExcelOperations.FromExcelFileToDataTable(_textFileG);
                    TB_Status.AddLine($"Acquired {_textDataTableG.Rows.Count} texts");
                    TB_StatusBar.Text = "Checking headers...";
                    _textDataTableG.CheckHeaders(_wsDataG);
                    UpdateUIDataWithWSData(_wsDataG);
                    CheckWSData();
                }
                catch (Exception ex)
                {
                    CatchExceptionBehaviour(ex);
                }
            }
            EnableButtonAndChangeCursor(sender);
        }
        private async void B_Translate_ClickAsync(object sender, RoutedEventArgs e)
        {
            DisableButtonAndChangeCursor(sender);
            string nameExtension = "_translated";
            if (_textFileG == null)
            {
                TB_Status.AddLine("Failed to duplicate (.xlsx) file! Selected file is null!");
                EnableButtonAndChangeCursor(sender);
                return;
            }
            if(_textFileG.IsFileLocked())
            {
                TB_Status.AddLine("Failed to duplicate (.xlsx) file! File not exist or is being used!");
                EnableButtonAndChangeCursor(sender);
                return;
            }
            var excelPackage = ExcelOperations.DuplicateExcelFile(_textFileG, _expFolderPath, nameExtension);
            if (excelPackage == null)
            {
                TB_Status.AddLine("Failed to duplicate (.xlsx) file! Selected file not correct!");
                EnableButtonAndChangeCursor(sender);
                return;
            }

            try
            {
                TB_StatusBar.Text = "Creating Textlist...";
                _textToTranslateList = _textDataTableG.GetTextList(_wsDataG);
                TB_Status.AddLine($"Acquired {_textToTranslateList.Count} non empty texts");
                TB_StatusBar.Text = "Removing duplicates in Textlist...";
                var shortVerTextList = _textToTranslateList.GetListWithoutDuplicatedSource();
                TB_Status.AddLine($"Acquired {shortVerTextList.Count} non empty UNIQUE texts");
                _stopWatchG.Reset();
                _stopWatchG.Start();
                TB_StatusBar.Text = "Translating Textlist...";
                PB_Status.Maximum = shortVerTextList.Count;
                await shortVerTextList.TranslateAsync(_wsDataG.SrcLangCode, _wsDataG.TrgLangCode, _progress);
                _stopWatchG.Stop();
                TB_Status.AddLine($"Translated in {_stopWatchG.Elapsed}.");
                TB_StatusBar.Text = "Filling Textlist with translations...";
                _textToTranslateList.FillListWithTranslationsList(shortVerTextList);
                TB_StatusBar.Text = "Updating DataTable with Textlist...";
                _textDataTableG.UpdateWithTextList(_textToTranslateList, _wsDataG);
                var ws = excelPackage.Workbook.Worksheets[0];
                TB_StatusBar.Text = "Loading DataTable to Excelfile...";
                var range = ws.Cells["A1"].LoadFromDataTable(_textDataTableG, true);
                range.AutoFitColumns();
                var newName = excelPackage.File.FullName;
                TB_StatusBar.Text = "Saving Excelfile...";
                await ExcelOperations.SaveExcelFile(excelPackage);
                TB_StatusBar.Text = "Translations made!";
                TB_Status.AddLine($"Created file : {newName}");
            }
            catch (Exception ex)
            {
                CatchExceptionBehaviour(ex);
            }
            finally
            {
                EnableButtonAndChangeCursor(sender);
            }
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
            if (_textDataTableG.Rows.Count >= 0 && _wsDataG.CheckData())
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
            if (TB_colId.Text != wSData.IdColumn.ToString())
            {
                TB_Status.AddLine($"Column id has been updated {TB_colId.Text} -> {wSData.IdColumn}");
                TB_colId.Text = wSData.IdColumn.ToString();
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
        private void CatchExceptionBehaviour(Exception ex)
        {
            TB_StatusBar.Text = "Operation failed. Check status text or log file";
            TB_Status.AddLine($"{ex.Message}");
            TB_Status.AddLine(ex.StackTrace != null ? $"\n{ex.StackTrace}" : string.Empty);
            Log.Error(ex.Message + ex.StackTrace);
        }
        #endregion
        #region UI Input Validation
        private void TB_colId_textChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (int.TryParse(textBox.Text, out var colId))
            {
                _wsDataG.IdColumn = colId;
                textBox.Background = Brushes.LightGreen;
                _wsDataG.ValOk[0] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                _wsDataG.ValOk[0] = false;
            }
        }
        private void TB_colSrc_textChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (int.TryParse(textBox.Text, out var colSrc))
            {
                _wsDataG.SrcColumn = colSrc;
                textBox.Background = Brushes.LightGreen;
                _wsDataG.ValOk[1] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                _wsDataG.ValOk[1] = false;
            }
        }
        private void TB_srcLang_textChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string inputValue = textBox.Text;
            if (inputValue == "auto")
            {
                _wsDataG.SrcLangCode = inputValue;
                textBox.Background = Brushes.LightGreen;
                _wsDataG.ValOk[2] = true;
            }
            else if (_langCodeListG.Contains(inputValue))
            {
                _wsDataG.SrcLangCode = inputValue;
                textBox.Background = Brushes.LightGreen;
                _wsDataG.ValOk[2] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                _wsDataG.ValOk[2] = false;
            }
        }
        private void TB_colTrg_textChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (int.TryParse(textBox.Text, out var colTrg))
            {
                _wsDataG.TrgColumn = colTrg;
                textBox.Background = Brushes.LightGreen;
                _wsDataG.ValOk[3] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                _wsDataG.ValOk[3] = false;
            }
        }
        private void TB_trgLang_textChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string inputValue = textBox.Text;
            if (_langCodeListG.Contains(inputValue))
            {
                _wsDataG.TrgLangCode = inputValue;
                textBox.Background = Brushes.LightGreen;
                _wsDataG.ValOk[4] = true;
            }
            else
            {
                textBox.Background = Brushes.IndianRed;
                _wsDataG.ValOk[4] = false;
            }
        }
        #endregion
    }
}
