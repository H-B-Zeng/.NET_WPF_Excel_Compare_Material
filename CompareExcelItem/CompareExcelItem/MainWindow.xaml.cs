using CompareExcelItem.Model;
using CompareExcelItem.Service;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows;
using System.Linq;

namespace CompareExcelItem
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnSelectExcel(object sender, RoutedEventArgs e)
        {
            OpenFileDialog chrooseFileDialog = new OpenFileDialog();
            chrooseFileDialog.DefaultExt = ".xlsx";
            chrooseFileDialog.Filter = "Excel files(.xlsx;)|*.xlsx;";
            chrooseFileDialog.Multiselect = false;
            chrooseFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            Nullable<bool> result = chrooseFileDialog.ShowDialog();
            string defaultSaveExcelPath = string.Empty;
            string errorMsg = "";
            List<string> viewSheets = new List<string>();

            try
            {
                if (result == true)
                {
                    txtFilePath.Text = chrooseFileDialog.FileName;
                    FileInfo filePath = new FileInfo(txtFilePath.Text);
                    ExcelPackage ep = new ExcelPackage(filePath);

                    foreach (var item in ep.Workbook.Worksheets)
                    {
                        viewSheets.Add(item.ToString());
                    }

                    //WPView dropdownlist
                    ddlExcelSheets.ItemsSource = viewSheets;

                    btnCheckExcelData();
                }
            }
            catch (Exception ex)
            {
                errorMsg = ex.ToString();
                throw;
            }
        }

        /// <summary>
        /// Compare Excel 檢查Excel前兩個欄位是否一樣
        /// </summary>
        private void btnCheckExcelData()
        {
            if (ddlExcelSheets.Items.Count == 0)
            {
                MessageBox.Show("Please Choose Excel File", "Info");
            }

            if (ddlExcelSheets.Items.Count > 0)
            {                
                ImportFileService importFileService = new ImportFileService();
                ExportFileService exportFileService = new ExportFileService();
                try
                {
                    //調整單
                    DataTable dtRevision = new DataTable();
                    dtRevision = importFileService.ExcelToDataTable(txtFilePath.Text, 0, 104);

                    //退料單
                    DataTable dtReturn = new DataTable();
                    dtReturn = importFileService.ExcelToDataTable(txtFilePath.Text, 1, 5);
                    var result = DataTableExtensions.ToList<Material>(dtReturn).ToList();
                    List<Material> returnList = result as List<Material>;

                    //Compare excel data
                    DataTable dt = importFileService.CompareRevisionAndReturn(dtRevision, returnList);
                    ResponseMessage response = exportFileService.DataTableToExcelFile(dt,txtFilePath.Text);
                }
                catch (Exception ex)
                {
                    throw;
                }

                //if (result.isSuccess)
                //{
                //    MessageBox.Show("檢查成功", "Info");
                //}
                //else
                //{
                //    MessageBox.Show(result.errorMsg, "error");
                //}

            }
        }
    }
}
