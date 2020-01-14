using CompareExcelItem.Model;

using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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

namespace CompareExcelItem.Service
{
    public class ExportFileService
    {
        public ResponseMessage DataTableToExcelFile(DataTable dt, string ExcelSavePath)
        {
            bool SavetSwitch = false;
            string getExcelSaveFullPath = ExcelSavePath;
            string getExcelSaveDirectory = string.Empty;
            ResponseMessage result = new ResponseMessage();

            int columnWidth = 25;

            try
            {
                #region -- Check Excel Save Path --
               
                if (!string.IsNullOrEmpty(getExcelSaveFullPath))
                {
                    getExcelSaveDirectory = System.IO.Path.GetDirectoryName(ExcelSavePath);
                    if (!Directory.Exists(getExcelSaveDirectory))
                    {
                        Directory.CreateDirectory(getExcelSaveDirectory);
                    }

                    //if Create Directory fail
                    if (!Directory.Exists(getExcelSaveDirectory))
                    {
                        getExcelSaveDirectory = string.Empty;
                        result.Success = false;
                    }
                    else
                    {
                        SavetSwitch = true;
                    }
                }
                #endregion

                if (SavetSwitch)
                {
                    //避免Excek檔名重複
                    string dtNow = "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    ExcelSavePath = ExcelSavePath.Replace(".xlsx", dtNow + ".xlsx");

                    FileInfo filePath = new FileInfo(ExcelSavePath);
                    ExcelPackage ep = new ExcelPackage(filePath);
                    ExcelWorksheet ws;

                    if (dt.TableName != string.Empty)
                    {
                        ws = ep.Workbook.Worksheets.Add(dt.TableName);
                    }
                    else
                    {
                        ws = ep.Workbook.Worksheets.Add("Sheet1");
                    }

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        ws.Cells[1, i + 1].Value = dt.Columns[i].ColumnName;
                        ws.Column(i + 1).Width = columnWidth;
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            ws.Cells[i + 2, j + 1].Value = dt.Rows[i][j].ToString();
                        }
                    }

                    ws.Cells[1, 1, dt.Rows.Count + 2, dt.Columns.Count + 1].Style.Font.Size = 12;
                    ws.Cells[1, 1, 2, dt.Columns.Count + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[3, 1, dt.Rows.Count + 2, dt.Columns.Count + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[1, 1, dt.Rows.Count + 2, dt.Columns.Count + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells.Style.WrapText = true;
                    ep.Save();
                    ep.Dispose();
                    ep = null;
                    result.Success = true;
                    result.Message = "Export excel Success";

                }
                else
                {
                    result.Success = false;
                    result.Message = "File already exists : " + getExcelSaveFullPath;
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return result;
        }
    }
}
