using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using CompareExcelItem.Model;
using System.Linq;

namespace CompareExcelItem.Service
{
    public class ImportFileService
    {
        public DataTable ExcelToDataTable(string txtFilePath, int sheetIndex, int columns)
        {
            DataTable dt = new DataTable();
            FileInfo filePath = new FileInfo(txtFilePath);
            ExcelPackage ep = new ExcelPackage(filePath);
            ExcelWorksheet sheet = ep.Workbook.Worksheets[sheetIndex + 1];
            int startRowNumber = sheet.Dimension.Start.Row + 1;//起始列編號，從1算起
            int endRowNumber = sheet.Dimension.End.Row;//結束列編號，從1算起
            int startColumn = 1; //sheet.Dimension.Start.Column;//開始欄編號，從1算起
            int endColumn = 2; //sheet.Dimension.End.Column;//結束欄編號，

            //建立欄位名稱
            for (int k = 1; k <= columns; k++)
            {
                dt.Columns.Add(sheet.Cells[1, k].Value.ToString());
            }

            try
            {
                //寫入資料到資料列
                for (int currentRow = startRowNumber; currentRow <= endRowNumber; currentRow++)
                {
                    dt.NewRow();
                    object[] cell = new object[columns];
                    int idx = 0;
                    for (int i = 1; i <= columns; i++)
                    {
                        cell[idx] = sheet.Cells[currentRow, i].Value;
                        idx++;
                    }
                    dt.Rows.Add(cell);
                }

            }
            catch (Exception ex)
            {
                throw;
            }

            return dt;
        }

        public DataTable CompareRevisionAndReturn(DataTable dtRevision, List<Material> returnList)
        {
            try
            {
                for (int i = 0; i < dtRevision.Rows.Count; i++)
                {
                    //去退料檔找序號
                    List<Material> findList = returnList.Where(x => x.Model == dtRevision.Rows[i]["品號"].ToString().Replace("-C","")).ToList();

                    dtRevision.Rows[i][3] = findList.Count;

                    int serialIndex = 4;
                    string strIndex = "\"" + serialIndex.ToString() + "\"";
                    foreach (var item in findList)
                    {
                        dtRevision.Rows[i][serialIndex] = item.SerialNumber;
                        serialIndex++;
                    }

                }
            }
            catch (Exception ex)
            {
                throw;
            }

            return dtRevision;
        }


    }
}
