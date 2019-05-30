using System;
using System.IO;
using System.Data;

using ClosedXML.Excel;

namespace dotnet_core_popgmail {

    public class ExcelAdapter {
        public DataTable GetDataFromExcel(string path){

            DataTable dt = new DataTable();
            using (XLWorkbook workBook = new XLWorkbook(path)){
                IXLWorksheet workSheet = workBook.Worksheet(0);

                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows()){
                    if (firstRow){
                        foreach (IXLCell cell in row.Cells())
                        {
                            if (!string.IsNullOrEmpty(cell.Value.ToString()))
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            else
                            {
                                break;
                            }
                        }
                        firstRow = false;
                    }else{
                        int i = 0;
                        DataRow toInsert = dt.NewRow();
                        foreach (IXLCell cell in row.Cells(1, dt.Columns.Count))
                        {
                            try
                            {
                                toInsert[i] = cell.Value.ToString();
                            }
                            catch
                            {
                            }
                            i++;
                        }
                        dt.Rows.Add(toInsert);
                    }
                }
            }

            return dt;
        }
    }
    
}