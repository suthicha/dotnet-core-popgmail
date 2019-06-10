using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace dotnet_core_popgmail {
    public class InvoiceExp{

        private string _sqlConnection;
        private int MAX_FILE = 10;
        public InvoiceExp(string sqlConnection){
            _sqlConnection = sqlConnection;
        }

        public void ReadExcelToDb(){

            var files = GetFiles();

            foreach (string file in files)
            {
                var obj = ConvertExcelToInvoiceObj(file);
                var execStatus = Insert(obj.ToArray());
                
                if (execStatus){
                    DeleteFile(file);
                }

                break;
            }


        }

        private void DeleteFile(string filename){
            try{
                File.Delete(filename);
            }catch{}
        }

        private bool Insert(params Invoice[] invoices){

            using (SqlConnection conn = new SqlConnection(_sqlConnection))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var trans = conn.BeginTransaction();
                cmd.Transaction = trans;

                try {

                    cmd.CommandText = "SP_DT010_WMLOT_INSERT";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;

                    cmd.Parameters.AddRange(new SqlParameter[]{
                        new SqlParameter("@INVNO", SqlDbType.NVarChar),
                        new SqlParameter("@ITEM", SqlDbType.NVarChar),
                        new SqlParameter("@PART", SqlDbType.NVarChar),
                        new SqlParameter("@SPEC", SqlDbType.NVarChar),
                        new SqlParameter("@INVENTORY_NO", SqlDbType.NVarChar),
                        new SqlParameter("@IMP_INVNO", SqlDbType.NVarChar),
                        new SqlParameter("@QTY", SqlDbType.Decimal),
                        new SqlParameter("@PONO", SqlDbType.NVarChar)
                    });

                    for(var i = 0; i < invoices.Length; i++){
                        var inv = invoices[i];
                        cmd.Parameters["@INVNO"].Value = inv.InvoiceNo;
                        cmd.Parameters["@ITEM"].Value = inv.Item;
                        cmd.Parameters["@PART"].Value = inv.Part;
                        cmd.Parameters["@SPEC"].Value = inv.Spec;
                        cmd.Parameters["@INVENTORY_NO"].Value = inv.InventoryNo;
                        cmd.Parameters["@IMP_INVNO"].Value = inv.ImpInvoiceNo;
                        cmd.Parameters["@QTY"].Value = inv.Qty;
                        cmd.Parameters["@PONO"].Value = inv.PONO;
                        cmd.ExecuteNonQuery();
                    }

                    trans.Commit();

                    return true;

                }catch {
                    trans.Rollback();
                }
            }
            return false;
        }

        private List<Invoice> ConvertExcelToInvoiceObj(string filename){
            
            List<Invoice> invObj = new List<Invoice>();
            try {

                using (var workbook = new XLWorkbook(filename))
                {
                    var dataRows = workbook.Worksheet(1).RowsUsed();
                    
                    foreach (var dataRow in dataRows)
                    {
                        if (dataRow.RowNumber() > 1){
                            var inv = new Invoice();
                            inv.InvoiceNo = dataRow.Cell(1).Value.ToString();
                            inv.Item = dataRow.Cell(2).Value.ToString();
                            inv.Part = dataRow.Cell(3).Value.ToString();
                            inv.Spec = dataRow.Cell(4).Value.ToString();
                            inv.InventoryNo =  dataRow.Cell(5).Value.ToString();
                            inv.ImpInvoiceNo = dataRow.Cell(6).Value.ToString();
                            inv.Qty = Convert.ToDecimal(dataRow.Cell(7).Value.ToString());
                            inv.PONO = dataRow.Cell(8).Value.ToString();
                            invObj.Add(inv);
                        }
                    }
                }

            }catch {
                
            }
            return invObj;
        }

        private List<String> GetFiles(){

            string path = Path.Combine(Utils.AssemblyDirectory, "temp");
            DirectoryInfo di = new DirectoryInfo(path);
            FileInfo[] fi = di.GetFiles("*.xlsx");

            List<String> files = new List<string>();
            int count = 0;
            foreach (var file in fi)
            {            
                files.Add(file.FullName);
                if (count == MAX_FILE) break;
                count++;
            }

            return files;
        }
    }
}