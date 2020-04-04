using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using ExcelFileRead.Models;

namespace ExcelFileRead.DAL
{
    public class FileOperationDAL
    {

        public string SaveExcellFileInDirectory(HttpFileCollectionBase postedFile, string folderName)
        {

            string fileName = null;
            for (int i = 0; i < postedFile.Count; i++)
            {
                HttpPostedFileBase file = postedFile[i];

                string fname = file.FileName;

                string path = @"D:\\" + folderName + "";

                if (Directory.Exists(path))
                {

                    fileName = Path.Combine(path, DateTime.Now.ToString("dd-MMM-yyyy-hh-mm-ss") + "_" + fname);
                    file.SaveAs(fileName);
                }
                else
                {
                    DirectoryInfo di = Directory.CreateDirectory(path);
                    fileName = Path.Combine(path, DateTime.Now.ToString("dd-MMM-yyyy-hh-mm-ss") + "_" + fname);
                    file.SaveAs(fileName);
                }

            }

            return fileName;
        }
        public DataTable GetDataListFromExcel(string fileName, string folderName)
        {
            string path = @"D:\\" + folderName + "\\";
            string extension = Path.GetExtension(fileName);

            string conString = string.Empty;
            switch (extension)
            {
                case ".xls": //Excel 97-03.
                    conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    break;
                case ".xlsx": //Excel 07 and above.
                    conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                    break;
            }
            string filePath = path + Path.GetFileName(fileName);
            conString = string.Format(conString, filePath);
            DataTable dt = new DataTable();
            try
            {
                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            //Get the name of First Sheet.
                            connExcel.Open();
                            var dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[2]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            //Read Data from First Sheet.
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT Requirement From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            connExcel.Close();
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
            return dt;
        }


        public List<FileDataModel> DataTableToObjList(DataTable dt)
        {
            try
            {
                List<FileDataModel> aList = new List<FileDataModel>();

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[i][0].ToString()))
                    {
                        FileDataModel aInfo = new FileDataModel();
                        aInfo.Requirement = dt.Rows[i][0] == DBNull.Value ? null : Convert.ToString(dt.Rows[i][0]).Trim(' ', '.');
                        aList.Add(aInfo);
                    }
                }
                return aList;
            }
            catch (Exception exception)
            {

                throw exception;
            }
        }
    }
}