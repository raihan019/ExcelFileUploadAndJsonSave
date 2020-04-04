using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelFileRead.Models;
using ExcelFileRead.DAL;
using System.Data;
using System.IO;
using Newtonsoft.Json;

namespace ExcelFileRead.Controllers
{
    public class FileTemplateController : Controller
    {

        FileOperationDAL fileOpDal = new FileOperationDAL();
        // GET: FileTemplate
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult RetrieveDataFromExcel()
        {
            List<FileDataModel> datalist = new List<FileDataModel>();
            if (Request.Files.Count > 0)
            {
               
                string fileName = fileOpDal.SaveExcellFileInDirectory(Request.Files, "ExcelFiles");
                DataTable dt = fileOpDal.GetDataListFromExcel(fileName, "ExcelFiles");
                datalist = fileOpDal.DataTableToObjList(dt);

            }

             return PartialView("_ExcelDataToTable", datalist);
          
        }

        public bool SaveDataInJson(List<FileDataModel> aList)
        {

            string json = JsonConvert.SerializeObject(aList);

            //write string to file
            System.IO.File.WriteAllText(@"D:\ExcelFiles\JsonFile.txt", json);
            return true;

        }
    }

   
}