using ExcelApi.Models;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;

namespace ExcelApi.Controllers
{
    public class ExceloperationController : ApiController
    {
     
        [HttpPost]
        public HttpResponseMessage ExcelUpload()
        {
            string message = "";
            HttpResponseMessage result = null;
            var httpRequest = HttpContext.Current.Request;
            List<ExcelEntity> excelEntity = new List<ExcelEntity>();
            if (httpRequest.Files.Count > 0)
            {
                HttpPostedFile file = httpRequest.Files[0];
                Stream stream = file.InputStream;

                IExcelDataReader reader = null;

                if (file.FileName.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (file.FileName.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                else
                {
                  return  result = Request.CreateResponse(HttpStatusCode.BadRequest, "This file format is not supported"); 
                }

                DataSet excelRecords = reader.AsDataSet();
                reader.Close();

                var finalRecords = excelRecords.Tables[0];
                for (int i = 0; i < finalRecords.Rows.Count; i++)
                {
                    ExcelEntity objUser = new ExcelEntity();
                    objUser.Sno =Convert.ToString(finalRecords.Rows[i][0]);
                    objUser.Company = Convert.ToString(finalRecords.Rows[i][1]);
                    objUser.Sector = Convert.ToString(finalRecords.Rows[i][2]);
                    objUser.SubSector = Convert.ToString(finalRecords.Rows[i][3]);
                    objUser.Region = Convert.ToString(finalRecords.Rows[i][4]);
                    objUser.NoOfEmployee = Convert.ToString(finalRecords.Rows[i][5]);
                    objUser.TotalRevenu = Convert.ToString(finalRecords.Rows[i][6]);
                    objUser.Websites = Convert.ToString(finalRecords.Rows[i][7]);
                    excelEntity.Add(objUser);
                }
                result = Request.CreateResponse(HttpStatusCode.OK, excelEntity);
            }
            else
            {
                result = Request.CreateResponse(HttpStatusCode.BadRequest);
            }
           
            return result;
        }
    }
}
