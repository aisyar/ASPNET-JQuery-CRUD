using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using TestAisyar.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TestAisyar.Controllers
{
    public class HomeController : Controller
    {
        private readonly Entities _db;
        public HomeController()
        {
            _db = new Entities();
        }

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public JsonResult getProvince()
        {
            var data = _db.Provinces.Select(fl => new { value = fl.ProvinceCode, text = fl.ProvinceDesc });
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult getCityByProvince(string provinceCode)
        {
            var data = _db.Cities.Where(fl => fl.ProvinceCode == provinceCode).Select(fl => new { value = fl.CityCode, text = fl.CityDesc });
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult getSuppliers(string SUPPLIER_CODE, string PROVINCE_CODE, string CITY_CODE)
        {
            IQueryable<ViewSupplier> data = _db.ViewSuppliers;
            if (!string.IsNullOrEmpty(SUPPLIER_CODE))
            {
                data = data.Where(fl => fl.SupplierCode == SUPPLIER_CODE);
            }
            if (!string.IsNullOrEmpty(PROVINCE_CODE))
            {
                data = data.Where(fl => fl.ProvinceCode == PROVINCE_CODE);
            }
            if (!string.IsNullOrEmpty(CITY_CODE))
            {
                data = data.Where(fl => fl.CityCode == CITY_CODE);
            }
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult Save(FormCollection form)
        {
            try
            {
                var suppcode = form["SUPPLIER_CODE"];
                var existing = _db.Suppliers.FirstOrDefault(fl => fl.SupplierCode == suppcode);
                if (existing != null)
                {
                    existing.SupplierCode = form["SUPPLIER_CODE"];
                    existing.SupplierName = form["SUPPLIER_NAME"];
                    existing.Address = form["ADDRESS"];
                    existing.ProvinceCode = form["PROVINCE_CODE"];
                    existing.CityCode = form["CITY_CODE"];
                    existing.PIC = form["PIC"];
                    _db.SaveChanges();

                    return Json(new Supplier());
                }
                else
                {
                    var supp = new Supplier();
                    supp.SupplierCode = form["SUPPLIER_CODE"];
                    supp.SupplierName = form["SUPPLIER_NAME"];
                    supp.Address = form["ADDRESS"];
                    supp.ProvinceCode = form["PROVINCE_CODE"];
                    supp.CityCode = form["CITY_CODE"];
                    supp.PIC = form["PIC"];
                    _db.Suppliers.Add(supp);
                    _db.SaveChanges();

                    return Json(supp);
                }

            }
            catch (Exception ex)
            {
                return Json(new { error = ex.Message + ex.InnerException != null ? ex.InnerException.Message : "" });
            }
        }

        [HttpPost]
        public ActionResult Delete(List<string> VALUES)
        {
            var result = new string[] { "OK", "" };
            _db.Suppliers.RemoveRange(_db.Suppliers.Where(fl => VALUES.Contains(fl.SupplierCode)));
            _db.SaveChanges();
            return Json(result);
        }

        public FileContentResult Download()
        {

            var fileDownloadName = String.Format($"Supp - {DateTime.Now.ToString("yyyy-MMM-dd")}.xlsx");
            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";


            // Pass your ef data to method
            ExcelPackage package = GenerateExcelFile(_db.ViewSuppliers.ToList());

            var fsr = new FileContentResult(package.GetAsByteArray(), contentType);
            fsr.FileDownloadName = fileDownloadName;

            return fsr;
        }

        private static ExcelPackage GenerateExcelFile(IEnumerable<ViewSupplier> datasource)
        {
            //EPPLUS Lib
            ExcelPackage pck = new ExcelPackage();

            //Create the worksheet 
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet 1");

            // Sets Headers
            ws.Cells[1, 1].Value = "SupplierCode";
            ws.Cells[1, 2].Value = "SupplierName";
            ws.Cells[1, 3].Value = "Address";
            ws.Cells[1, 4].Value = "ProvinceCode";
            ws.Cells[1, 5].Value = "ProvinceDesc";
            ws.Cells[1, 6].Value = "CityCode";
            ws.Cells[1, 7].Value = "CityDesc";
            ws.Cells[1, 8].Value = "PIC";

            // Inserts Data
            for (int i = 0; i < datasource.Count(); i++)
            {
                ws.Cells[i + 2, 1].Value = datasource.ElementAt(i).SupplierCode;
                ws.Cells[i + 2, 2].Value = datasource.ElementAt(i).SupplierName;
                ws.Cells[i + 2, 3].Value = datasource.ElementAt(i).Address;
                ws.Cells[i + 2, 4].Value = datasource.ElementAt(i).ProvinceCode;
                ws.Cells[i + 2, 5].Value = datasource.ElementAt(i).ProvinceDesc;
                ws.Cells[i + 2, 6].Value = datasource.ElementAt(i).CityCode;
                ws.Cells[i + 2, 7].Value = datasource.ElementAt(i).CityDesc;
                ws.Cells[i + 2, 8].Value = datasource.ElementAt(i).PIC;
            }

            // Format Header of Table
            using (ExcelRange rng = ws.Cells["A1:H1"])
            {

                rng.Style.Font.Bold = true;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid; //Set Pattern for the background to Solid 
                rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gold); //Set color to DarkGray 
                rng.Style.Font.Color.SetColor(System.Drawing.Color.Black);
            }
            return pck;
        }

        [HttpPost]
        public JsonResult Upload()
        {
            var result = new string[] { "OK", "Upload Failed." };

            HttpPostedFileBase file = Request.Files["fileUpload"];
            if (file != null)
            {
                string fileName = file.FileName;
                string fileContentType = file.ContentType;
                byte[] fileBytes = new byte[file.ContentLength];
                var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));
                List<Supplier> suppList = new List<Supplier>();

                try
                {
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var workSheet = currentSheet.First();
                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;
                        for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                        {
                            var supp = new Supplier();
                            supp.SupplierCode = workSheet.Cells[rowIterator, 1].Value.ToString();
                            supp.SupplierName = workSheet.Cells[rowIterator, 2].Value.ToString();
                            supp.Address = workSheet.Cells[rowIterator, 3].Value.ToString();
                            supp.ProvinceCode = workSheet.Cells[rowIterator, 4].Value.ToString();
                            supp.CityCode = workSheet.Cells[rowIterator, 6].Value.ToString();
                            supp.PIC = workSheet.Cells[rowIterator, 8].Value.ToString();
                            suppList.Add(supp);

                        }
                    }

                    _db.Suppliers.AddRange(suppList);
                    _db.SaveChanges();
                    result = new string[] { "OK", "Upload Successfull." };
                }
                catch (Exception ex)
                {
                    return Json(new string[] { "error", ex.Message });
                }
            }

            return Json(result);
        }
    }
}