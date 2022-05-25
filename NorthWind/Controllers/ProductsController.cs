using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace NorthWind.Controllers {
    public class ProductsController : Controller {
        // GET: Products
        public ActionResult Index() {
            //宣告商品列表
            List<Models.Products> products = new List<Models.Products>();

            //接收新增完成的訊息
            ViewBag.ResultMsg = TempData["ResultMsg"];

            using(Models.NorthWindDBEntities db = new Models.NorthWindDBEntities()) {
                //Linq語法取資料
                products = (from p in db.Products select p).ToList();
            }

            return View(products);
        }

        //新增商品頁面
        public ActionResult Create() {
            return View();
        }

        //新增商品資料-資料傳回處理
        [HttpPost]  //指定只有使用Post方法才能進入
        public ActionResult Create(Models.Products postBack) {
            if(ModelState.IsValid) //驗證資料是否成功
            {
                using(Models.NorthWindDBEntities db = new Models.NorthWindDBEntities()) {

                    string productName = postBack.ProductName;
                    int? supplierID = postBack.SupplierID;
                    int? categoryID = postBack.CategoryID;
                    string errMsg = CheckData(productName, supplierID, categoryID);

                    if(!string.IsNullOrEmpty(errMsg)) {
                        //失敗訊息
                        ViewBag.ResultMsg = errMsg;

                        //停留在此頁面
                        return View(postBack);
                    }

                    //將回傳資料加入product
                    var product = new Models.Products() { };

                    product.ProductID = (from p in db.Products select p).Max(p => p.ProductID) + 1;
                    product.ProductName = productName;
                    product.SupplierID = supplierID;
                    product.CategoryID = postBack.CategoryID;
                    product.QuantityPerUnit = postBack.QuantityPerUnit;
                    product.UnitPrice = postBack.UnitPrice;
                    product.UnitsInStock = postBack.UnitsInStock;
                    product.UnitsOnOrder = postBack.UnitsOnOrder;
                    product.ReorderLevel = postBack.ReorderLevel;
                    product.Discontinued = postBack.Discontinued;
                    db.Products.Add(product);

                    //儲存異動資料
                    db.SaveChanges();

                    //設定成功訊息
                    TempData["ResultMsg"] = string.Format("商品[{0}]成功建立", postBack.ProductName);

                    //跳轉Product/Index頁面
                    return RedirectToAction("Index");
                }
            }
            //失敗訊息
            ViewBag.ResultMsg = "資料異常，請檢查";

            //停留在此頁面
            return View(postBack);
        }

        //編輯商品頁面
        public ActionResult Edit(int id) {
            using(Models.NorthWindDBEntities db = new Models.NorthWindDBEntities()) {
                //抓取Product.ID等於輸入ID的資料
                var product = (from p in db.Products where p.ProductID == id select p).FirstOrDefault();

                if(product != default(Models.Products)) //判斷是否有此筆資料
                {
                    return View(product);
                } else {   //沒資料則顯示錯誤訊息並導回Index頁面
                    TempData["ResultMsg"] = "資料有誤，請重新操作";
                    return RedirectToAction("Index");
                }
            }
        }

        //編輯商品資料-資料傳回處理
        [HttpPost]  //指定只有使用Post方法才能進入
        public ActionResult Edit(Models.Products postBack) {
            if(ModelState.IsValid) //驗證資料是否成功
            {
                using(Models.NorthWindDBEntities db = new Models.NorthWindDBEntities()) {
                    //抓取Product.ID等於輸入ID的資料
                    Models.Products product = (from p in db.Products where p.ProductID == postBack.ProductID select p).FirstOrDefault();

                    string productName = postBack.ProductName;
                    int? supplierID = postBack.SupplierID;
                    int? categoryID = postBack.CategoryID;
                    string errMsg = CheckData(productName, supplierID, categoryID);

                    if(!string.IsNullOrEmpty(errMsg)) {
                        //失敗訊息
                        ViewBag.ResultMsg = errMsg;

                        //停留在此頁面
                        return View(postBack);
                    }


                    product.ProductName = productName;
                    product.SupplierID = supplierID;
                    product.CategoryID = categoryID;
                    product.QuantityPerUnit = postBack.QuantityPerUnit;
                    product.UnitPrice = postBack.UnitPrice;
                    product.UnitsInStock = postBack.UnitsInStock;
                    product.UnitsOnOrder = postBack.UnitsOnOrder;
                    product.ReorderLevel = postBack.ReorderLevel;
                    product.Discontinued = postBack.Discontinued;

                    //儲存異動資料
                    db.SaveChanges();

                    //設定成功訊息
                    TempData["ResultMsg"] = string.Format("商品[{0}]成功編輯", postBack.ProductName);

                    //跳轉Product/Index頁面
                    return RedirectToAction("Index");
                }
            } else {
                //失敗訊息
                ViewBag.ResultMsg = "資料有誤，請重新操作";

                //停留在此頁面
                return View(postBack);
            }
        }

        //刪除商品資料
        [HttpPost]  //指定只有使用Post方法才能進入
        public ActionResult Delete(int id) {
            using(Models.NorthWindDBEntities db = new Models.NorthWindDBEntities()) {
                //抓取Product.ID等於輸入ID的資料
                Models.Products product = (from p in db.Products where p.ProductID == id select p).FirstOrDefault();

                if(product != default(Models.Products)) //判斷是否有此筆資料
                {
                    db.Products.Remove(product);

                    db.SaveChanges();

                    //設定成功訊息
                    TempData["ResultMsg"] = string.Format("商品[{0}]成功刪除", product.ProductName);

                    //跳轉Product/Index頁面
                    return RedirectToAction("Index");
                } else {
                    //失敗訊息
                    TempData["ResultMsg"] = "指定資料不存在，無法刪除，請重新操作";

                    //停留在此頁面
                    return RedirectToAction("Index");
                }
            }
        }

        private string CheckData(string productName, int? supplierID, int? categoryID) {
            if(string.IsNullOrEmpty(productName)) return "產品名稱不得為空";
            using(Models.NorthWindDBEntities db = new Models.NorthWindDBEntities()) {
                var supplier = (from s in db.Suppliers where s.SupplierID == supplierID select s).FirstOrDefault();
                if(supplierID.HasValue && supplier == default(Models.Suppliers)) {
                    return "無此廠商編號，請重新操作";
                }

                var category = (from c in db.Categories where c.CategoryID == categoryID select c).FirstOrDefault();
                if(categoryID.HasValue && category == default(Models.Categories)) {
                    return "無此類別編號，請重新操作";
                }
                return null;
            }
        }

        //public FileResult ExportExcel() {
        //    var sbHtml = new System.Text.StringBuilder();
        //    sbHtml.Append("<table border='1' cellspacing='0' cellpadding='0'>");
        //    sbHtml.Append("<tr>");
        //    var lstTitle = new List<string> { "編號", "姓名", "年齡", "建立時間" };
        //    foreach(var item in lstTitle) {
        //        sbHtml.AppendFormat("<td style='font-size: 14px;text-align:center;background-color: #DCE0E2; font-weight:bold;' height='25'>{0}</td>", item);
        //    }
        //    sbHtml.Append("</tr>");

        //    for(int i = 0; i < 1000; i++) {
        //        sbHtml.Append("<tr>");
        //        sbHtml.AppendFormat("<td style='font-size: 12px;height:20px;'>{0}</td>", i);
        //        sbHtml.AppendFormat("<td style='font-size: 12px;height:20px;'>屌絲{0}號</td>", i);
        //        sbHtml.AppendFormat("<td style='font-size: 12px;height:20px;'>{0}</td>", new Random().Next(20, 30) + i);
        //        sbHtml.AppendFormat("<td style='font-size: 12px;height:20px;'>{0}</td>", DateTime.Now);
        //        sbHtml.Append("</tr>");
        //    }
        //    sbHtml.Append("</table>");

        //    //第一種:使用FileContentResult
        //    byte[] fileContents = System.Text.Encoding.Default.GetBytes(sbHtml.ToString());
        //    return File(fileContents, "application/ms-excel", "fileContents.xls");

        //    //第二種:使用FileStreamResult
        //    var fileStream = new System.IO.MemoryStream(fileContents);
        //    return File(fileStream, "application/ms-excel", "fileStream.xls");

        //    //第三種:使用FilePathResult
        //    //伺服器上首先必須要有這個Excel檔案,然會通過Server.MapPath獲取路徑返回.
        //    var fileName = Server.MapPath("~/Files/fileName.xls");
        //    return File(fileName, "application/ms-excel", "fileName.xls");
        //}

        public ActionResult ExportExcel() {
            using(Models.NorthWindDBEntities db = new Models.NorthWindDBEntities()) {
                //取出要匯出Excel的資料
                List<Models.Products> products = (from p in db.Products select p).ToList();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //建立Excel
                ExcelPackage ep = new ExcelPackage();

                //建立第一個Sheet，後方為定義Sheet的名稱
                ExcelWorksheet sheet = ep.Workbook.Worksheets.Add("FirstSheet");

                int col = 1;    //欄:直的，因為要從第1欄開始，所以初始為1

                //第1列是標題列 
                //標題列部分，是取得DataAnnotations中的DisplayName，這樣比較一致，
                //這也可避免後期有修改欄位名稱需求，但匯出excel標題忘了改的問題發生。
                //取得做法可參考最後的參考連結。
                sheet.Cells[1, col++].Value = "產品編號";
                sheet.Cells[1, col++].Value = "產品名稱";
                sheet.Cells[1, col++].Value = "廠商編號";
                sheet.Cells[1, col++].Value = "類別編號";
                sheet.Cells[1, col++].Value = "每單位數量";
                sheet.Cells[1, col++].Value = "單價";
                sheet.Cells[1, col++].Value = "庫存單位";
                sheet.Cells[1, col++].Value = "訂購單位";
                sheet.Cells[1, col++].Value = "再次訂購量";
                sheet.Cells[1, col++].Value = "是否停產";

                //資料從第2列開始
                int row = 2;    //列:橫的
                foreach(var product in products) {
                    col = 1;//每換一列，欄位要從1開始
                            //指定Sheet的欄與列(欄名列號ex.A1,B20，在這邊都是用數字)，將資料寫入
                    sheet.Cells[row, col++].Value = product.ProductID;
                    sheet.Cells[row, col++].Value = product.ProductName;
                    sheet.Cells[row, col++].Value = product.SupplierID;
                    sheet.Cells[row, col++].Value = product.CategoryID; ;
                    sheet.Cells[row, col++].Value = product.QuantityPerUnit;
                    sheet.Cells[row, col++].Value = product.UnitPrice;
                    sheet.Cells[row, col++].Value = product.UnitsInStock;
                    sheet.Cells[row, col++].Value = product.UnitsOnOrder;
                    sheet.Cells[row, col++].Value = product.ReorderLevel;
                    sheet.Cells[row, col++].Value = product.Discontinued;
                    row++;
                }

                //因為ep.Stream是加密過的串流，故要透過SaveAs將資料寫到MemoryStream，在將MemoryStream使用FileStreamResult回傳到前端
                MemoryStream fileStream = new MemoryStream();
                ep.SaveAs(fileStream);
                ep.Dispose();//如果這邊不下Dispose，建議此ep要用using包起來，但是要記得先將資料寫進MemoryStream在Dispose。
                fileStream.Position = 0;//不重新將位置設為0，excel開啟後會出現錯誤
                return File(fileStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", DateTime.Now + "產品列表.xlsx");
            }
        }

    }
}