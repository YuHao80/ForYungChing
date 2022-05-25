using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NorthWind.Controllers
{
    public class ProductsController : Controller
    {
        // GET: Products
        public ActionResult Index()
        {
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
                    var product = (from p in db.Products where p.ProductID == postBack.ProductID select p).FirstOrDefault();

                    if(product != default(Models.Products)) //判斷是否有此筆資料
                    {
                        TempData["ResultMsg"] = "已有此產品編號，請重新操作";
                        return RedirectToAction("Index");
                    } else {   //沒資料則繼續新增                        
                        //將回傳資料加入product
                        product = new Models.Products() { };

                        product.ProductID = (from p in db.Products select p).Max(p => p.ProductID) + 1;
                        product.ProductName = postBack.ProductName;
                        product.SupplierID = postBack.SupplierID;
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
            }
            //失敗訊息
            ViewBag.ResultMsg = "資料有誤，請檢查";

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

                    product.ProductName = postBack.ProductName;
                    product.SupplierID = postBack.SupplierID;
                    product.CategoryID = postBack.CategoryID;
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

    }
}