﻿using MVC.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace MVC.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        // GET: Home  
        public ActionResult UploadFiles()
        {
            return View();
        }
        [HttpPost]
        public ActionResult UploadFiles(HttpPostedFileBase[] files)
        {

            //Ensure model state is valid  
            if (ModelState.IsValid)
            {   //iterating through multiple file collection   
                List<FirmInfo> firmInfos = new List<FirmInfo>();
                List<FirmMap> firmMaps = new List<FirmMap>();
                foreach (HttpPostedFileBase file in files)
                {
                    //Checking file is available to save.  
                    if (file != null)
                    {
                        if (file.FileName.Contains("Firm Info")) firmInfos = getFile1(file);
                        else if (file.FileName.Contains("Asset Class - Firm Map")) firmMaps = getFile2(file);
                    }

                }

                if (firmInfos.Count > 0 && firmMaps.Count > 0)
                {
                    Dictionary<string, string> firms = new Dictionary<string, string>();
                    Dictionary<string, string> assets = new Dictionary<string, string>();
                    Dictionary<string, Dictionary<string, string>> relations = new Dictionary<string, Dictionary<string, string>>();
                    foreach (FirmInfo firm in firmInfos)
                    {
                        firms.Add(firm.FirmID, firm.FirmName);
                        relations.Add(firm.FirmID, new Dictionary<string, string>());
                    }
                    foreach (FirmMap firmMap in firmMaps)
                    {
                        if (!assets.ContainsKey(firmMap.AssetClassID)) assets.Add(firmMap.AssetClassID, firmMap.AssetClassName);
                        Dictionary<string, string> tmpDic = relations[firmMap.InterestedFirmsID];
                        tmpDic.Add(firmMap.AssetClassID, "true");
                        relations[firmMap.InterestedFirmsID] = tmpDic;
                    }
                    var orderedfirms = firms.OrderBy(x => x.Value);
                    firms = orderedfirms.ToDictionary(t => t.Key, t => t.Value);
                    var orderedassets = assets.OrderBy(x => x.Value);
                    assets = orderedassets.ToDictionary(t => t.Key, t => t.Value);

                    FirmData firmData = new FirmData();
                    firmData.Firms = firms;
                    firmData.Assets = assets;
                    firmData.Relations = relations;
                    ViewBag.FirmData = firmData;
                    string requestSent = "<?xml version=\"1.0\" encoding=\"UTF - 8\" standalone=\"yes\"?><PaymentRequestCommand><ScheduleId>UAT_1</ScheduleId><ClientId>NIBSS_V2001</ClientId><DebitBankCode>044</DebitBankCode><DebitAccountNumber>0123456789</DebitAccountNumber></PaymentRequestCommand>";
                    TempData["message"] = "Upload was successful";                    
                    return View(nameof(Index));

                }
                else
                {
                    TempData["message"] = "No file was uploaded";
                    string request = "<?xml version=\"1.0\" encoding=\"UTF - 8\" standalone=\"yes\"?><PaymentRequestCommand><ScheduleId>UAT_1</ScheduleId><ClientId>NIBSS_V2001</ClientId><DebitBankCode>044</DebitBankCode><DebitAccountNumber>0123456789</DebitAccountNumber></PaymentRequestCommand>";
                    return View(nameof(Index));
                }
            }
            return View();
        }

        public List<FirmInfo> getFile1(HttpPostedFileBase postedFile)
        {
            try
            {
                string path = Server.MapPath("~/Uploads/");
                string filePath = string.Empty;
                if (postedFile != null)
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    filePath = path + DateTime.Now.Ticks + "-" + Path.GetFileName(postedFile.FileName);
                    postedFile.SaveAs(filePath);

                    //read data from excel
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(filePath);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;

                    List<FirmInfo> firmInfos = new List<FirmInfo>();
                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        FirmInfo firminfo = new FirmInfo();
                        firminfo.FirmID = ((Excel.Range)range.Cells[row, 1]).Text;
                        firminfo.FirmName = ((Excel.Range)range.Cells[row, 2]).Text;
                        firmInfos.Add(firminfo);
                    }
                    workbook.Close();
                    System.IO.File.Delete(filePath);
                    return firmInfos;
                }

                return null;

            }
            catch (Exception e)
            {
                return null;
            }
        }
        public List<FirmMap> getFile2(HttpPostedFileBase postedFile)
        {
            try
            {
                string path = Server.MapPath("~/Uploads/");
                string filePath = string.Empty;
                if (postedFile != null)
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    filePath = path + DateTime.Now.Ticks + "-" + Path.GetFileName(postedFile.FileName);
                    postedFile.SaveAs(filePath);

                    //read data from excel
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(filePath);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;

                    List<FirmMap> firmMaps = new List<FirmMap>();
                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        FirmMap firmMap = new FirmMap();
                        firmMap.AssetClassID = ((Excel.Range)range.Cells[row, 1]).Text;
                        firmMap.AssetClassName = ((Excel.Range)range.Cells[row, 2]).Text;
                        firmMap.InterestedFirmsID = ((Excel.Range)range.Cells[row, 3]).Text;
                        firmMaps.Add(firmMap);
                    }
                    workbook.Close();
                    System.IO.File.Delete(filePath);
                    return firmMaps;
                }

                return null;

            }
            catch (Exception e)
            {
                return null;
            }
        }




    }
}