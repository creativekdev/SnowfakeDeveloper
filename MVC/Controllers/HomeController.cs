using Microsoft.Office.Interop.Excel;
using MVC.Models;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;
using System.Web.UI.WebControls;
using ExcelDataReader;

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
                        if (file.FileName.Contains("Firm Info"))
                        {
                            if (Path.GetExtension(file.FileName).ToUpper().Equals(".TXT") || Path.GetExtension(file.FileName).ToUpper().Equals(".CSV")) firmInfos = getFile1FromTxt(file);
                            else if(Path.GetExtension(file.FileName).ToUpper().Equals(".XLSX") || Path.GetExtension(file.FileName).ToUpper().Equals(".XLS")) firmInfos = getFile1(file);
                        }
                        else if (file.FileName.Contains("Asset Class - Firm Map"))
                        {
                            if (Path.GetExtension(file.FileName).ToUpper().Equals(".TXT") || Path.GetExtension(file.FileName).ToUpper().Equals(".CSV")) firmMaps = getFile2FromTxt(file);                            
                            else if(Path.GetExtension(file.FileName).ToUpper().Equals(".XLSX") || Path.GetExtension(file.FileName).ToUpper().Equals(".XLS")) firmMaps = getFile2(file);
                        }
                    }

                }

                if (firmInfos != null && firmMaps != null && firmInfos.Count > 0 && firmMaps.Count > 0)
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
                    return View(nameof(Index));

                }
                else
                {
                    TempData["message"] = "No file was uploaded";
                    ViewBag.Message = "Please select correct files";
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
                    List<FirmInfo> firmInfos = new List<FirmInfo>();

                    using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {

                            Boolean isTitle = true;
                            while (reader.Read()) //Each row of the file
                            {
                                if (isTitle)
                                {
                                    isTitle = false;
                                    continue;
                                }
                                FirmInfo firminfo = new FirmInfo();
                                firminfo.FirmID = reader.GetValue(0).ToString();
                                firminfo.FirmName = reader.GetValue(1).ToString();
                                firmInfos.Add(firminfo);
                            }
                        }
                    }

                    System.IO.File.Delete(filePath);
                    return firmInfos;

                }
                else return null;
            }
            catch (Exception e)
            {
                return null;
            }
        }
        public List<FirmInfo> getFile1FromTxt(HttpPostedFileBase postedFile)
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

                    //read data from txt
                    List<FirmInfo> firmInfos = new List<FirmInfo>();
                    using (var reader = new StreamReader(filePath))
                    {
                        bool flag = false;
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            if (!flag)
                            {
                                flag = true;
                                continue;
                            }
                            var temp = line.Split(',');
                            FirmInfo firminfo = new FirmInfo();
                            firminfo.FirmID = temp[0];
                            firminfo.FirmName = temp[1];
                            firmInfos.Add(firminfo);
                        }
                    }
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
                    List<FirmMap> firmMaps = new List<FirmMap>();
                    using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {

                            Boolean isTitle = true;
                            while (reader.Read()) //Each row of the file
                            {
                                if (isTitle)
                                {
                                    isTitle = false;
                                    continue;
                                }
                                FirmMap firmMap = new FirmMap();
                                firmMap.AssetClassID = reader.GetValue(0).ToString();
                                firmMap.AssetClassName = reader.GetValue(1).ToString();
                                firmMap.InterestedFirmsID = reader.GetValue(2).ToString();
                                firmMaps.Add(firmMap);
                            }
                        }
                    }

                    System.IO.File.Delete(filePath);
                    return firmMaps;


                }
                else return null;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public List<FirmMap> getFile2FromTxt(HttpPostedFileBase postedFile)
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

                    //read data from txt
                    List<FirmMap> firmMaps = new List<FirmMap>();
                    using (var reader = new StreamReader(filePath))
                    {
                        bool flag = false;
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            if (!flag)
                            {
                                flag = true;
                                continue;
                            }
                            var temp = line.Split(',');
                            FirmMap firmMap = new FirmMap();
                            firmMap.AssetClassID = temp[0];
                            firmMap.AssetClassName = temp[1];
                            firmMap.InterestedFirmsID = temp[2];
                            firmMaps.Add(firmMap);
                        }
                    }


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