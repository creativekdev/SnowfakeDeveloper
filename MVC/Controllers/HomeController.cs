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
                    //                   postedFile.SaveAs(filePath);


                    //Coneection String by default empty
                    string ConStr = filePath;
                    //Extantion of the file upload control saving into ext because 
                    //there are two types of extation .xls and .xlsx of excel 
                    string ext = Path.GetExtension(filePath).ToLower();
                    //saving the file inside the MyFolder of the server
                    postedFile.SaveAs(filePath);
                    // Label1.Text = FileUpload1.FileName + "\'s Data showing into the GridView";
                    //checking that extantion is .xls or .xlsx

                    if (ext.Trim() == ".xls")
                    {
                        //connection string for that file which extantion is .xls
                        ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }
                    else if (ext.Trim() == ".xlsx")
                    {
                        //connection string for that file which extantion is .xlsx
                        ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }
                    //making query
                    string query = "SELECT * FROM [Sheet1$]";
                    //Providing connection
                    OleDbConnection conn = new OleDbConnection(ConStr);
                    //checking that connection state is closed or not if closed the 
                    //open the connection
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }
                    //create command object
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    // create a data adapter and get the data into dataadapter
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    //fill the excel data to data set
                    da.Fill(ds);
                    List<FirmInfo> firmInfos = new List<FirmInfo>();

                    if (ds.Tables != null && ds.Tables.Count > 0)
                    {
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            FirmInfo firminfo = new FirmInfo();
                            firminfo.FirmID = ds.Tables[0].Rows[i][0].ToString();
                            firminfo.FirmName = ds.Tables[0].Rows[i][1].ToString();
                            firmInfos.Add(firminfo);

                        }
                    }

                    conn.Close();
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
                    //                   postedFile.SaveAs(filePath);


                    //Coneection String by default empty
                    string ConStr = filePath;
                    //Extantion of the file upload control saving into ext because 
                    //there are two types of extation .xls and .xlsx of excel 
                    string ext = Path.GetExtension(filePath).ToLower();
                    //saving the file inside the MyFolder of the server
                    postedFile.SaveAs(filePath);
                    // Label1.Text = FileUpload1.FileName + "\'s Data showing into the GridView";
                    //checking that extantion is .xls or .xlsx

                    if (ext.Trim() == ".xls")
                    {
                        //connection string for that file which extantion is .xls
                        ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }
                    else if (ext.Trim() == ".xlsx")
                    {
                        //connection string for that file which extantion is .xlsx
                        ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }
                    //making query
                    string query = "SELECT * FROM [Sheet1$]";
                    //Providing connection
                    OleDbConnection conn = new OleDbConnection(ConStr);
                    //checking that connection state is closed or not if closed the 
                    //open the connection
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }
                    //create command object
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    // create a data adapter and get the data into dataadapter
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    //fill the excel data to data set
                    da.Fill(ds);
                    List<FirmMap> firmMaps = new List<FirmMap>();
                    if (ds.Tables != null && ds.Tables.Count > 0)
                    {
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            FirmMap firmMap = new FirmMap();
                            firmMap.AssetClassID = ds.Tables[0].Rows[i][0].ToString();
                            firmMap.AssetClassName = ds.Tables[0].Rows[i][1].ToString();
                            firmMap.InterestedFirmsID = ds.Tables[0].Rows[i][2].ToString();
                            firmMaps.Add(firmMap);

                        }
                    }

                    conn.Close();
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