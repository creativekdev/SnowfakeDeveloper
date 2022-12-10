using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Xml.Linq;

namespace MVC.Models
{
    public class FirmData
    {
        public Dictionary<string, string> Firms { get; set; }
        public Dictionary<string, string> Assets { get; set; }
        public Dictionary<string, Dictionary<string, string>> Relations { get; set; }
    }
    public class FirmInfo
    {
        public string FirmID { get; set; }
        public string FirmName { get; set; }
    }

    public class FirmMap
    {
        public string AssetClassID { get; set; }
        public string AssetClassName { get; set; }
        public string InterestedFirmsID { get; set; }
    }
    public class FileModel
    {
        [Required(ErrorMessage = "Please select file.")]
        [Display(Name = "Browse File")]
        public HttpPostedFileBase[] files { get; set; }
    }
}