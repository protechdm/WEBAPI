using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CompareCloudwareWebAPI.Models
{
    public class SiteAnalyticTypeModel
    {
        public int SiteAnalyticTypeID { get; set; }
        public string SiteAnalyticTypeName { get; set; }
        public DateTime AddDate { get; set; }
        public DateTime? LastUpdateDate { get; set; }
    }
}