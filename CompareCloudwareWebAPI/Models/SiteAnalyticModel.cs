using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CompareCloudwareWebAPI.Models
{
    public class SiteAnalyticModel
    {
        public int SiteAnalyticID { get; set; }
        public DateTime SiteAnalyticDate { get; set; }
        public string SiteAnalyticType { get; set; }
        public int? CloudApplicationID { get; set; }
        public int? CategoryID { get; set; }
        public int? PersonID { get; set; }
        public string SessionID { get; set; }
        public int? ReferenceDataRowID { get; set; }
        public int? FeatureTypeID { get; set; }
    }
}