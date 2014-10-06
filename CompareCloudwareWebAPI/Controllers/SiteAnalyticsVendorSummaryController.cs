using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CompareCloudware.Domain.Models;
using CompareCloudware.Domain.Contracts.Repositories;
using CompareCloudware.POCOQueryRepository;
using CompareCloudwareWebAPI.Models;
using CompareCloudwareWebAPI.Helpers;
using System.IO;

namespace CompareCloudwareWebAPI.Controllers
{
    public class SiteAnalyticsVendorSummaryController : ApiController
    {
        protected readonly ICompareCloudwareRepository _repository;
        protected readonly ICompareCloudwareContext _context;

        //public SiteAnalyticsController(ICustomSession session, ICompareCloudwareRepository repository, ISiteAnalyticsLogger _SiteAnalyticsLogger)
        public SiteAnalyticsVendorSummaryController()
        {
            _context = new CompareCloudwareContext();
            _repository = new QueryRepository(_context);
        }

        public SiteAnalyticsVendorSummaryController(ICompareCloudwareRepository repository)
        {
            _repository = repository;
            
        }




        public SiteAnalyticsVendorSummary[] GetSiteAnalyticsVendorSummary(int vendorID, DateTime startDate, DateTime endDate)
        {
            try
            {
                string vendorName = _repository.FindVendorByID(vendorID).VendorName;
                List<SiteAnalyticsVendorSummary> siteAnalytics = _repository.GetSiteAnalyticsForVendor(vendorID,startDate,endDate);

                //ExcelCreate eh = new ExcelCreate();
                //eh.CreateVendorAnalyticsSummary(siteAnalytics,vendorName,startDate,endDate);
                //MemoryStream ms = eh.CreateVendorAnalyticsSummaryAsStream(siteAnalytics, vendorName, startDate, endDate);

                return siteAnalytics.ToArray();
            }
            catch (Exception e)
            {
                return null;
            }
        }


    }
}
