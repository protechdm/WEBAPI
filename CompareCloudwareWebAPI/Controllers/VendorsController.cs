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
using System.Web.Http.Tracing;

namespace CompareCloudwareWebAPI.Controllers
{
    public class VendorsController : ApiController
    {
        protected readonly ICompareCloudwareRepository _repository;
        protected readonly ICompareCloudwareContext _context;

        //public SiteAnalyticsController(ICustomSession session, ICompareCloudwareRepository repository, ISiteAnalyticsLogger _SiteAnalyticsLogger)
        public VendorsController()
        {
            _context = new CompareCloudwareContext();
            _repository = new QueryRepository(_context);
        }

        public VendorsController(ICompareCloudwareRepository repository)
        {
            _repository = repository;
            
        }




        public VendorModel[] GetVendors()
        {
            try
            {
              GlobalConfiguration.Configuration.Services.GetTraceWriter().Info(
            Request, "VendorsController", "Get the list of vendors.");

                IList<Vendor> vendors = _repository.GetAllVendors();
                return vendors.Select(x => new VendorModel()
                    {
                        VendorID = x.VendorID,
                        VendorName = x.VendorName,
                    }
                    )
                    .ToArray();
            }
            catch (Exception e)
            {
                return null;
            }
        }


    }
}
