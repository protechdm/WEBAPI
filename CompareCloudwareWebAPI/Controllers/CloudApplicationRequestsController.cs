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
    public class CloudApplicationRequestsController : ApiController
    {
        protected readonly ICompareCloudwareRepository _repository;
        protected readonly ICompareCloudwareContext _context;

        //public SiteAnalyticsController(ICustomSession session, ICompareCloudwareRepository repository, ISiteAnalyticsLogger _SiteAnalyticsLogger)
        public CloudApplicationRequestsController()
        {
            _context = new CompareCloudwareContext();
            _repository = new QueryRepository(_context);
        }

        public CloudApplicationRequestsController(ICompareCloudwareRepository repository)
        {
            _repository = repository;
            
        }




        public CloudApplicationRequestModel[] GetCloudApplicationRequests(DateTime startDate, DateTime endDate)
        {
            try
            {
                List<CloudApplicationRequestModel> cloudApplicationRequests = _repository.GetWEBAPICloudApplicationRequests(startDate, endDate).Select(x => new CloudApplicationRequestModel()
                    {
                        Brand = x.Brand,   
                        CloudApplicationID = x.CloudApplicationID,   
                        CloudApplicationRequestID = x.CloudApplicationRequestID,   
                        Company = x.Company,   
                        EMail = x.EMail,   
                        Forename = x.Forename,   
                        NumberOfEmployees = x.NumberOfEmployees,   
                        PersonAddress1 = x.PersonAddress1,   
                        PersonAddress2 = x.PersonAddress2,   
                        PersonCountry = x.PersonCountry,   
                        PersonID = x.PersonID,   
                        PersonPostCode = x.PersonPostCode,   
                        PersonRegion = x.PersonRegion,   
                        Position = x.Position,   
                        RequestTypeID = x.RequestTypeID,   
                        Serviced = x.Serviced,   
                        ServiceName = x.ServiceName,   
                        Servicing = x.Servicing,   
                        Surname = x.Surname,   
                        Telephone = x.Telephone,   
                        UserName = x.UserName,   
                        VendorName = x.VendorName,
                        RequestType = x.RequestTypeID == 1 ? "TRIAL" : "BUY",
                    }
                    ).ToList();

                //ExcelCreate eh = new ExcelCreate();
                //eh.CreateVendorAnalyticsSummary(siteAnalytics,vendorName,startDate,endDate);
                //MemoryStream ms = eh.CreateVendorAnalyticsSummaryAsStream(siteAnalytics, vendorName, startDate, endDate);

                return cloudApplicationRequests.ToArray();
            }
            catch (Exception e)
            {
                return null;
            }
        }


    }
}
