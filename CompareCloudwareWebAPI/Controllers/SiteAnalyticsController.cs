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

namespace CompareCloudwareWebAPI.Controllers
{
    public class SiteAnalyticsController : ApiController
    {
        protected readonly ICompareCloudwareRepository _repository;
        protected readonly ICompareCloudwareContext _context;

        //public SiteAnalyticsController(ICustomSession session, ICompareCloudwareRepository repository, ISiteAnalyticsLogger _SiteAnalyticsLogger)
        public SiteAnalyticsController()
        {
            _context = new CompareCloudwareContext();
            _repository = new QueryRepository(_context);
        }

        public SiteAnalyticsController(ICompareCloudwareRepository repository)
        {
            _repository = repository;
            
        }



        public SiteAnalyticOutput[] GetSiteAnalytics()
        {
            try
            {
                List<SiteAnalyticOutput> siteAnalytics = _repository.GetAllSiteAnalytics();
                return siteAnalytics.ToArray();
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public SiteAnalyticOutput[] GetSiteAnalytics(string sessionID)
        {
            try
            {
                ExcelCreate eh = new ExcelCreate();
                eh.Main();
                List<SiteAnalyticOutput> siteAnalytics = _repository.GetAllSiteAnalyticsBySession(sessionID);

                return siteAnalytics.ToArray();
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public SiteAnalyticOutput[] xGet()
        {
            try
            {
                //List<SiteAnalyticModel> siteAnalytics = _repository.GetAllSiteAnalytics().Select(x => new SiteAnalyticModel()
                //{
                //    CategoryID = x.CategoryID,
                //    CloudApplicationID = x.CloudApplicationID,
                //    FeatureTypeID = x.FeatureTypeID,
                //    PersonID = x.PersonID,
                //    ReferenceDataRowID = x.ReferenceDataRowID,
                //    SessionID = x.SessionID,
                //    SiteAnalyticDate = x.SiteAnalyticDate,
                //    SiteAnalyticID = x.SiteAnalyticID,
                //    //SiteAnalyticType = new SiteAnalyticTypeModel()
                //    //    {
                //    //        AddDate = x.SiteAnalyticType.AddDate,
                //    //        LastUpdateDate = x.SiteAnalyticType.LastUpdateDate,
                //    //        SiteAnalyticTypeID = x.SiteAnalyticType.SiteAnalyticTypeID,
                //    //        SiteAnalyticTypeName = x.SiteAnalyticType.SiteAnalyticTypeName,
                //    //    },
                //    SiteAnalyticType = x.SiteAnalyticType.SiteAnalyticTypeName,
                //}).OrderByDescending(x => x.SiteAnalyticDate).ToList();
                List<SiteAnalyticOutput> siteAnalytics = _repository.GetAllSiteAnalytics();
                return siteAnalytics.ToArray();


            }
            catch (Exception e)
            {
                return null;
            }

        //    return new SiteAnalytic[]
        //{
        //    new SiteAnalytic
        //    {
        //     CategoryID = 1,
        //     CloudApplicationID = 2,
        //     FeatureTypeID = 3,
        //     PersonID = 4,
        //     ReferenceDataRowID = 5,
        //     SessionID = "qwertyuiop",
        //    },
        //    new SiteAnalytic
        //    {
        //     CategoryID = 6,
        //     CloudApplicationID = 7,
        //     FeatureTypeID = 8,
        //     PersonID = 9,
        //     ReferenceDataRowID = 10,
        //     SessionID = "asdfghjkl",
        //    },
        //};
        }

    }
}
