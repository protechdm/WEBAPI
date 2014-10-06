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
using System.Net.Http.Headers;
//using System.Web.Mvc;

namespace CompareCloudwareWebAPI.Controllers
{
    public class SiteAnalyticsVendorSummaryExcelController : ApiController
    {
        protected readonly ICompareCloudwareRepository _repository;
        protected readonly ICompareCloudwareContext _context;

        //public SiteAnalyticsController(ICustomSession session, ICompareCloudwareRepository repository, ISiteAnalyticsLogger _SiteAnalyticsLogger)
        public SiteAnalyticsVendorSummaryExcelController()
        {
            _context = new CompareCloudwareContext();
            _repository = new QueryRepository(_context);
        }

        public SiteAnalyticsVendorSummaryExcelController(ICompareCloudwareRepository repository)
        {
            _repository = repository;
            
        }




        //public MemoryStream GetSiteAnalyticsVendorSummary(int vendorID, DateTime startDate, DateTime endDate)
        //{
        //    try
        //    {
        //        string vendorName = _repository.FindVendorByID(vendorID).VendorName;
        //        List<SiteAnalyticsVendorSummary> siteAnalytics = _repository.GetSiteAnalyticsForVendor(vendorID,startDate,endDate);

        //        ExcelCreate eh = new ExcelCreate();
        //        //eh.CreateVendorAnalyticsSummary(siteAnalytics,vendorName,startDate,endDate);
        //        MemoryStream ms = eh.CreateVendorAnalyticsSummaryAsStream(siteAnalytics, vendorName, startDate, endDate);
        //        return File(ms.GetBuffer(), "application/msexcel");
        //        return new ExcelResult(ms.GetBuffer(), "filename.xls");

        //        return ms;
        //    }
        //    catch (Exception e)
        //    {
        //        return null;
        //    }
        //}
        
        [HttpPost]
        public HttpResponseMessage GetSiteAnalyticsVendorSummary(int vendorID, DateTime startDate, DateTime endDate)
        //public FileStreamResult GetSiteAnalyticsVendorSummary(int vendorID, DateTime startDate, DateTime endDate)
        {
            string vendorName = _repository.FindVendorByID(vendorID).VendorName;
            List<SiteAnalyticsVendorSummary> siteAnalytics = _repository.GetSiteAnalyticsForVendor(vendorID, startDate, endDate);
            ExcelCreate eh = new ExcelCreate();
            MemoryStream ms = eh.CreateVendorAnalyticsSummaryAsStream(siteAnalytics, vendorName, startDate, endDate);
            ms.Position = 0;
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
            result.Content = new StreamContent(ms);
            //result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");
            //result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/force-download");

            result.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");

            string fileName = vendorName + "_" +
                startDate.Day.ToString() + "-" + startDate.Month.ToString() + "-" + startDate.Year.ToString() +
                "_to_" +
                endDate.Day.ToString() + "-" + endDate.Month.ToString() + "-" + endDate.Year.ToString()
                + ".xlsx"
                ;

            result.Content.Headers.ContentDisposition.FileName = fileName;

            return result;

            
        }

    }


    #region CRAP
    ///// <summary>
    ///// Custom ActionResult for saving excel files
    ///// </summary>
    //public class ExcelResult : ActionResult
    //{
    //    private Stream excelStream;
    //    private String fileName;
    //    private bool saveAsXML;

    //    /// <summary>
    //    /// Creates a new ActionResult for saving excel files
    //    /// </summary>
    //    /// <param name="excel">byte array from excel workbook</param>
    //    /// <param name="fileName">string defining file name</param>
    //    public ExcelResult(byte[] excel, String fileName)
    //    {
    //        excelStream = new MemoryStream(excel);
    //        this.fileName = fileName;
    //        saveAsXML = false;
    //    }

    //    /// <summary>
    //    /// Creates a new ActionResult for saving excel files
    //    /// </summary>
    //    /// <param name="excel">byte array from excel workbook</param>
    //    /// <param name="fileName">string defining file name</param>
    //    /// <param name="saveAsXML">defines the content type as XML</param>
    //    public ExcelResult(byte[] excel, String fileName, bool saveAsXML)
    //    {
    //        excelStream = new MemoryStream(excel);
    //        this.fileName = fileName;
    //        this.saveAsXML = saveAsXML;
    //    }

    //    public override void ExecuteResult(ControllerContext context)
    //    {
    //        if (context == null)
    //        {
    //            throw new ArgumentNullException("context");
    //        }

    //        HttpResponseBase response = context.HttpContext.Response;

    //        response.ContentType = (saveAsXML) ? "text/xml" : "application/vnd.ms-excel";

    //        response.AddHeader("content-disposition", "attachment; filename=" + fileName);

    //        byte[] buffer = new byte[4096];

    //        while (true)
    //        {
    //            int read = this.excelStream.Read(buffer, 0, buffer.Length);
    //            if (read == 0)
    //            {
    //                break;
    //            }

    //            response.OutputStream.Write(buffer, 0, read);
    //        }

    //        response.End();
    //    }

    //}
    #endregion
}
