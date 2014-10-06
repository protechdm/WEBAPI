using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
//using System.Web.Mvc;
using System.Web.Http;
using System.Web.Http.Tracing;
using System.Net.Http;
using System.IO;

using System.Net;
using CompareCloudware.Domain.Models;
using CompareCloudware.Domain.Contracts.Repositories;
using CompareCloudware.POCOQueryRepository;
using CompareCloudwareWebAPI.Models;
using CompareCloudwareWebAPI.Helpers;

namespace CompareCloudwareWebAPI.Helpers
{
    public class Tracing
    {
    }

    public class SimpleTracer : ITraceWriter
    {
        public void Trace(HttpRequestMessage request, string category, TraceLevel level,
            Action<TraceRecord> traceAction)
        {
            TraceRecord rec = new TraceRecord(request, category, level);
            traceAction(rec);
            WriteTrace(rec);
        }

        protected void WriteTrace(TraceRecord rec)
        {
            var message = string.Format("{0};{1};{2}",
                rec.Operator, rec.Operation, rec.Message);
            System.Diagnostics.Trace.WriteLine(message, rec.Category);

            string path = HttpContext.Current.Server.MapPath("~/Logs/MyTestLog.txt");
            File.AppendAllText(path, rec.Status + " - " + rec.Message + "\r\n");
        }
    }

}