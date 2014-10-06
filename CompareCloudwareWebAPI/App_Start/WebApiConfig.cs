using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using CompareCloudwareWebAPI.Controllers;
using System.Web.Http.Tracing;
//using System.Net.Http;
using CompareCloudwareWebAPI.Helpers;

namespace CompareCloudwareWebAPI
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );

            config.Routes.MapHttpRoute(
                name: "GetSiteAnalytics",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { action = "GetSiteAnalytics", id = RouteParameter.Optional }
            );

            config.Routes.MapHttpRoute(
    name: "GetVendors",
    routeTemplate: "api/{controller}/{id}",
    defaults: new { action = "GetVendors", id = RouteParameter.Optional }
);

            config.Services.Replace(typeof(ITraceWriter), new SimpleTracer());
            //config.EnableSystemDiagnosticsTracing();
        }
    }
}
