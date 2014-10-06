using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace CompareCloudwareWebAPI
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );

            
            routes.MapRoute(
                name: "Analytics",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "SiteAnalytics", action = "GetSiteAnalytics", id = UrlParameter.Optional }
            );

            routes.MapRoute(
    name: "Vendors",
    url: "{controller}/{action}/{id}",
    defaults: new { controller = "Vendors", action = "GetVendors", id = UrlParameter.Optional }
);

            routes.MapRoute(
    name: "Analytics2",
    url: "{controller}/{action}/{id}",
    defaults: new { controller = "SiteAnalyticsVendorSummary", action = "GetSiteAnalyticsVendorSummary", id = UrlParameter.Optional }
);

           // config.Routes.MapHttpRoute(
           //name: "ActionApi",
           //routeTemplate: "api/{controller}/{action}/{id}",
           //defaults: new { id = RouteParameter.Optional }
           // );
        }
    }
}