using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Http;
using System.Web.Http.Tracing;
using System.Net.Http;
using System.IO;

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
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
    }

}
