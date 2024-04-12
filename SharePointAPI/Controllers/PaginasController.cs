using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SharePointAPI.Controllers
{
    public class PaginasController : Controller
    {
        // GET: Paginas
        public ActionResult Index()
        {
            return View();
        }
    }
}