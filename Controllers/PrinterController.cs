using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Rotativa1.Models;

// For more information on enabling MVC for empty projects, visit http://go.microsoft.com/fwlink/?LinkID=397860

namespace Rotativa1.Controllers
{
    

    public class PrinterController : RazorPDFCore.Controller
    {
        // GET: /<controller>/
        public IActionResult Pdf()
        {

            User2 user = new User2 { Name = "Ivan", Surname = "Males" };

            ViewBag.test = "test";

            return ViewPdf(user,"Pdf");
        }
    }
}
