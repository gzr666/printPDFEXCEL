using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.NodeServices;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using OfficeOpenXml;

namespace Rotativa1.Controllers
{

    public class User {

        public string Name { get; set; }
        public string Street { get; set; }


    }
    public class HomeController : Controller
    {
        private readonly IHostingEnvironment _env;

        public HomeController(IHostingEnvironment env)
        {
            _env = env;


        }


        public IActionResult Index()
        {
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        [HttpGet]
        public async Task<IActionResult> MyAction([FromServices] INodeServices nodeServices)
        {
            List<User> mylist = new List<User> {

                new User { Name = "Ivan", Street = "Brace RADIĆ" },
                new User { Name = "Šumar", Street = "Brace RADIČ" }

                };

            var result = await nodeServices.InvokeAsync<byte[]>("./pdf",mylist);

            HttpContext.Response.ContentType = "application/pdf";

            string filename = @"report.pdf";
            HttpContext.Response.Headers.Add("x-filename", filename);
            HttpContext.Response.Headers.Add("Access-Control-Expose-Headers", "x-filename");
            HttpContext.Response.Body.Write(result, 0, result.Length);
            return new ContentResult();
        }

        [HttpGet]
        [Route("home/single/{id}")]
        public async Task<IActionResult> MyActionS([FromServices] INodeServices nodeServices,[FromRoute] int id)
        {


            User user = new User { Name = "Ivan", Street = "Brace RADIĆ" };
               

           

            var result = await nodeServices.InvokeAsync<byte[]>("./pdfSingle", user);

            HttpContext.Response.ContentType = "application/pdf";

            string filename = @"report.pdf";
            HttpContext.Response.Headers.Add("x-filename", filename);
            HttpContext.Response.Headers.Add("Access-Control-Expose-Headers", "x-filename");
            HttpContext.Response.Body.Write(result, 0, result.Length);
            return new ContentResult();
        }

        [HttpGet]
        public String Excel()
        {
            string webRoot = _env.WebRootPath;
            string filename = @"demo1.xlsx";
            string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, filename);
            FileInfo file = new FileInfo(Path.Combine(webRoot, filename));

            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(webRoot, filename));
            }

            using (ExcelPackage package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Employee");
                //First add the headers
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Gender";
                worksheet.Cells[1, 4].Value = "Salary (in $)";

                //Add values
                worksheet.Cells["A2"].Value = 1000;
                worksheet.Cells["B2"].Value = "Jon";
                worksheet.Cells["C2"].Value = "M";
                worksheet.Cells["D2"].Value = 5000;

                worksheet.Cells["A3"].Value = 1001;
                worksheet.Cells["B3"].Value = "Graham";
                worksheet.Cells["C3"].Value = "M";
                worksheet.Cells["D3"].Value = 10000;

                worksheet.Cells["A4"].Value = 1002;
                worksheet.Cells["B4"].Value = "Jenny";
                worksheet.Cells["C4"].Value = "F";
                worksheet.Cells["D4"].Value = 5000;

                package.Save(); //Save the workbook.
            }

            return URL;




        }



        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Error()
        {
            return View();
        }
    }
}
