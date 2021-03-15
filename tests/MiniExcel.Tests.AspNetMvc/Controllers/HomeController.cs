using System;
using System.Web.Mvc;
using System.Collections.Generic;
using System.IO;
using MiniExcelLibs;


public class HomeController : Controller
{
    [HttpGet]
    public ActionResult Index() => View();

    public ActionResult Download()
    {
	   var values = new[] {
				new { Column1 = "MiniExcel", Column2 = 1 },
				new { Column1 = "Github", Column2 = 2}
			};
	   var stream = new MemoryStream();
	   stream.SaveAs(values);
	   return File(stream,
		   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		   "demo.xlsx");
    }
}
