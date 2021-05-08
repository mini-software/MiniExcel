using MiniExcelLibs;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebFormTest
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            var fileName = "Demo.xlsx";
            var sheetName = "Sheet1";
            HttpResponse response = HttpContext.Current.Response;
            response.Clear();
            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            response.AddHeader("Content-Disposition", $"attachment;filename=\"{fileName}\"");
            var values = new[] {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            };
            var memoryStream = new MemoryStream();
            memoryStream.SaveAs(values, sheetName: sheetName);
            memoryStream.Seek(0, SeekOrigin.Begin);
            memoryStream.CopyTo(Response.OutputStream);
            response.End();
        }


        protected void Button2_Click(object sender, EventArgs e)
        {
            var path = HttpContext.Current.ApplicationInstance.Server.MapPath("~/TestIssue223.xlsx");
            var dt = MiniExcelLibs.MiniExcel.QueryAsDataTable(path);
            this.GridView1.DataSource = dt;
            this.GridView1.DataBind();
        }
    }
}