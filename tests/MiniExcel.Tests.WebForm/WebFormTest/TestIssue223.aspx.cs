using System;
using System.Collections.Generic;
using System.Data;
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
            var path = HttpContext.Current.ApplicationInstance.Server.MapPath("~/TestIssue223.xlsx");
            var dt = MiniExcelLibs.MiniExcel.QueryAsDataTable(path);
            this.GridView1.DataSource = dt;
            this.GridView1.DataBind();
        }

        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}