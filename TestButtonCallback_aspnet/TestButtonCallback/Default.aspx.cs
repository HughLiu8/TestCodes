using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TestButtonCallback
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.Button1.Attributes.Add("onclick", "return onSubmit('test')"); 
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Response.Write("<script>alert('asp click')</script>");
        }
    }
}
