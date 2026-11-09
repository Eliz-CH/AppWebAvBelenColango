using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using EntityLayer.PoliMusicV3;

namespace WebAppPoliMusicV3
{
    public partial class SiteMaster : MasterPage
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadUserName();
            }
        }

        /// <summary>
        /// Shows the name of the authenticated user in the navbar.
        /// Uses Session[USER].UserName if available, otherwise User.Identity.Name.
        /// </summary>
        private void LoadUserName()
        {
            EntityLayer.PoliMusicV3.User sessionUser = Session[EntityLayer.PoliMusicV3.Util.Constants.USER] as EntityLayer.PoliMusicV3.User;
            if (sessionUser != null && !string.IsNullOrEmpty(sessionUser.UserName))
            {
                lblUserName.Text = sessionUser.UserName;
            }
            else if (Page.User != null && Page.User.Identity.IsAuthenticated)
            {
                lblUserName.Text = Page.User.Identity.Name;
            }
            else
            {
                lblUserName.Text = "Guest";
            }
        }
    }
}
