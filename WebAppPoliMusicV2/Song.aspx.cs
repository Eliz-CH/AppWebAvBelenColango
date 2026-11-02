using System;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        private string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                ValidateSession();
                LoadUserName();
                LoadSongs();
            }
        }

        /// <summary>
        /// Validates that the user is authenticated. If Session[USER] does not exist
        /// and the user identity is not available, redirect to Login.aspx.
        /// </summary>
        private void ValidateSession()
        {
            Entities.User sessionUser = Session[Util.Constants.USER] as Entities.User;
            if (sessionUser == null && (User == null || !User.Identity.IsAuthenticated))
            {
                Response.Redirect("Login.aspx", true);
            }
        }

        /// <summary>
        /// Loads the authenticated user's name into the welcome label.
        /// Uses Session[USER].UserName if available, otherwise User.Identity.Name.
        /// </summary>
        private void LoadUserName()
        {
            Entities.User sessionUser = Session[Util.Constants.USER] as Entities.User;
            if (sessionUser != null && !string.IsNullOrEmpty(sessionUser.UserName))
            {
                lblUserName.Text = sessionUser.UserName;
            }
            else if (User != null && User.Identity.IsAuthenticated)
            {
                lblUserName.Text = User.Identity.Name;
            }
            else
            {
                lblUserName.Text = "Guest";
            }
        }

        /// <summary>
        /// Loads the list of songs from BLL.Song.Read() and binds it to the GridView.
        /// </summary>
        private void LoadSongs()
        {
            BLL.Song songBLL = new BLL.Song(strConnString);
            DataTable dtSongs = songBLL.Read();
            gvSongs.DataSource = dtSongs;
            gvSongs.DataBind();
        }

        /// <summary>
        /// Handles the RowDataBound event to inject the audio source path
        /// into the HTML5 <audio> element for each song row.
        /// </summary>
        protected void gvSongs_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DataRowView rowView = e.Row.DataItem as DataRowView;
                if (rowView != null)
                {
                    string songPath = rowView["SONG_PATH"].ToString();

                    // If the path stored in DB is relative (e.g. ../songFiles/name.mp3),
                    // resolve it from the app setting as a fallback.
                    if (string.IsNullOrEmpty(songPath))
                    {
                        string songName = rowView["SONG_NAME"].ToString();
                        string songFolder = ConfigurationManager.AppSettings["songFolderPath"];
                        songPath = songFolder + songName + ".mp3";
                    }

                    // Ensure the path is usable from the browser
                    if (!songPath.StartsWith("http") && !songPath.StartsWith("/"))
                    {
                        // Normalize: remove leading "../" if present, prepend the application root
                        songPath = ResolveClientUrl("~/songFiles/" + System.IO.Path.GetFileName(songPath));
                    }

                    // Find the <audio> control in the last column and set the source
                    System.Web.UI.HtmlControls.HtmlGenericControl audioCtrl =
                        e.Row.Cells[e.Row.Cells.Count - 1].FindControl("audio") as System.Web.UI.HtmlControls.HtmlGenericControl;
                    // The <audio> element is not a server control, so we set the <source> src via literal
                    // We use a Literal control approach instead

                    // Since the audio element is plain HTML (not runat=server), we inject via inner HTML
                    // of the cell's last child by finding the source element
                    // The approach: replace the cell content with the full audio tag
                    string audioHtml = "<audio controls preload=\"none\">" +
                        "<source src=\"" + songPath + "\" type=\"audio/mpeg\" />" +
                        "Your browser does not support the audio element." +
                        "</audio>";
                    e.Row.Cells[e.Row.Cells.Count - 1].Text = audioHtml;
                }
            }
        }
    }
}
