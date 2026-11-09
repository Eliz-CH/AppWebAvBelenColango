using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV3
{
    public partial class Song : System.Web.UI.Page
    {
        private string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                ValidateSession();
                LoadSongs();
            }
        }

        /// <summary>
        /// Validates that the user is authenticated. If Session[USER] does not exist
        /// and the user identity is not available, redirect to Login.aspx.
        /// </summary>
        private void ValidateSession()
        {
            EntityLayer.PoliMusicV3.User sessionUser = Session[EntityLayer.PoliMusicV3.Util.Constants.USER] as EntityLayer.PoliMusicV3.User;
            if (sessionUser == null && (User == null || !User.Identity.IsAuthenticated))
            {
                Response.Redirect("~/Login.aspx", true);
            }
        }

        /// <summary>
        /// Loads the list of songs from BLL.PoliMusicV3.Song.Read()
        /// and binds it to the GridView.
        /// </summary>
        private void LoadSongs()
        {
            BLL.PoliMusicV3.Song songBLL = new BLL.PoliMusicV3.Song(strConnString);
            DataTable dtSongs = songBLL.Read();
            gvSongs.DataSource = dtSongs;
            gvSongs.DataBind();
        }

        /// <summary>
        /// Builds the URL used by the HTML5 &lt;audio&gt; element for a song.
        ///
        /// The [SONG_PATH] column already stores an application-relative path such as
        /// "~/songFiles/bensound-adventure.mp3". Prepending the song folder again would
        /// produce an invalid duplicated path ("~/songFiles/~/songFiles/..."), which is why
        /// the value is only normalized here:
        ///   - absolute URL (http/https)  -> used as is
        ///   - application-relative (~/)   -> resolved against the application root
        ///   - bare file name (x.mp3)      -> resolved against the ~/songFiles folder
        /// </summary>
        protected string ResolveSongPath(string songPath)
        {
            if (string.IsNullOrEmpty(songPath))
            {
                return ResolveUrl("~/songFiles/");
            }

            songPath = songPath.Trim();

            // Already an absolute URL - keep it untouched.
            if (songPath.StartsWith("http://") || songPath.StartsWith("https://"))
            {
                return songPath;
            }

            // Application-relative path stored in the database (e.g. "~/songFiles/song.mp3").
            if (songPath.StartsWith("~/"))
            {
                return ResolveUrl(songPath);
            }

            // Root-relative path (e.g. "/songFiles/song.mp3").
            if (songPath.StartsWith("/"))
            {
                return ResolveUrl("~" + songPath);
            }

            // Legacy relative path (e.g. "../songFiles/song.mp3").
            if (songPath.StartsWith("../"))
            {
                return ResolveUrl("~/" + songPath.Substring(3));
            }

            // Bare file name - assume it lives in the ~/songFiles folder.
            return ResolveUrl("~/songFiles/" + Path.GetFileName(songPath));
        }

        /// <summary>
        /// Handles the RowDataBound event. Reserved for future enhancements
        /// (e.g. incrementing the play counter when a track starts).
        /// </summary>
        protected void gvSongs_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                // The audio element src is resolved via ResolveSongPath in the ASPX.
            }
        }
    }
}
