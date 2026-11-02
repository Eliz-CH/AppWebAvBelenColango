<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>PoliMusic - Song Library</title>
    <style type="text/css">
        /* ===== Reset & Base ===== */
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(160deg, #1a0b2e 0%, #2d1b4e 40%, #1a0b2e 100%);
            color: #E6E6FA;
            min-height: 100vh;
        }

        /* ===== Header / Navbar ===== */
        .navbar {
            background: rgba(30, 15, 50, 0.85);
            backdrop-filter: blur(10px);
            border-bottom: 1px solid rgba(147, 112, 219, 0.3);
            padding: 18px 40px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            position: sticky;
            top: 0;
            z-index: 100;
        }

        .navbar-brand {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 22px;
            font-weight: 700;
            color: #9370DB;
        }

        .navbar-brand .logo-icon {
            width: 32px;
            height: 32px;
            background: linear-gradient(135deg, #9370DB, #4B0082);
            border-radius: 50%;
            display: inline-block;
            box-shadow: 0 0 12px rgba(147, 112, 219, 0.5);
        }

        .navbar-right {
            display: flex;
            align-items: center;
            gap: 24px;
        }

        .welcome-text {
            color: #E6E6FA;
            font-size: 15px;
            font-weight: 400;
        }

        .welcome-text strong {
            color: #9370DB;
            font-weight: 600;
        }

        .logout-link {
            text-decoration: none;
            color: #E6E6FA;
            font-size: 14px;
            padding: 8px 20px;
            border: 1px solid rgba(147, 112, 219, 0.4);
            border-radius: 22px;
            background: rgba(75, 0, 130, 0.15);
            transition: all 0.3s ease;
            display: inline-block;
        }

        .logout-link:hover {
            background: rgba(147, 112, 219, 0.25);
            border-color: #9370DB;
            color: #9370DB;
            box-shadow: 0 0 10px rgba(147, 112, 219, 0.3);
        }

        /* ===== Main Container ===== */
        .main-container {
            max-width: 1000px;
            margin: 40px auto;
            padding: 0 20px;
        }

        .page-title {
            font-size: 32px;
            font-weight: 700;
            color: #E6E6FA;
            margin-bottom: 8px;
            text-align: center;
        }

        .page-subtitle {
            font-size: 15px;
            color: #9370DB;
            text-align: center;
            margin-bottom: 30px;
            font-weight: 300;
        }

        /* ===== GridView / Table Container ===== */
        .table-container {
            background: rgba(30, 15, 50, 0.6);
            border-radius: 16px;
            border: 1px solid rgba(147, 112, 219, 0.15);
            overflow: hidden;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4);
        }

        /* The actual GridView rendered as table */
        .table-container table {
            width: 100%;
            border-collapse: collapse;
        }

        .table-container th {
            background: rgba(75, 0, 130, 0.4);
            color: #E6E6FA;
            font-size: 13px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 16px 14px;
            text-align: center;
            border-bottom: 2px solid rgba(147, 112, 219, 0.3);
        }

        .table-container th:first-child { width: 80px; }
        .table-container th:nth-child(4) { width: 100px; }
        .table-container th:last-child { width: 280px; }

        .table-container td {
            padding: 14px;
            font-size: 14px;
            color: #E6E6FA;
            border-bottom: 1px solid rgba(147, 112, 219, 0.08);
            text-align: center;
            vertical-align: middle;
        }

        .table-container tr:nth-child(even) td {
            background: rgba(147, 112, 219, 0.04);
        }

        .table-container tr:hover td {
            background: rgba(147, 112, 219, 0.12);
            transition: background 0.2s ease;
        }

        .table-container tr:last-child td {
            border-bottom: none;
        }

        /* ===== Song Name cell (left aligned) ===== */
        .song-name-cell {
            text-align: left !important;
            font-weight: 500;
            color: #E6E6FA;
        }

        /* ===== Plays badge ===== */
        .plays-badge {
            display: inline-block;
            min-width: 40px;
            padding: 4px 12px;
            border-radius: 14px;
            background: rgba(147, 112, 219, 0.15);
            color: #9370DB;
            font-weight: 600;
            font-size: 13px;
        }

        /* ===== Audio player ===== */
        .table-container audio {
            width: 250px;
            height: 36px;
            border-radius: 20px;
        }

        /* ===== Empty data message ===== */
        .empty-message {
            text-align: center;
            padding: 60px 20px;
            color: rgba(230, 230, 250, 0.5);
            font-size: 16px;
        }

        /* ===== Footer ===== */
        .footer {
            text-align: center;
            padding: 20px;
            color: rgba(230, 230, 250, 0.3);
            font-size: 12px;
            margin-top: 20px;
        }

        /* ===== Responsive ===== */
        @media screen and (max-width: 768px) {
            .navbar { padding: 14px 18px; }
            .navbar-brand { font-size: 18px; }
            .welcome-text { font-size: 13px; }
            .main-container { padding: 0 12px; }
            .page-title { font-size: 24px; }
            .table-container { overflow-x: auto; border-radius: 12px; }
            .table-container th, .table-container td { padding: 10px 8px; font-size: 12px; }
            .table-container audio { width: 180px; }
        }
    </style>
</head>
<body>
    <form id="formSong" runat="server">
        <!-- ===== Navbar ===== -->
        <div class="navbar">
            <div class="navbar-brand">
                <span class="logo-icon"></span>
                <span>PoliMusic</span>
            </div>
            <div class="navbar-right">
                <span class="welcome-text">Welcome <strong><asp:Label ID="lblUserName" runat="server"></asp:Label></strong></span>
                <asp:HyperLink ID="lnkLogout" runat="server" NavigateUrl="~/Logout.aspx" CssClass="logout-link">Logout</asp:HyperLink>
            </div>
        </div>

        <!-- ===== Main Content ===== -->
        <div class="main-container">
            <h1 class="page-title">Song Library</h1>
            <p class="page-subtitle">Browse and play your favorite tracks</p>

            <div class="table-container">
                <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False"
                    DataKeyNames="ID_SONG" GridLines="None" BorderStyle="None"
                    EmptyDataText="No songs available in the library."
                    OnRowDataBound="gvSongs_RowDataBound">
                    <Columns>
                        <asp:BoundField DataField="ID_SONG" HeaderText="ID Cancion" ReadOnly="True" />
                        <asp:BoundField DataField="SONG_NAME" HeaderText="Nombre" ReadOnly="True" />
                        <asp:BoundField DataField="SONG_PATH" HeaderText="Ruta" ReadOnly="True" />
                        <asp:BoundField DataField="PLAYS" HeaderText="Reproducciones" ReadOnly="True" />
                        <asp:TemplateField HeaderText="Reproducir">
                            <ItemTemplate>
                                <audio controls preload="none">
                                    <source type="audio/mpeg" />
                                </audio>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>

            <div class="footer">
                &copy; 2026 PoliMusic - All rights reserved
            </div>
        </div>
    </form>
</body>
</html>
