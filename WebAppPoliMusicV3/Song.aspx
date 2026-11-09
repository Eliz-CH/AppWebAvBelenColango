<%@ Page Title="Canciones" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV3.Song" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <style type="text/css">
        /* ===== Song page - PoliMusic dark lavender theme ===== */
        .song-page {
            background: linear-gradient(160deg, #1a0b2e 0%, #2d1b4e 45%, #1a0b2e 100%);
            min-height: calc(100vh - 120px);
            margin: -15px;
            padding: 0 0 40px 0;
        }

        .song-container {
            max-width: 940px;
            margin: 0 auto;
            padding: 40px 20px 0 20px;
        }

        /* ===== Titles ===== */
        .song-page-title {
            font-size: 30px;
            font-weight: 700;
            color: #E6E6FA;
            text-align: center;
            margin: 0 0 6px 0;
            letter-spacing: 0.3px;
        }

        .song-page-subtitle {
            font-size: 13px;
            font-weight: 300;
            color: #9370DB;
            text-align: center;
            margin: 0 0 28px 0;
            opacity: 0.85;
        }

        /* ===== Table container ===== */
        .song-table-wrap {
            background: rgba(30, 15, 50, 0.55);
            border: 1px solid rgba(147, 112, 219, 0.12);
            border-radius: 14px;
            overflow: hidden;
            box-shadow: 0 10px 34px rgba(0, 0, 0, 0.45);
        }

        /* ===== GridView overrides (Bootstrap .table on top) ===== */
        .song-table-wrap .table {
            margin-bottom: 0;
            color: #E6E6FA;
            border-collapse: collapse;
        }

        .song-table-wrap .table > thead > tr > th {
            background: linear-gradient(180deg, #4B0082 0%, #3a0065 100%);
            color: #E6E6FA;
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            text-align: center;
            padding: 14px 12px;
            border: none;
            border-bottom: 2px solid rgba(147, 112, 219, 0.35);
            vertical-align: middle;
        }

        .song-table-wrap .table > tbody > tr > td {
            background: transparent;
            color: #E6E6FA;
            font-size: 13.5px;
            padding: 13px 12px;
            border: none;
            border-bottom: 1px solid rgba(147, 112, 219, 0.07);
            vertical-align: middle;
            text-align: center;
        }

        /* alternating rows */
        .song-table-wrap .table > tbody > tr:nth-child(even) > td {
            background: rgba(147, 112, 219, 0.05);
        }

        /* hover */
        .song-table-wrap .table > tbody > tr:hover > td {
            background: rgba(147, 112, 219, 0.13);
            color: #fff;
        }

        .song-table-wrap .table > tbody > tr:last-child > td {
            border-bottom: none;
        }

        /* column alignment */
        .col-id { width: 90px; }
        .col-name { text-align: left !important; font-weight: 500; }
        .col-path { text-align: left !important; font-size: 12px !important; opacity: 0.8; }
        .col-plays { width: 130px; }
        .col-player { width: 300px; }

        /* ===== Plays counter ===== */
        .plays-pill {
            display: inline-block;
            min-width: 38px;
            padding: 3px 12px;
            border-radius: 12px;
            background: rgba(147, 112, 219, 0.18);
            color: #B79CE8;
            font-weight: 600;
            font-size: 12.5px;
        }

        /* ===== Audio player ===== */
        .song-table-wrap audio {
            width: 265px;
            height: 34px;
            vertical-align: middle;
            border-radius: 18px;
            outline: none;
            filter: none;
        }

        /* ===== Empty state ===== */
        .song-empty {
            background: rgba(30, 15, 50, 0.55);
            border: 1px solid rgba(147, 112, 219, 0.12);
            border-radius: 14px;
            padding: 55px 20px;
            text-align: center;
            color: rgba(230, 230, 250, 0.45);
            font-size: 15px;
        }

        /* ===== Footer ===== */
        .song-footer {
            text-align: center;
            padding: 26px 0 10px 0;
            color: rgba(230, 230, 250, 0.25);
            font-size: 12px;
        }

        @media screen and (max-width: 768px) {
            .song-page-title { font-size: 23px; }
            .song-table-wrap { overflow-x: auto; }
            .song-table-wrap .table > tbody > tr > td,
            .song-table-wrap .table > thead > tr > th { padding: 10px 8px; font-size: 12px; }
            .song-table-wrap audio { width: 190px; }
        }
    </style>

    <div class="song-page">
        <div class="song-container">
            <h1 class="song-page-title">Song Library</h1>
            <p class="song-page-subtitle">Browse and play your favorite tracks</p>

            <div class="song-table-wrap">
                <asp:GridView ID="gvSongs" runat="server"
                    AutoGenerateColumns="false"
                    CssClass="table"
                    GridLines="None"
                    BorderStyle="None"
                    EmptyDataText="No hay canciones disponibles."
                    OnRowDataBound="gvSongs_RowDataBound">
                    <Columns>
                        <asp:BoundField DataField="ID_SONG" HeaderText="ID Cancion" ReadOnly="True" ItemStyle-CssClass="col-id" HeaderStyle-CssClass="col-id" />
                        <asp:BoundField DataField="SONG_NAME" HeaderText="Nombre" ReadOnly="True" ItemStyle-CssClass="col-name" />
                        <asp:BoundField DataField="SONG_PATH" HeaderText="Ruta" ReadOnly="True" ItemStyle-CssClass="col-path" />
                        <asp:TemplateField HeaderText="Reproducciones" ItemStyle-CssClass="col-plays" HeaderStyle-CssClass="col-plays">
                            <ItemTemplate>
                                <span class="plays-pill"><%# Eval("PLAYS") %></span>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Reproducir" ItemStyle-CssClass="col-player" HeaderStyle-CssClass="col-player">
                            <ItemTemplate>
                                <audio controls preload="none">
                                    <source src="<%# Server.HtmlEncode(ResolveSongPath(Eval("SONG_PATH").ToString())) %>" type="audio/mpeg" />
                                    Your browser does not support the audio element.
                                </audio>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>

            <div class="song-footer">
                &copy; <%: DateTime.Now.Year %> PoliMusic - All rights reserved
            </div>
        </div>
    </div>
</asp:Content>
