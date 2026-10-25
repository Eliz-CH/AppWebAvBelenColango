<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/conexion.asp"-->
<%
'==============================================================================
' Archivo:      canciones.asp
' Descripcion:  Modulo de gestion de canciones para PoliMusic v1
'               - Autenticacion por sesion (Session("Username")/Session("Usuario"))
'               - Roles: normal (solo lectura + reproducir) / admin (CRUD completo)
'               - Lista de canciones con portada, reproductor de audio y acciones
' Proyecto:     PoliMusic v1 - ASP Clasico (VBScript)
'==============================================================================

'==============================================================================
' FUNCIONES AUXILIARES
'   Se definen aqui, al inicio del archivo, porque ASP Clasico evalua el
'   codigo en orden: una Function debe existir ANTES de su primer uso.
'==============================================================================

'--- IIf: emula el operador ternario que VBScript no tiene ---------------
Function IIf(condition, trueVal, falseVal)
    If condition Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function

'--- NormalizeAssetPath: convierte un path de la BD ("~/carpeta/archivo") ---
'    en una ruta web relativa valida ("carpeta/archivo"), ajustando el nombre
'    de la carpeta cuando el path guardado no coincide con la carpeta real.
Function NormalizeAssetPath(ByVal rawPath, ByVal realFolder)
    Dim p
    p = Trim(CStr(rawPath))
    p = Replace(p, "~/", "")
    p = Replace(p, "~\", "")
    p = Replace(p, "\", "/")
    If realFolder <> "" Then
        ' Reemplazar la carpeta de assets guardada en la BD por la carpeta real
        ' indicada, de modo que el path siempre resuelva en el disco.
        ' El orden importa: primero los nombres especificos, luego el generico.
        p = Replace(p, "albumImages/", realFolder & "/")
        p = Replace(p, "ArtistImages/", realFolder & "/")
        p = Replace(p, "userImages/", realFolder & "/")
        p = Replace(p, "playlistImages/", realFolder & "/")
        p = Replace(p, "songFiles/", realFolder & "/")
    End If
    NormalizeAssetPath = p
End Function

'--- Cache off ---------------------------------------------------------------
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CharSet = "utf-8"

'==============================================================================
' 1. AUTENTICACION Y ROLES
'==============================================================================

'--- Leer nombre de usuario desde sesion -------------------------------------
Dim strUsername
strUsername = ""
If Session("Username") <> "" And Not IsNull(Session("Username")) Then
    strUsername = CStr(Session("Username"))
ElseIf Session("Usuario") <> "" And Not IsNull(Session("Usuario")) Then
    strUsername = CStr(Session("Usuario"))
End If

'--- Leer rol desde sesion (acepta "Rol" o "UserRole") ----------------------
Dim strRol
strRol = ""
If Session("Rol") <> "" And Not IsNull(Session("Rol")) Then
    strRol = LCase(CStr(Session("Rol")))
ElseIf Session("UserRole") <> "" And Not IsNull(Session("UserRole")) Then
    strRol = LCase(CStr(Session("UserRole")))
End If

'--- Si no hay sesion activa, redirigir a login.asp --------------------------
If strUsername = "" Then
    Response.Redirect "login.asp"
    Response.End
End If

'--- Normalizar el rol a "normal" o "admin" ----------------------------------
'    La BD usa USER_TYPE 0=USER (normal), 1=ADMIN, pero login.asp pudo haber
'    guardado el texto del catalogo ("ADMIN"/"USER") o un numero.
Dim bIsAdmin
bIsAdmin = False
If strRol = "admin" Or strRol = "1" Or strRol = "administrator" Then
    bIsAdmin = True
    strRol = "admin"
Else
    strRol = "normal"
End If

'==============================================================================
' 2. PROCESAR ACCIONES CRUD (solo admin)
'==============================================================================

Dim strAction, strMessage, bError
strAction = Request.QueryString("action")
If strAction = "" Then strAction = Request.Form("action")
strMessage = ""
bError = False

If bIsAdmin Then
    '--- CREATE -----------------------------------------------------------------
    If strAction = "create" Then
        Dim newSongName, newSongPath, newPlays
        newSongName = Trim(Request.Form("song_name"))
        newSongPath = Trim(Request.Form("song_path"))
        newPlays    = Trim(Request.Form("plays"))
        If newSongName = "" Or newSongPath = "" Then
            strMessage = "El nombre y la ruta de la cancion son obligatorios."
            bError = True
        Else
            If newPlays = "" Or Not IsNumeric(newPlays) Then newPlays = 0
            Dim cmdInsert
            Set cmdInsert = Server.CreateObject("ADODB.Command")
            cmdInsert.ActiveConnection = conn
            cmdInsert.CommandType = 1 ' adCmdText
            cmdInsert.CommandText = "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, ?)"
            cmdInsert.Parameters.Append cmdInsert.CreateParameter("@name", 200, 1, 50, Left(newSongName, 50))
            cmdInsert.Parameters.Append cmdInsert.CreateParameter("@path", 200, 1, 255, Left(newSongPath, 255))
            cmdInsert.Parameters.Append cmdInsert.CreateParameter("@plays", 3, 1, , CLng(newPlays))
            On Error Resume Next
            cmdInsert.Execute
            If Err.Number <> 0 Then
                strMessage = "Error al crear la cancion: " & Err.Description
                bError = True
            Else
                strMessage = "Cancion creada correctamente."
            End If
            On Error GoTo 0
            Set cmdInsert = Nothing
        End If

    '--- UPDATE -----------------------------------------------------------------
    ElseIf strAction = "update" Then
        Dim updId, updName, updPath, updPlays
        updId   = Trim(Request.Form("id_song"))
        updName = Trim(Request.Form("song_name"))
        updPath = Trim(Request.Form("song_path"))
        updPlays= Trim(Request.Form("plays"))
        If updId = "" Or Not IsNumeric(updId) Or updName = "" Or updPath = "" Then
            strMessage = "Datos insuficientes para actualizar."
            bError = True
        Else
            If updPlays = "" Or Not IsNumeric(updPlays) Then updPlays = 0
            Dim cmdUpdate
            Set cmdUpdate = Server.CreateObject("ADODB.Command")
            cmdUpdate.ActiveConnection = conn
            cmdUpdate.CommandType = 1
            cmdUpdate.CommandText = "UPDATE TBL_SONG SET SONG_NAME=?, SONG_PATH=?, PLAYS=? WHERE ID_SONG=?"
            cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("@name", 200, 1, 50, Left(updName, 50))
            cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("@path", 200, 1, 255, Left(updPath, 255))
            cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("@plays", 3, 1, , CLng(updPlays))
            cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("@id", 3, 1, , CLng(updId))
            On Error Resume Next
            cmdUpdate.Execute
            If Err.Number <> 0 Then
                strMessage = "Error al actualizar: " & Err.Description
                bError = True
            Else
                strMessage = "Cancion actualizada correctamente."
            End If
            On Error GoTo 0
            Set cmdUpdate = Nothing
        End If

    '--- DELETE -----------------------------------------------------------------
    '     TBL_SONG es referenciada por 3 tablas hijas:
    '       TBL_TRACKLIST, TBL_PLAYER, TBL_PLAYLIST_SONG
    '     Se eliminan primero las filas hijas de forma explicita y luego la
    '     cancion. Esto funciona aunque las FKs no tengan ON DELETE CASCADE
    '     (la BD ya las tiene en CASCADE, pero el borrado explicito es la
    '     garantia extra para que la operacion nunca falle).
    ElseIf strAction = "delete" Then
        Dim delId
        delId = Trim(Request.Form("id_song"))
        If delId = "" Or Not IsNumeric(delId) Then
            strMessage = "ID de cancion no valido."
            bError = True
        Else
            Dim cmdChild, cmdDel, bDeleteOk
            bDeleteOk = True
            On Error Resume Next

            ' 1) Borrar las referencias hijas (tracklist / player / playlist_song)
            Set cmdChild = Server.CreateObject("ADODB.Command")
            cmdChild.ActiveConnection = conn
            cmdChild.CommandType = 1
            cmdChild.CommandText = "DELETE FROM TBL_TRACKLIST WHERE ID_SONG=?"
            cmdChild.Parameters.Append cmdChild.CreateParameter("@id", 3, 1, , CLng(delId))
            cmdChild.Execute
            If Err.Number <> 0 Then Err.Clear

            cmdChild.CommandText = "DELETE FROM TBL_PLAYER WHERE ID_SONG=?"
            cmdChild.Parameters.Delete 0
            cmdChild.Parameters.Append cmdChild.CreateParameter("@id", 3, 1, , CLng(delId))
            cmdChild.Execute
            If Err.Number <> 0 Then Err.Clear

            cmdChild.CommandText = "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG=?"
            cmdChild.Parameters.Delete 0
            cmdChild.Parameters.Append cmdChild.CreateParameter("@id", 3, 1, , CLng(delId))
            cmdChild.Execute
            If Err.Number <> 0 Then Err.Clear
            Set cmdChild = Nothing

            ' 2) Borrar la cancion
            Set cmdDel = Server.CreateObject("ADODB.Command")
            cmdDel.ActiveConnection = conn
            cmdDel.CommandType = 1
            cmdDel.CommandText = "DELETE FROM TBL_SONG WHERE ID_SONG=?"
            cmdDel.Parameters.Append cmdDel.CreateParameter("@id", 3, 1, , CLng(delId))
            cmdDel.Execute

            If Err.Number <> 0 Then
                strMessage = "Error al eliminar la cancion: " & Err.Description
                bError = True
                Err.Clear
            Else
                strMessage = "Cancion eliminada correctamente."
            End If
            On Error GoTo 0
            Set cmdDel = Nothing
        End If
    End If
End If

'==============================================================================
' 3. CONSULTAR CANCIONES (subqueries TOP 1: evita productos cartesianos)
'==============================================================================

Dim rs, sql
'--- NOTA SOBRE DUPLICADOS ---------------------------------------------------
' Un JOIN directo con TBL_PLAYER (artistas) y TBL_TRACKLIST (albumes) genera
' un producto cartesiano: "Adventure" tiene 2 artistas x 2 albumes = 4 filas.
' Para evitarlo se usan SUBQUERIES con TOP 1: cada cancion aparece UNA sola
' vez, tomando el primer artista y el primer album asociados.
' Ademas se agrega la imagen del ARTISTA (no solo la del album) para el avatar.
sql = "SELECT S.ID_SONG, S.SONG_NAME, S.SONG_PATH, S.PLAYS, " & _
      "       (SELECT TOP 1 A.ARTIST_NAME + ' ' + ISNULL(A.ARTIST_LASTNAME,'') " & _
      "          FROM TBL_PLAYER P INNER JOIN TBL_ARTIST A ON A.ID_ARTIST = P.ID_ARTIST " & _
      "         WHERE P.ID_SONG = S.ID_SONG ORDER BY A.ARTIST_NAME) AS ARTIST_FULL, " & _
      "       (SELECT TOP 1 ISNULL(A.ARTIST_IMAGE,'') " & _
      "          FROM TBL_PLAYER P INNER JOIN TBL_ARTIST A ON A.ID_ARTIST = P.ID_ARTIST " & _
      "         WHERE P.ID_SONG = S.ID_SONG ORDER BY A.ARTIST_NAME) AS ARTIST_IMAGE, " & _
      "       (SELECT TOP 1 AL.TITLE_ALBUM " & _
      "          FROM TBL_TRACKLIST T INNER JOIN TBL_ALBUM AL ON AL.ID_ALBUM = T.ID_ALBUM " & _
      "         WHERE T.ID_SONG = S.ID_SONG ORDER BY AL.ID_ALBUM) AS TITLE_ALBUM, " & _
      "       (SELECT TOP 1 AL.ALBUM_IMAGE_PATH " & _
      "          FROM TBL_TRACKLIST T INNER JOIN TBL_ALBUM AL ON AL.ID_ALBUM = T.ID_ALBUM " & _
      "         WHERE T.ID_SONG = S.ID_SONG ORDER BY AL.ID_ALBUM) AS ALBUM_IMAGE_PATH, " & _
      "       (SELECT COUNT(DISTINCT P2.ID_ARTIST) " & _
      "          FROM TBL_PLAYER P2 WHERE P2.ID_SONG = S.ID_SONG) AS ARTIST_COUNT " & _
      "FROM TBL_SONG S " & _
      "ORDER BY S.ID_SONG"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorType = 1  ' adOpenKeyset
rs.LockType   = 1  ' adLockReadOnly
On Error Resume Next
rs.Open sql, conn
If Err.Number <> 0 Then
    strMessage = "Error al consultar canciones: " & Err.Description
    bError = True
End If
On Error GoTo 0

'==============================================================================
' 4. RENDER HTML
'==============================================================================
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PoliMusic &middot; Listado de Canciones</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        /* ===================================================================
           PALETA DE COLORES - PoliMusic (Purpura/Indigo oscuro + azul vibrante)
           =================================================================== */
        :root {
            --bg-base:       #1E1B2E;
            --bg-surface:    #28253E;
            --bg-elevated:   #332F4F;
            --bg-hover:      #3D3859;
            --accent:        #6C5CE7;
            --accent-light:  #8B7FF0;
            --accent-glow:   rgba(108,92,231,0.35);
            --blue-vibrant:  #00B4D8;
            --blue-glow:     rgba(0,180,216,0.25);
            --text-primary:  #F5F3FA;
            --text-secondary:#B4B0CA;
            --text-muted:    #7C7894;
            --border:        rgba(255,255,255,0.08);
            --danger:        #FF5C7C;
            --danger-glow:   rgba(255,92,124,0.25);
            --success:       #4ADE80;
            --warning:       #FFB547;
            --radius:        14px;
            --radius-sm:     8px;
            --shadow:        0 8px 32px rgba(0,0,0,0.35);
            --transition:    all 0.22s ease;
        }

        * { margin:0; padding:0; box-sizing:border-box; }

        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: var(--bg-base);
            color: var(--text-primary);
            min-height: 100vh;
            line-height: 1.6;
        }

        /* ---- Background decorativo ---- */
        body::before {
            content: '';
            position: fixed; top: -120px; right: -120px;
            width: 420px; height: 420px;
            background: radial-gradient(circle, var(--accent-glow), transparent 70%);
            border-radius: 50%;
            z-index: 0; pointer-events: none;
        }
        body::after {
            content: '';
            position: fixed; bottom: -100px; left: -100px;
            width: 360px; height: 360px;
            background: radial-gradient(circle, var(--blue-glow), transparent 70%);
            border-radius: 50%;
            z-index: 0; pointer-events: none;
        }

        /* ===================================================================
           HEADER
           =================================================================== */
        .header {
            position: relative; z-index: 1;
            background: var(--bg-surface);
            border-bottom: 1px solid var(--border);
            box-shadow: var(--shadow);
        }
        .header-inner {
            max-width: 1280px; margin: 0 auto;
            display: flex; align-items: center; justify-content: space-between;
            padding: 18px 32px;
        }
        .logo {
            display: flex; align-items: center; gap: 12px;
            font-size: 1.4rem; font-weight: 800;
            letter-spacing: -0.5px;
        }
        .logo-icon {
            width: 42px; height: 42px;
            background: linear-gradient(135deg, var(--accent), var(--blue-vibrant));
            border-radius: 12px;
            display: flex; align-items: center; justify-content: center;
            font-size: 1.3rem; flex-shrink: 0;
        }
        .logo-text { color: var(--text-primary); }
        .logo-text span { color: var(--blue-vibrant); }

        .header-user {
            display: flex; align-items: center; gap: 16px;
        }
        .welcome-badge {
            display: flex; align-items: center; gap: 10px;
            padding: 8px 16px;
            background: var(--bg-elevated);
            border: 1px solid var(--border);
            border-radius: 50px;
            font-size: 0.88rem; font-weight: 500;
        }
        .welcome-badge .avatar {
            width: 28px; height: 28px;
            background: linear-gradient(135deg, var(--accent), var(--blue-vibrant));
            border-radius: 50%;
            display: flex; align-items: center; justify-content: center;
            font-size: 0.8rem; font-weight: 700;
            text-transform: uppercase;
            overflow: hidden;
            flex-shrink: 0;
        }
        .welcome-badge .avatar img {
            width: 100%; height: 100%;
            object-fit: cover;
            border-radius: 50%;
            display: block;
        }
        .role-badge {
            font-size: 0.7rem; font-weight: 700;
            padding: 3px 10px; border-radius: 50px;
            text-transform: uppercase; letter-spacing: 0.5px;
        }
        .role-badge.admin {
            background: var(--accent);
            color: #fff;
        }
        .role-badge.normal {
            background: var(--bg-hover);
            color: var(--text-secondary);
        }

        .btn-logout {
            display: inline-flex; align-items: center; gap: 8px;
            padding: 8px 18px;
            background: var(--danger);
            color: #fff;
            border: none; border-radius: 50px;
            font-size: 0.85rem; font-weight: 600;
            cursor: pointer; text-decoration: none;
            transition: var(--transition);
        }
        .btn-logout:hover {
            background: #ff3355;
            transform: translateY(-1px);
            box-shadow: 0 4px 16px var(--danger-glow);
        }

        /* ===================================================================
           MAIN CONTAINER
           =================================================================== */
        .container {
            position: relative; z-index: 1;
            max-width: 1280px; margin: 0 auto;
            padding: 32px;
        }

        .page-title {
            font-size: 1.8rem; font-weight: 700;
            margin-bottom: 4px;
        }
        .page-title span { color: var(--accent-light); }
        .page-subtitle {
            color: var(--text-secondary);
            font-size: 0.95rem;
            margin-bottom: 24px;
        }

        /* ---- Alert / Message ---- */
        .alert {
            padding: 14px 20px;
            border-radius: var(--radius-sm);
            margin-bottom: 20px;
            font-size: 0.9rem; font-weight: 500;
            display: flex; align-items: center; gap: 10px;
        }
        .alert.success {
            background: rgba(74,222,128,0.12);
            border: 1px solid rgba(74,222,128,0.3);
            color: var(--success);
        }
        .alert.error {
            background: rgba(255,92,124,0.12);
            border: 1px solid rgba(255,92,124,0.3);
            color: var(--danger);
        }

        /* ---- Admin toolbar ---- */
        .admin-toolbar {
            display: flex; align-items: center; justify-content: space-between;
            margin-bottom: 20px;
            gap: 16px;
        }
        .admin-toolbar .info {
            color: var(--text-muted);
            font-size: 0.85rem;
        }
        .admin-toolbar .info b { color: var(--accent-light); }

        .btn {
            display: inline-flex; align-items: center; gap: 8px;
            padding: 10px 22px;
            border: none; border-radius: 50px;
            font-size: 0.88rem; font-weight: 600;
            cursor: pointer; text-decoration: none;
            transition: var(--transition);
            font-family: inherit;
        }
        .btn-primary {
            background: linear-gradient(135deg, var(--accent), var(--accent-light));
            color: #fff;
        }
        .btn-primary:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 20px var(--accent-glow);
        }
        .btn-edit {
            background: rgba(0,180,216,0.15);
            color: var(--blue-vibrant);
            border: 1px solid rgba(0,180,216,0.3);
        }
        .btn-edit:hover { background: rgba(0,180,216,0.25); }
        .btn-delete {
            background: rgba(255,92,124,0.12);
            color: var(--danger);
            border: 1px solid rgba(255,92,124,0.25);
        }
        .btn-delete:hover { background: rgba(255,92,124,0.22); }
        .btn-sm { padding: 7px 14px; font-size: 0.8rem; }

        /* ===================================================================
           TABLA DE CANCIONES
           =================================================================== */
        .songs-wrapper {
            background: var(--bg-surface);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            overflow: hidden;
            box-shadow: var(--shadow);
        }

        .songs-table {
            width: 100%;
            border-collapse: collapse;
        }
        .songs-table thead th {
            background: var(--bg-elevated);
            color: var(--text-secondary);
            font-size: 0.78rem; font-weight: 600;
            text-transform: uppercase; letter-spacing: 0.8px;
            padding: 14px 18px;
            text-align: left;
            border-bottom: 1px solid var(--border);
        }
        .songs-table tbody tr {
            border-bottom: 1px solid var(--border);
            transition: var(--transition);
        }
        .songs-table tbody tr:hover {
            background: var(--bg-elevated);
        }
        .songs-table tbody tr:last-child { border-bottom: none; }
        .songs-table td {
            padding: 14px 18px;
            vertical-align: middle;
        }

        /* ---- Cover image cell ---- */
        .cover-cell {
            width: 64px;
        }
        .cover-cell img {
            width: 48px; height: 48px;
            border-radius: 8px;
            object-fit: cover;
            border: 1px solid var(--border);
            display: block;
        }

        /* ---- Song info ---- */
        .song-info .song-title {
            font-weight: 600; font-size: 0.95rem;
            color: var(--text-primary);
        }
        .song-info .song-meta {
            font-size: 0.8rem; color: var(--text-muted);
        }

        /* ---- Artist cell with avatar ---- */
        .artist-cell {
            display: flex; align-items: center; gap: 10px;
        }
        .artist-avatar {
            width: 32px; height: 32px;
            border-radius: 50%;
            object-fit: cover;
            border: 1px solid var(--border);
            flex-shrink: 0;
            display: block;
        }
        .artist-avatar.placeholder {
            display: flex; align-items: center; justify-content: center;
            background: linear-gradient(135deg, var(--accent), var(--blue-vibrant));
            color: #fff;
            font-size: 0.8rem; font-weight: 700;
            text-transform: uppercase;
        }
        .artist-name {
            font-size: 0.88rem;
            white-space: nowrap;
        }

        /* ---- Audio player ---- */
        .audio-cell audio {
            width: 240px; height: 36px;
            border-radius: 50px;
        }
        .audio-cell audio::-webkit-media-controls-panel {
            background: var(--bg-elevated);
        }

        /* ---- Plays badge ---- */
        .plays-badge {
            display: inline-flex; align-items: center; gap: 5px;
            padding: 4px 12px;
            background: var(--bg-elevated);
            border-radius: 50px;
            font-size: 0.78rem; font-weight: 600;
            color: var(--text-secondary);
        }

        /* ---- Actions cell ---- */
        .actions-cell {
            display: flex; gap: 8px;
        }

        .no-songs {
            padding: 48px; text-align: center;
            color: var(--text-muted);
        }
        .no-songs .icon { font-size: 3rem; margin-bottom: 12px; }

        /* ===================================================================
           MODAL
           =================================================================== */
        .modal-overlay {
            display: none;
            position: fixed; inset: 0;
            background: rgba(0,0,0,0.65);
            backdrop-filter: blur(6px);
            z-index: 1000;
            align-items: center; justify-content: center;
            padding: 20px;
        }
        .modal-overlay.active {
            display: flex;
        }
        .modal {
            background: var(--bg-surface);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            width: 100%; max-width: 480px;
            box-shadow: 0 24px 64px rgba(0,0,0,0.5);
            overflow: hidden;
            animation: modalIn 0.25s ease;
        }
        @keyframes modalIn {
            from { transform: scale(0.95) translateY(10px); opacity: 0; }
            to   { transform: scale(1) translateY(0); opacity: 1; }
        }
        .modal-header {
            padding: 20px 24px;
            border-bottom: 1px solid var(--border);
            display: flex; align-items: center; justify-content: space-between;
        }
        .modal-header h3 {
            font-size: 1.15rem; font-weight: 700;
        }
        .modal-close {
            background: none; border: none;
            color: var(--text-muted);
            font-size: 1.4rem; cursor: pointer;
            line-height: 1; padding: 0;
            transition: var(--transition);
        }
        .modal-close:hover { color: var(--text-primary); }
        .modal-body {
            padding: 24px;
        }
        .form-group {
            margin-bottom: 18px;
        }
        .form-group label {
            display: block;
            font-size: 0.82rem; font-weight: 600;
            color: var(--text-secondary);
            margin-bottom: 6px;
        }
        .form-group input, .form-group select {
            width: 100%;
            padding: 10px 14px;
            background: var(--bg-elevated);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            color: var(--text-primary);
            font-size: 0.9rem;
            font-family: inherit;
            transition: var(--transition);
        }
        .form-group input:focus, .form-group select:focus {
            outline: none;
            border-color: var(--accent);
            box-shadow: 0 0 0 3px var(--accent-glow);
        }
        .form-group .hint {
            font-size: 0.75rem; color: var(--text-muted);
            margin-top: 4px;
        }
        .modal-footer {
            padding: 18px 24px;
            border-top: 1px solid var(--border);
            display: flex; gap: 10px; justify-content: flex-end;
        }

        /* ===================================================================
           RESPONSIVE
           =================================================================== */
        @media (max-width: 900px) {
            .header-inner { flex-direction: column; gap: 14px; padding: 16px; }
            .container { padding: 20px; }
            .songs-table { font-size: 0.85rem; }
            .songs-table th, .songs-table td { padding: 10px 12px; }
            .audio-cell audio { width: 160px; }
            .col-album, .col-plays { display: none; }
        }
        @media (max-width: 600px) {
            .audio-cell audio { width: 120px; height: 32px; }
            .col-artist { display: none; }
            .actions-cell { flex-direction: column; }
        }
    </style>
</head>
<body>

<!-- ========================== HEADER ========================== -->
<header class="header">
    <div class="header-inner">
        <div class="logo">
            <div class="logo-icon">&#9835;</div>
            <div class="logo-text">Poli<span>Music</span></div>
        </div>
        <div class="header-user">
            <div class="welcome-badge">
                <div class="avatar">
                    <%
                    ' --- Foto de perfil del usuario (si existe en UserImages) ---
                    Dim strUserPhoto, strUserPhotoWeb
                    strUserPhoto = ""
                    If Not IsNull(Session("UserPhoto")) Then strUserPhoto = Trim(CStr(Session("UserPhoto")))
                    strUserPhotoWeb = ""
                    If strUserPhoto <> "" Then
                        strUserPhotoWeb = NormalizeAssetPath(strUserPhoto, "UserImages")
                        ' Verificar que el archivo exista realmente
                        Dim fsHdr
                        On Error Resume Next
                        Set fsHdr = Server.CreateObject("Scripting.FileSystemObject")
                        If Not (fsHdr Is Nothing) Then
                            If InStr(strUserPhotoWeb, ":") = 0 Then
                                If Not fsHdr.FileExists(Server.MapPath(strUserPhotoWeb)) Then
                                    strUserPhotoWeb = ""
                                End If
                            End If
                            Set fsHdr = Nothing
                        End If
                        On Error GoTo 0
                    End If
                    %>
                    <% If strUserPhotoWeb <> "" Then %>
                        <img src="<%= strUserPhotoWeb %>" alt="<%= Server.HTMLEncode(strUsername) %>">
                    <% Else %>
                        <%= Left(strUsername, 1) %>
                    <% End If %>
                </div>
                <span>&iexcl;Bienvenido/a, <%= Server.HTMLEncode(strUsername) %>!</span>
                <span class="role-badge <%= strRol %>"><%= strRol %></span>
            </div>
            <a href="logout.asp" class="btn-logout" title="Cerrar sesion">
                &#8634; Cerrar Sesi&oacute;n
            </a>
        </div>
    </div>
</header>

<!-- ========================== MAIN ========================== -->
<div class="container">

    <h1 class="page-title">Listado de <span>Canciones</span></h1>
    <p class="page-subtitle">Explora y reproduce la biblioteca musical de PoliMusic.</p>

    <% If strMessage <> "" Then %>
        <div class="alert <%= IIf(bError, "error", "success") %>">
            <span><%= IIf(bError, "&#9888;", "&#10003;") %></span>
            <%= Server.HTMLEncode(strMessage) %>
        </div>
    <% End If %>

    <% If bIsAdmin Then %>
        <div class="admin-toolbar">
            <div class="info">
                Modo <b>Administrador</b> &mdash; Puede crear, editar y eliminar canciones.
            </div>
            <button class="btn btn-primary" onclick="openCreateModal()">
                &#43; Nueva Canci&oacute;n
            </button>
        </div>
    <% End If %>

    <!-- ===================== TABLA DE CANCIONES ===================== -->
    <div class="songs-wrapper">
        <table class="songs-table">
            <thead>
                <tr>
                    <th>#</th>
                    <th>Portada</th>
                    <th>T&iacute;tulo</th>
                    <th class="col-artist">Artista</th>
                    <th class="col-album">&Aacute;lbum</th>
                    <th>Reproductor</th>
                    <th class="col-plays">Reproducciones</th>
                    <% If bIsAdmin Then %><th>Acciones</th><% End If %>
                </tr>
            </thead>
            <tbody>
                <%
                Dim songIndex, displayArtist, displayAlbum, songPathClean, coverImg
                Dim artistImg, fsObj
                songIndex = 0

                ' Objeto de sistema de archivos para verificar que el asset exista
                On Error Resume Next
                Set fsObj = Server.CreateObject("Scripting.FileSystemObject")
                On Error GoTo 0

                Do While Not rs.EOF
                    songIndex = songIndex + 1

                    ' --- Artista (viene del subquery TOP 1, ya concatenado) ---
                    If IsNull(rs("ARTIST_FULL")) Or Trim(CStr(rs("ARTIST_FULL"))) = "" Then
                        displayArtist = "Artista desconocido"
                    Else
                        displayArtist = Trim(CStr(rs("ARTIST_FULL")))
                    End If

                    ' --- Indicador de colaboradores adicionales (+N) ---
                    Dim collabCount
                    collabCount = 0
                    If Not IsNull(rs("ARTIST_COUNT")) Then
                        If IsNumeric(rs("ARTIST_COUNT")) Then
                            collabCount = CLng(rs("ARTIST_COUNT"))
                        End If
                    End If
                    If collabCount > 1 Then
                        displayArtist = displayArtist & " +" & (collabCount - 1)
                    End If

                    ' --- Album ---
                    If IsNull(rs("TITLE_ALBUM")) Or Trim(CStr(rs("TITLE_ALBUM"))) = "" Then
                        displayAlbum = "Sin album"
                    Else
                        displayAlbum = Trim(CStr(rs("TITLE_ALBUM")))
                    End If

                    ' --- Ruta del audio: ~/songFiles/xxx.mp3 -> SongFiles/xxx.mp3 ---
                    songPathClean = NormalizeAssetPath(rs("SONG_PATH"), "SongFiles")

                    ' --- Portada: imagen del ALBUM, con fallback a default.jpg ---
                    If IsNull(rs("ALBUM_IMAGE_PATH")) Or Trim(CStr(rs("ALBUM_IMAGE_PATH"))) = "" Then
                        coverImg = "PlaylistImages/default.jpg"
                    Else
                        coverImg = NormalizeAssetPath(rs("ALBUM_IMAGE_PATH"), "albumImages")
                    End If

                    ' --- Avatar del ARTISTA (se muestra junto al nombre) ---
                    If IsNull(rs("ARTIST_IMAGE")) Or Trim(CStr(rs("ARTIST_IMAGE"))) = "" Then
                        artistImg = ""
                    Else
                        artistImg = NormalizeAssetPath(rs("ARTIST_IMAGE"), "ArtistImages")
                    End If

                    ' --- Verificar que el archivo de portada exista en disco ---
                    '     (si no existe, usar la imagen por defecto)
                    If Not (fsObj Is Nothing) Then
                        If InStr(coverImg, ":") = 0 Then
                            On Error Resume Next
                            If Not fsObj.FileExists(Server.MapPath(coverImg)) Then
                                coverImg = "PlaylistImages/default.jpg"
                            End If
                            On Error GoTo 0
                        End If
                        ' Verificar tambien la imagen del artista
                        If artistImg <> "" And InStr(artistImg, ":") = 0 Then
                            On Error Resume Next
                            If Not fsObj.FileExists(Server.MapPath(artistImg)) Then
                                artistImg = ""
                            End If
                            On Error GoTo 0
                        End If
                    End If
                %>
                <tr>
                    <td><%= songIndex %></td>
                    <td class="cover-cell">
                        <img src="<%= coverImg %>" alt="Portada" loading="lazy">
                    </td>
                    <td>
                        <div class="song-info">
                            <div class="song-title"><%= Server.HTMLEncode(rs("SONG_NAME")) %></div>
                            <div class="song-meta">ID #<%= rs("ID_SONG") %></div>
                        </div>
                    </td>
                    <td class="col-artist">
                        <div class="artist-cell">
                            <% If artistImg <> "" Then %>
                                <img class="artist-avatar" src="<%= artistImg %>" alt="<%= Server.HTMLEncode(displayArtist) %>" loading="lazy">
                            <% Else %>
                                <span class="artist-avatar placeholder"><%= Left(displayArtist, 1) %></span>
                            <% End If %>
                            <span class="artist-name"><%= Server.HTMLEncode(displayArtist) %></span>
                        </div>
                    </td>
                    <td class="col-album"><%= Server.HTMLEncode(displayAlbum) %></td>
                    <td class="audio-cell">
                        <audio controls preload="none">
                            <source src="<%= songPathClean %>" type="audio/mpeg">
                            Tu navegador no soporta audio HTML5.
                        </audio>
                    </td>
                    <td class="col-plays">
                        <span class="plays-badge">&#9658; <%= rs("PLAYS") %></span>
                    </td>
                    <% If bIsAdmin Then %>
                    <td>
                        <div class="actions-cell">
                            <button class="btn btn-edit btn-sm"
                                    onclick="openEditModal(
                                        <%= rs("ID_SONG") %>,
                                        '<%= Replace(Server.HTMLEncode(rs("SONG_NAME")), "'", "\'") %>',
                                        '<%= Replace(rs("SONG_PATH"), "'", "\'") %>',
                                        <%= rs("PLAYS") %>
                                    )">
                                &#9998; Editar
                            </button>
                            <button class="btn btn-delete btn-sm"
                                    onclick="confirmDelete(
                                        <%= rs("ID_SONG") %>,
                                        '<%= Replace(Server.HTMLEncode(rs("SONG_NAME")), "'", "\'") %>'
                                    )">
                                &#128465; Eliminar
                            </button>
                        </div>
                    </td>
                    <% End If %>
                </tr>
                <%
                    rs.MoveNext
                Loop

                If songIndex = 0 Then
                %>
                <tr>
                    <td colspan="<%= IIf(bIsAdmin, 8, 7) %>" class="no-songs">
                        <div class="icon">&#9835;</div>
                        <div>No hay canciones registradas en la base de datos.</div>
                    </td>
                </tr>
                <%
                End If

                rs.Close
                Set rs = Nothing
                conn.Close
                Set conn = Nothing
                %>
            </tbody>
        </table>
    </div>

</div>

<!-- ===================== MODAL: CREAR / EDITAR ===================== -->
<% If bIsAdmin Then %>
<div class="modal-overlay" id="songModal" onclick="if(event.target===this) closeModal()">
    <div class="modal">
        <div class="modal-header">
            <h3 id="modalTitle">Nueva Canci&oacute;n</h3>
            <button class="modal-close" onclick="closeModal()">&times;</button>
        </div>
        <form method="POST" action="canciones.asp" id="songForm" onsubmit="return validateForm()">
            <div class="modal-body">
                <input type="hidden" name="action" id="formAction" value="create">
                <input type="hidden" name="id_song" id="fldId" value="">

                <div class="form-group">
                    <label for="fldName">Nombre de la canci&oacute;n *</label>
                    <input type="text" id="fldName" name="song_name" maxlength="50"
                           placeholder="Ej: Adventure" required>
                </div>
                <div class="form-group">
                    <label for="fldPath">Ruta del archivo (.mp3) *</label>
                    <input type="text" id="fldPath" name="song_path" maxlength="255"
                           placeholder="~/songFiles/bensound-adventure.mp3" required>
                    <div class="hint">Use ~/songFiles/ como raiz relativa. El archivo debe existir en la carpeta <code>SongFiles/</code>.</div>
                </div>
                <div class="form-group">
                    <label for="fldPlays">Reproducciones iniciales</label>
                    <input type="number" id="fldPlays" name="plays" min="0" value="0">
                </div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-sm" style="background:var(--bg-elevated);color:var(--text-secondary)" onclick="closeModal()">Cancelar</button>
                <button type="submit" class="btn btn-primary btn-sm" id="btnSave">Guardar</button>
            </div>
        </form>
    </div>
</div>

<!-- ===================== MODAL: CONFIRMAR ELIMINACION ===================== -->
<div class="modal-overlay" id="deleteModal" onclick="if(event.target===this) closeDeleteModal()">
    <div class="modal" style="max-width:420px">
        <div class="modal-header">
            <h3>Confirmar eliminaci&oacute;n</h3>
            <button class="modal-close" onclick="closeDeleteModal()">&times;</button>
        </div>
        <div class="modal-body">
            <p>&iquest;Seguro que deseas eliminar la canci&oacute;n?</p>
            <p style="margin-top:8px;font-weight:600;color:var(--blue-vibrant);font-size:1.05rem;" id="deleteSongName"></p>
            <p style="margin-top:10px;font-size:0.8rem;color:var(--danger);">
                &#9888; Esta acci&oacute;n no se puede deshacer.
            </p>
        </div>
        <div class="modal-footer">
            <button type="button" class="btn btn-sm" style="background:var(--bg-elevated);color:var(--text-secondary)" onclick="closeDeleteModal()">Cancelar</button>
            <form method="POST" action="canciones.asp" id="deleteForm" style="margin:0">
                <input type="hidden" name="action" value="delete">
                <input type="hidden" name="id_song" id="deleteId" value="">
                <button type="submit" class="btn btn-delete btn-sm">S&iacute;, eliminar</button>
            </form>
        </div>
    </div>
</div>

<script>
    // ===== Funciones del modal Crear/Editar =====
    function openCreateModal() {
        document.getElementById('modalTitle').innerHTML = 'Nueva Canci&oacute;n';
        document.getElementById('formAction').value = 'create';
        document.getElementById('fldId').value = '';
        document.getElementById('fldName').value = '';
        document.getElementById('fldPath').value = '~/songFiles/';
        document.getElementById('fldPlays').value = '0';
        document.getElementById('btnSave').textContent = 'Guardar';
        document.getElementById('songModal').classList.add('active');
        document.getElementById('fldName').focus();
    }

    function openEditModal(id, name, path, plays) {
        document.getElementById('modalTitle').innerHTML = 'Editar Canci&oacute;n';
        document.getElementById('formAction').value = 'update';
        document.getElementById('fldId').value = id;
        document.getElementById('fldName').value = name;
        document.getElementById('fldPath').value = path;
        document.getElementById('fldPlays').value = plays;
        document.getElementById('btnSave').textContent = 'Actualizar';
        document.getElementById('songModal').classList.add('active');
    }

    function closeModal() {
        document.getElementById('songModal').classList.remove('active');
    }

    function validateForm() {
        var name = document.getElementById('fldName').value.trim();
        var path = document.getElementById('fldPath').value.trim();
        if (name === '' || path === '') {
            alert('El nombre y la ruta son obligatorios.');
            return false;
        }
        return true;
    }

    // ===== Funciones del modal Eliminar =====
    function confirmDelete(id, name) {
        document.getElementById('deleteId').value = id;
        document.getElementById('deleteSongName').textContent = name;
        document.getElementById('deleteModal').classList.add('active');
    }

    function closeDeleteModal() {
        document.getElementById('deleteModal').classList.remove('active');
    }

    // ===== Cerrar modal con tecla ESC =====
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            closeModal();
            closeDeleteModal();
        }
    });
</script>
<% End If %>

</body>
</html>
<%
' Las funciones auxiliares (IIf y NormalizeAssetPath) se definieron al inicio
' del archivo, antes de su primer uso.
%>