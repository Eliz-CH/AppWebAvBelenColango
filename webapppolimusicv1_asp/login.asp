<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="includes/conexion.asp"-->
<%
'==============================================================================
' Archivo:      login.asp
' Descripcion:  Inicio de sesion para PoliMusic v1
'               - Formulario de login (Usuario + Password)
'               - Validacion contra TBL_USER en SQL Server (BDD_PoliMusic)
'               - Guarda Session("Username"), Session("Rol") y redirige a canciones.asp
'               - Maneja roles normal (USER_TYPE=0) y admin (USER_TYPE=1)
' Proyecto:     PoliMusic v1 - ASP Clasico (VBScript)
'==============================================================================

'--- Cache off ---------------------------------------------------------------
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CharSet = "utf-8"

'--- Si ya hay sesion activa, ir directo a canciones.asp ---------------------
If Session("Username") <> "" And Not IsNull(Session("Username")) Then
    If Trim(Session("Username")) <> "" Then
        Response.Redirect "canciones.asp"
        Response.End
    End If
End If

'==============================================================================
' PROCESAR LOGIN (POST)
'==============================================================================
Dim strErrorMessage, bShowError
strErrorMessage = ""
bShowError = False

If UCase(Request.ServerVariables("HTTP_METHOD")) = "POST" Then

    Dim strInputUser, strInputPass
    strInputUser = Trim(Request.Form("username"))
    strInputPass = Trim(Request.Form("password"))

    '--- Validar que no vengan vacios ---------------------------------------
    If strInputUser = "" Or strInputPass = "" Then
        strErrorMessage = "Por favor, complete todos los campos."
        bShowError = True
    Else
        '--- Consultar el usuario en TBL_USER mediante parametros ------------
        '    La tabla TBL_USER tiene:
        '      USERNAME    varchar(50)
        '      PASSWORD    varchar(50) -> puede ser texto plano o SHA1 hash
        '      USER_TYPE   int (0=USER/normal, 1=ADMIN)
        '      EMAIL       varchar(50)
        '      USER_PHOTO  varchar(255)

        Dim cmd, rsUser
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.CommandType = 1 ' adCmdText

        ' --- Consulta parametrizada para evitar SQL Injection ---
        '     Busca por USERNAME o EMAIL (el usuario puede escribir cualquiera)
        cmd.CommandText = _
            "SELECT ID_USER, USERNAME, PASSWORD, EMAIL, USER_PHOTO, USER_TYPE " & _
            "FROM TBL_USER " & _
            "WHERE USERNAME = ? OR EMAIL = ?"

        Dim pUser, pEmail
        Set pUser = cmd.CreateParameter("@username", 200, 1, 50, Left(strInputUser, 50)) ' adVarChar, adParamInput
        Set pEmail = cmd.CreateParameter("@email", 200, 1, 50, Left(strInputUser, 50))
        cmd.Parameters.Append pUser
        cmd.Parameters.Append pEmail

        Set rsUser = Server.CreateObject("ADODB.Recordset")
        rsUser.CursorType = 0  ' adOpenForwardOnly
        rsUser.LockType   = 1  ' adLockReadOnly

        On Error Resume Next
        rsUser.Open cmd
        If Err.Number <> 0 Then
            strErrorMessage = "Error de base de datos: " & Err.Description
            bShowError = True
        End If
        On Error GoTo 0

        If Not bShowError Then
            If rsUser.EOF Then
                '--- No existe el usuario ---
                strErrorMessage = "Usuario o contrase&ntilde;a incorrectos."
                bShowError = True
            Else
                '--- Verificar contrasena: texto plano o SHA1 hash ---
                Dim strDbPassword, bPasswordMatch
                strDbPassword  = rsUser("PASSWORD")
                bPasswordMatch = False

                ' Comparacion directa (texto plano, ej: "12345678")
                If StrComp(strDbPassword, strInputPass, vbBinaryCompare) = 0 Then
                    bPasswordMatch = True
                End If

                ' Comparacion con hash SHA1 (la BD almacena el hash hex en uppercase)
                If Not bPasswordMatch Then
                    Dim strSha1Hash
                    strSha1Hash = UCase(Sha1Hex(strInputPass))
                    If StrComp(strDbPassword, strSha1Hash, vbBinaryCompare) = 0 Then
                        bPasswordMatch = True
                    End If
                End If

                If bPasswordMatch Then
                    '--- CREDENCIALES CORRECTAS: guardar sesion y redirigir -------

                    ' Guardar nombre de usuario (en ambas variables para compatibilidad)
                    Session("Username") = CStr(rsUser("USERNAME"))
                    Session("Usuario")  = CStr(rsUser("USERNAME"))

                    ' Determinar el rol a partir de USER_TYPE
                    '   0 = USER  (normal)
                    '   1 = ADMIN (admin)
                    ' Tambien aceptamos texto en caso de que el catalogo lo guarde asi.
                    Dim intUserType, strRol
                    intUserType = rsUser("USER_TYPE")

                    If IsNull(intUserType) Then
                        strRol = "normal"
                    ElseIf IsNumeric(intUserType) Then
                        If CLng(intUserType) = 1 Then
                            strRol = "admin"
                        Else
                            strRol = "normal"
                        End If
                    Else
                        ' Si USER_TYPE guarda texto ("ADMIN"/"USER")
                        Dim strTypeText
                        strTypeText = UCase(CStr(intUserType))
                        If strTypeText = "ADMIN" Then
                            strRol = "admin"
                        Else
                            strRol = "normal"
                        End If
                    End If

                    Session("Rol")      = strRol
                    Session("UserRole")  = strRol

                    ' Datos extra utiles para el resto de la app
                    Session("UserId")    = CLng(rsUser("ID_USER"))
                    Session("Email")     = ""
                    If Not IsNull(rsUser("EMAIL")) Then
                        Session("Email") = CStr(rsUser("EMAIL"))
                    End If
                    Session("UserPhoto") = ""
                    If Not IsNull(rsUser("USER_PHOTO")) Then
                        Session("UserPhoto") = CStr(rsUser("USER_PHOTO"))
                    End If

                    ' Cerrar recordset antes de redirigir
                    rsUser.Close
                    Set rsUser = Nothing
                    Set cmd    = Nothing

                    ' --- Redirigir al modulo de canciones ---
                    Response.Redirect "canciones.asp"
                    Response.End

                Else
                    strErrorMessage = "Usuario o contrase&ntilde;a incorrectos."
                    bShowError = True
                End If
            End If
        End If

        ' Limpieza si aun abiertos
        If rsUser.State = 1 Then rsUser.Close
        If Not (rsUser Is Nothing) Then Set rsUser = Nothing
        If Not (cmd Is Nothing) Then Set cmd = Nothing
    End If
End If

'==============================================================================
' FUNCION: Sha1Hex
'   Calcula el hash SHA1 de una cadena y lo devuelve como string hex (uppercase)
'   Usa el proveedor CAPICOM si esta disponible; si no, usa .NET via COM interop.
'==============================================================================
Function Sha1Hex(ByVal strInput)
    On Error Resume Next

    ' --- Intento 1: CAPICOM ---
    Dim objHash
    Set objHash = Server.CreateObject("CAPICOM.HashedData")
    If Err.Number = 0 Then
        objHash.Algorithm = 0  ' CAPICOM_HASH_ALGORITHM_SHA1
        objHash.Hash(strInput)
        Sha1Hex = UCase(objHash.Value)
        Set objHash = Nothing
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' --- Intento 2: System.Security.Cryptography via .NET interop ---
    On Error Resume Next
    Dim objSha1, objEncoding
    Set objSha1 = Server.CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    If Err.Number = 0 Then
        Dim arrBytes
        Set objEncoding = Server.CreateObject("System.Text.UTF8Encoding")
        arrBytes = objEncoding.GetBytes_4(strInput)
        Dim arrHash
        arrHash = objSha1.ComputeHash_2((arrBytes))
        ' Convertir bytes a hex
        Dim i, hexStr
        hexStr = ""
        For i = LBound(arrHash) To UBound(arrHash)
            hexStr = hexStr & Right("0" & Hex(arrHash(i)), 2)
        Next
        Sha1Hex = UCase(hexStr)
        Set objSha1 = Nothing
        Set objEncoding = Nothing
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' --- Fallback: devolver la cadena original (comparacion de texto plano) ---
    Sha1Hex = strInput
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PoliMusic &middot; Iniciar Sesi&oacute;n</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        /* ===================================================================
           PALETA DE COLORES - PoliMusic
           Purpura/Indigo oscuro + acentos azul vibrante
           =================================================================== */
        :root {
            --bg-base:       #1E1B2E;
            --bg-surface:    #28253E;
            --bg-elevated:   #332F4F;
            --bg-input:      #2A2742;
            --accent:        #6C5CE7;
            --accent-light:  #8B7FF0;
            --accent-glow:   rgba(108,92,231,0.40);
            --blue-vibrant:  #00B4D8;
            --blue-glow:     rgba(0,180,216,0.25);
            --text-primary:  #F5F3FA;
            --text-secondary:#B4B0CA;
            --text-muted:    #7C7894;
            --border:        rgba(255,255,255,0.08);
            --danger:        #FF5C7C;
            --danger-glow:   rgba(255,92,124,0.25);
            --success:       #4ADE80;
            --radius:        16px;
            --radius-sm:     10px;
            --shadow:        0 12px 40px rgba(0,0,0,0.40);
            --transition:    all 0.25s ease;
        }

        * { margin:0; padding:0; box-sizing:border-box; }

        html, body {
            height: 100%;
        }

        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: var(--bg-base);
            color: var(--text-primary);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            position: relative;
        }

        /* ---- Background decorativo animado ---- */
        .bg-decoration {
            position: fixed; inset: 0;
            overflow: hidden;
            z-index: 0;
            pointer-events: none;
        }
        .bg-decoration::before,
        .bg-decoration::after {
            content: '';
            position: absolute;
            border-radius: 50%;
            filter: blur(80px);
        }
        .bg-decoration::before {
            top: -100px; right: -100px;
            width: 500px; height: 500px;
            background: radial-gradient(circle, var(--accent-glow), transparent 70%);
            animation: float1 8s ease-in-out infinite;
        }
        .bg-decoration::after {
            bottom: -120px; left: -120px;
            width: 450px; height: 450px;
            background: radial-gradient(circle, var(--blue-glow), transparent 70%);
            animation: float2 10s ease-in-out infinite;
        }
        @keyframes float1 {
            0%, 100% { transform: translate(0, 0); }
            50% { transform: translate(-30px, 30px); }
        }
        @keyframes float2 {
            0%, 100% { transform: translate(0, 0); }
            50% { transform: translate(30px, -20px); }
        }

        /* ---- Notas musicales flotantes ---- */
        .floating-note {
            position: absolute;
            font-size: 2.5rem;
            color: rgba(108,92,231,0.12);
            animation: noteFloat 15s linear infinite;
            user-select: none;
        }
        .floating-note:nth-child(1) { left: 10%; top: 20%; animation-delay: 0s; }
        .floating-note:nth-child(2) { left: 85%; top: 15%; animation-delay: 3s; }
        .floating-note:nth-child(3) { left: 20%; top: 75%; animation-delay: 6s; }
        .floating-note:nth-child(4) { left: 75%; top: 70%; animation-delay: 9s; }
        @keyframes noteFloat {
            0% { transform: translateY(0) rotate(0deg); opacity: 0; }
            10% { opacity: 1; }
            90% { opacity: 1; }
            100% { transform: translateY(-120px) rotate(15deg); opacity: 0; }
        }

        /* ===================================================================
           LOGIN CARD
           =================================================================== */
        .login-card {
            position: relative; z-index: 1;
            width: 100%; max-width: 440px;
            background: var(--bg-surface);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            box-shadow: var(--shadow);
            overflow: hidden;
            margin: 20px;
            animation: cardIn 0.5s ease;
        }
        @keyframes cardIn {
            from { transform: translateY(20px) scale(0.98); opacity: 0; }
            to   { transform: translateY(0) scale(1); opacity: 1; }
        }

        /* ---- Header de la card ---- */
        .card-header {
            text-align: center;
            padding: 40px 32px 24px;
        }
        .logo-circle {
            width: 72px; height: 72px;
            margin: 0 auto 20px;
            background: linear-gradient(135deg, var(--accent), var(--blue-vibrant));
            border-radius: 20px;
            display: flex; align-items: center; justify-content: center;
            font-size: 2rem; color: #fff;
            box-shadow: 0 8px 24px var(--accent-glow);
        }
        .card-header h1 {
            font-size: 1.6rem; font-weight: 800;
            letter-spacing: -0.5px;
            margin-bottom: 6px;
        }
        .card-header h1 span { color: var(--blue-vibrant); }
        .card-header p {
            color: var(--text-secondary);
            font-size: 0.9rem;
        }

        /* ---- Formulario ---- */
        .card-body {
            padding: 8px 32px 32px;
        }

        .alert-error {
            display: flex; align-items: center; gap: 10px;
            padding: 12px 16px;
            background: rgba(255,92,124,0.12);
            border: 1px solid rgba(255,92,124,0.3);
            border-radius: var(--radius-sm);
            color: var(--danger);
            font-size: 0.88rem; font-weight: 500;
            margin-bottom: 20px;
            animation: shake 0.4s ease;
        }
        @keyframes shake {
            0%, 100% { transform: translateX(0); }
            25% { transform: translateX(-6px); }
            75% { transform: translateX(6px); }
        }

        .form-group {
            margin-bottom: 20px;
        }
        .form-group label {
            display: flex; align-items: center; gap: 6px;
            font-size: 0.82rem; font-weight: 600;
            color: var(--text-secondary);
            margin-bottom: 8px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .input-wrapper {
            position: relative;
        }
        .input-wrapper .input-icon {
            position: absolute;
            left: 14px; top: 50%;
            transform: translateY(-50%);
            font-size: 1.1rem;
            color: var(--text-muted);
            pointer-events: none;
            transition: var(--transition);
        }
        .form-group input {
            width: 100%;
            padding: 13px 16px 13px 44px;
            background: var(--bg-input);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            color: var(--text-primary);
            font-size: 0.95rem;
            font-family: inherit;
            transition: var(--transition);
        }
        .form-group input::placeholder {
            color: var(--text-muted);
        }
        .form-group input:focus {
            outline: none;
            border-color: var(--accent);
            background: var(--bg-elevated);
            box-shadow: 0 0 0 3px var(--accent-glow);
        }
        .form-group input:focus + .input-focus-glow {
            opacity: 1;
        }
        .form-group input:focus ~ .input-icon {
            color: var(--accent-light);
        }

        .btn-submit {
            width: 100%;
            padding: 14px;
            background: linear-gradient(135deg, var(--accent), var(--accent-light));
            color: #fff;
            border: none;
            border-radius: var(--radius-sm);
            font-size: 1rem; font-weight: 700;
            font-family: inherit;
            cursor: pointer;
            transition: var(--transition);
            letter-spacing: 0.3px;
            position: relative;
            overflow: hidden;
        }
        .btn-submit:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 28px var(--accent-glow);
        }
        .btn-submit:active {
            transform: translateY(0);
        }
        .btn-submit::after {
            content: '';
            position: absolute; inset: 0;
            background: linear-gradient(120deg, transparent, rgba(255,255,255,0.2), transparent);
            transform: translateX(-100%);
            transition: transform 0.6s;
        }
        .btn-submit:hover::after {
            transform: translateX(100%);
        }

        /* ---- Divider ---- */
        .divider {
            display: flex; align-items: center;
            margin: 24px 0;
            gap: 14px;
        }
        .divider::before, .divider::after {
            content: '';
            flex: 1;
            height: 1px;
            background: var(--border);
        }
        .divider span {
            font-size: 0.75rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1px;
        }

        /* ---- Info de credenciales demo ---- */
        .demo-creds {
            background: var(--bg-elevated);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            padding: 14px 16px;
            font-size: 0.8rem;
            color: var(--text-secondary);
        }
        .demo-creds .creds-title {
            font-weight: 600;
            color: var(--accent-light);
            margin-bottom: 8px;
            display: flex; align-items: center; gap: 6px;
        }
        .demo-creds .cred-row {
            display: flex; justify-content: space-between;
            padding: 3px 0;
        }
        .demo-creds .cred-row .label { color: var(--text-muted); }
        .demo-creds .cred-row .value {
            font-family: 'Courier New', monospace;
            color: var(--blue-vibrant);
            font-weight: 600;
        }

        /* ---- Footer ---- */
        .card-footer {
            text-align: center;
            padding: 0 32px 28px;
        }
        .card-footer a {
            color: var(--accent-light);
            text-decoration: none;
            font-size: 0.85rem;
            font-weight: 500;
            transition: var(--transition);
        }
        .card-footer a:hover {
            color: var(--blue-vibrant);
        }
        .card-footer .copyright {
            margin-top: 16px;
            font-size: 0.75rem;
            color: var(--text-muted);
        }

        /* ===================================================================
           RESPONSIVE
           =================================================================== */
        @media (max-width: 480px) {
            .login-card { margin: 0; border-radius: 0; min-height: 100vh; }
            .card-header { padding: 32px 24px 16px; }
            .card-body   { padding: 8px 24px 24px; }
            .card-footer { padding: 0 24px 24px; }
        }
    </style>
</head>
<body>

    <!-- Background decorativo -->
    <div class="bg-decoration">
        <div class="floating-note">&#9835;</div>
        <div class="floating-note">&#9834;</div>
        <div class="floating-note">&#9833;</div>
        <div class="floating-note">&#9836;</div>
    </div>

    <!-- ===================== LOGIN CARD ===================== -->
    <div class="login-card">

        <!-- Header -->
        <div class="card-header">
            <div class="logo-circle">&#9835;</div>
            <h1>Poli<span>Music</span></h1>
            <p>Tu plataforma de m&uacute;sica &mdash; Inicia sesi&oacute;n para continuar</p>
        </div>

        <!-- Body / Form -->
        <div class="card-body">

            <% If bShowError Then %>
                <div class="alert-error">
                    <span>&#9888;</span>
                    <%= strErrorMessage %>
                </div>
            <% End If %>

            <form method="POST" action="login.asp" id="loginForm" autocomplete="on">
                <div class="form-group">
                    <label for="username">&#128100; Usuario o Correo</label>
                    <div class="input-wrapper">
                        <input type="text"
                               id="username"
                               name="username"
                               placeholder="usuario o correo@epn.edu.ec"
                               required
                               autofocus
                               maxlength="50"
                               value="<%= Server.HTMLEncode(Request.Form("username")) %>">
                        <span class="input-icon">&#9993;</span>
                    </div>
                </div>

                <div class="form-group">
                    <label for="password">&#128274; Contrase&ntilde;a</label>
                    <div class="input-wrapper">
                        <input type="password"
                               id="password"
                               name="password"
                               placeholder="&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull;"
                               required
                               maxlength="50">
                        <span class="input-icon">&#128274;</span>
                    </div>
                </div>

                <button type="submit" class="btn-submit">
                    Iniciar Sesi&oacute;n &nbsp;&rarr;
                </button>
            </form>

            <!-- Credenciales demo -->
            <div class="divider"><span>Credenciales de prueba</span></div>
            <div class="demo-creds">
                <div class="creds-title">&#9432; Usuarios registrados en BDD_PoliMusic</div>
                <div class="cred-row">
                    <span class="label">Administrador:</span>
                    <span class="value">Belen.Cholango / 12345678</span>
                </div>
                <div class="cred-row">
                    <span class="label">Admin (demo):</span>
                    <span class="value">userAdmin / 12345678</span>
                </div>
                <div class="cred-row">
                    <span class="label">Usuario normal:</span>
                    <span class="value">userNormal / 12345678</span>
                </div>
            </div>

        </div>

        <!-- Footer -->
        <div class="card-footer">
            <a href="#">&#191;Olvidaste tu contrase&ntilde;a?</a>
            <div class="copyright">&copy; 2026 PoliMusic &middot; Proyecto Acad&eacute;mico</div>
        </div>

    </div>

</body>
</html>