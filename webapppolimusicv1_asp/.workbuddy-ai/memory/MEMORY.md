# PoliMusic v1 — Notas del proyecto

## Stack
- ASP Clasico (VBScript) sobre IIS + SQL Server (base `BDD_PoliMusic`)
- Driver: `SQLOLEDB.1` / SQL Server local `(local)`
- Credenciales BD: `usr_polimusic` / `Politecnica1` (login SQL con rol db_owner)

## Estructura
```
webapppolimusicv1_asp/
├── login.asp          # Autenticacion (POST) -> Session -> canciones.asp
├── logout.asp         # Session.Abandon() -> login.asp
├── canciones.asp      # Listado + CRUD (solo admin)
├── includes/
│   └── conexion.asp   # ADODB.Connection compartida
├── SongFiles/         # .mp3 (bensound-*.mp3)
├── albumImages/       # Portadas de album (album1-5.png)
├── ArtistImages/      # Fotos de artistas (artist1-4.jpg, artist5.png)
├── PlaylistImages/    # Portadas de playlists (playlist1-5.png, default.jpg)
├── UserImages/        # Fotos de perfil (default.jpg, root.png)
└── BDD_PoliMusic.sql  # Script de la BD (encoding UTF-16LE con BOM)
```

## Roles
- `TBL_USER.USER_TYPE`: 0 = USER (normal), 1 = ADMIN
- En sesion se guardan strings: `Session("Rol")` = `"normal"` o `"admin"`
- El rol "admin" habilita los botones Crear/Editar/Eliminar

## Convenciones importantes
1. **Las Functions VBScript deben declararse ANTES de su primer uso.** ASP Clasico evalua el codigo en orden; una Function definida al final del archivo no existe cuando se llama arriba.
2. `BDD_PoliMusic.sql` esta en **UTF-16LE** (BOM `fffe`). Para editarlo: `iconv -f UTF-16LE -t UTF-8` -> editar -> `iconv -f UTF-8 -t UTF-16LE` de vuelta.
3. Conexion por consola: `sqlcmd -S localhost -U usr_polimusic -P "Politecnica1" -d BDD_PoliMusic -C` — el flag `-C` es obligatorio (ODBC Driver 18 exige confiar en el certificado del servidor).
4. **Evitar JOIN entre TBL_PLAYER y TBL_TRACKLIST** en la misma query: el producto cartesiano duplica las canciones. Usar subqueries con `TOP 1`.
5. Las FKs hacia `TBL_SONG` tienen `ON DELETE CASCADE` (TBL_TRACKLIST, TBL_PLAYER, TBL_PLAYLIST_SONG). El ASP tambien hace borrado explicito en cascada por si acaso.
6. Paths de assets en la BD usan formato `~/carpeta/archivo`. La funcion `NormalizeAssetPath` los convierte a rutas web; el parametro `realFolder` indica la carpeta real del proyecto.
7. `Option Explicit` activo: toda variable debe declararse con `Dim`.
8. Consultas siempre parametrizadas con `ADODB.Command` + `CreateParameter` (evita SQL injection).
9. Passwords en `TBL_USER`: algunos en texto plano, otros en SHA1 hex uppercase. `login.asp` verifica ambos formatos (CAPICOM o .NET interop).

## Usuarios de prueba
| USERNAME | PASSWORD | USER_TYPE | Email |
|---|---|---|---|
| Belen.Cholango | 12345678 | 1 (admin) | belen.cholango@epn.edu.ec |
| userAdmin | 12345678 | 1 (admin) | usrfakeadmin@epn.edu.ec |
| userNormal | 12345678 | 0 (normal) | usrnormal@epn.edu.ec |
| username1 | 7C222FB2927D... (SHA1) | 0 (normal) | username1@epn.edu.ec |

## Paleta de diseño
- Fondo base: `#1E1B2E` (purpura/indigo oscuro)
- Superficie: `#28253E` / elevada `#332F4F`
- Acento: `#6C5CE7` (purpura) + `#00B4D8` (azul vibrant)
- Texto: `#F5F3FA` / `#B4B0CA` / muted `#7C7894`
- Peligro: `#FF5C7C`, exito: `#4ADE80`
- Tipografia: Inter (Google Fonts)

## Flujo
`login.asp` (POST valida) -> `canciones.asp` -> `logout.asp` -> `login.asp`
