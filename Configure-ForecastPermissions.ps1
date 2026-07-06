# Script de PowerShell para configurar permisos en SharePoint Online - Forecast 2026
# Requisito: Módulo PnP.PowerShell (el script lo instala si no está presente)

# ── CONFIGURACIÓN GENERAL ──────────────────────────────────────────────────
$SiteUrl = "https://provexpress.sharepoint.com/sites/ProvexpressIntranet"
$LibraryName = "Documentos compartidos"
$BasePath = "COMERCIAL/FORECAST 2026"

# Administradores con acceso total (pueden editar todo en todas las carpetas)
$Admins = @(
    "c.estrategica@provexpress.com.co",
    "juannovoa@provexpress.com.co",
    "oscar.perez@provexpress.com.co",
    "especialista.preventa@provexpress.com.co",
    "preventa.software@provexpress.com.co"
)

# Estructura de grupos de directores, ejecutivos y sus archivos Excel
$EstructuraComercial = @(
    # GRUPO 1: Rafael Novoa
    @{
        DirectorEmail = "rafael.novoa@provexpress.com.co"
        FolderName    = "Grupo Rafael Novoa"
        Executives    = @{
            "rosmira.rojas@provexpress.com.co" = "Rosmira Rojas.xlsx"
            "mario.reyes@provexpress.com.co"   = "Mario Reyes.xlsx"
            "wilson.sanchez@provexpress.com.co" = "Wilson Fernando Sánchez.xlsx"
            "maria.cruz@provexpress.com.co"     = "Maria Eugenia Cruz.xlsx"
            "javier.cortes@provexpress.com.co"   = "Javier Cortés.xlsx"
            "rosa.mendoza@provexpress.com.co"   = "Rosa María Mendoza.xlsx"
            "mariela.ramirez@provexpress.com.co" = "Mariela Ramírez.xlsx"
            "jenny.gonzalez@provexpress.com.co"  = "Jenny Gónzalez.xlsx"
            "julieth.galindo@provexpress.com.co" = "Julieth Galindo.xlsx"
        }
        UnitSupport   = @()  # No hay soporte asignado a nivel de unidad en grupo 1
    },
    # GRUPO 2: Angélica Caballero
    @{
        DirectorEmail = "angelica.caballero@provexpress.com.co"
        FolderName    = "Grupo Angelica Caballero"
        Executives    = @{
            "angela.torres@provexpress.com.co"       = "Ángela Torres.xlsx"
            "andrea.vargas@provexpress.com.co"       = "Yurani Vargas.xlsx"
            "alejandra.velasquez@provexpress.com.co" = "Alejandra Velásquez.xlsx"
            "fernando.quinonez@provexpress.com.co"   = "Fernando Quiñonez.xlsx"
            "johana.mojica@provexpress.com.co"       = "Jasbleidy Mójica.xlsx"
            "johanna.jaime@provexpress.com.co"       = "Johanna Jaime.xlsx"
            "dayana.chala@provexpress.com.co"        = "Dayana Chala.xlsx"
            "yovanny.herrera@provexpress.com.co"     = "Yovanny Herrera.xlsx"
            "cesar.cespedes@provexpress.com.co"      = "César Cespedes.xlsx"
            "daniel.galindo@provexpress.com.co"      = "Daniel Galindo.xlsx"
        }
        UnitSupport   = @("soporte.comercial@provexpress.com.co") # Karen Cagua
    },
    # GRUPO 3: Óscar Beltrán
    @{
        DirectorEmail = "oscar.beltran@provexpress.com.co"
        FolderName    = "Grupo Oscar Beltran"
        Executives    = @{
            "paola.garcia@provexpress.com.co"     = "Gina García.xlsx"
            "karen.carrillo@provexpress.com.co"   = "Karent Carrillo.xlsx"
            "lington.linares@provexpress.com.co"  = "Lington Linares.xlsx"
            "angelica.alvarez@provexpress.com.co" = "Angélica Álvarez.xlsx"
            "andres.pena@provexpress.com.co"      = "Andrés Peña.xlsx"
            "tatiana.parra@provexpress.com.co"     = "Tatiana Parra.xlsx"
            "claudia.triana@provexpress.com.co"   = "Claudia Triana.xlsx"
            "dilma.cuesta@provexpress.com.co"     = "Dilma Cuesta.xlsx"
            "juan.martinez@provexpress.com.co"    = "Juan Martínez.xlsx"
        }
        UnitSupport   = @("soporte.comercial2@provexpress.com.co") # Alexandra Vargas
    },
    # GRUPO 4: Miller Romero
    @{
        DirectorEmail = "miller.romero@provexpress.com.co"
        FolderName    = "Grupo Miller Romero"
        Executives    = @{
            "astrid.jimenez@provexpress.com.co"   = "Astrid Jiménez.xlsx"
            "maria.briceno@provexpress.com.co"    = "María Paola Briceño.xlsx"
            "dafne.ruiz@provexpress.com.co"       = "Dafne Lizeth Ruiz.xlsx"
            "jessica.valencia@provexpress.com.co" = "Jessica Valencia.xlsx"
            "jhonatan.acevedo@provexpress.com.co" = "Jhonatan Acevedo.xlsx"
            "camilo.hernandez@provexpress.com.co" = "Jhonatan Camilo Hernández.xlsx"
            "yeison.urrego@provexpress.com.co"    = "Yeison Urrego.xlsx"
            "diana.castro@provexpress.com.co"     = "Diana Catalina Castro.xlsx"
        }
        UnitSupport   = @() # No tiene soporte de unidad, sino soporte comercial específico
    }
)

# Mapeo de Soporte Comercial específico (soporta a ciertos ejecutivos)
$SoporteComercialEspecifico = @{
    "soporte.comercial4@provexpress.com.co" = @("yeison.urrego@provexpress.com.co") # Nury Marcela Vargas
    "soporte.comercial3@provexpress.com.co" = @("camilo.hernandez@provexpress.com.co") # Janira Alejandra Maldonado
    "soporte.comercial5@provexpress.com.co" = @("maria.briceno@provexpress.com.co", "jessica.valencia@provexpress.com.co") # Johanna Alcocer
}

# ── INICIALIZACIÓN Y CONEXIÓN ──────────────────────────────────────────────
Write-Host "Verificando el módulo PnP.PowerShell..." -ForegroundColor Cyan
if (-not (Get-Module -Name PnP.PowerShell -ListAvailable)) {
    Write-Host "Instalando módulo PnP.PowerShell para el usuario actual..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}

Write-Host "Conectando a SharePoint Online ($SiteUrl)..." -ForegroundColor Cyan
# Conexión interactiva (abrirá el navegador para autenticación de Microsoft 365)
Connect-PnPOnline -Url $SiteUrl -Interactive

# ── PROCESAMIENTO DE PERMISOS ──────────────────────────────────────────────
foreach ($grupo in $EstructuraComercial) {
    $director = $grupo.DirectorEmail
    $folderName = $grupo.FolderName
    $executives = $grupo.Executives
    $unitSupports = $grupo.UnitSupport
    $folderRelativePath = "$BasePath/$folderName"

    Write-Host "`n=== PROCESANDO GRUPO: $folderName ===" -ForegroundColor Green

    # 1. Obtener la carpeta del Director/Grupo
    Write-Host "Obteniendo carpeta: $folderRelativePath" -ForegroundColor Gray
    try {
        $folder = Get-PnPFolder -Url $folderRelativePath -Includes ListItemAllFields -ErrorAction Stop
    } catch {
        Write-Warning "La carpeta '$folderRelativePath' no existe o no se pudo acceder. Saltando grupo..."
        continue
    }

    $folderItemId = $folder.ListItemAllFields.Id
    if (-not $folderItemId) {
        Write-Warning "No se pudo obtener el ID del elemento de la carpeta '$folderName'. Saltando..."
        continue
    }

    # Romper herencia de permisos en la carpeta del grupo
    Write-Host "Rompiendo herencia de la carpeta y limpiando permisos existentes..." -ForegroundColor Yellow
    Set-PnPListItemPermission -List $LibraryName -Identity $folderItemId -BreakInheritance -ClearExisting

    # Asignar permisos en la carpeta del grupo:
    # - Administradores: Editar
    foreach ($admin in $Admins) {
        Set-PnPListItemPermission -List $LibraryName -Identity $folderItemId -User $admin -AddRole "Edit"
    }
    # - Director: Editar
    Set-PnPListItemPermission -List $LibraryName -Identity $folderItemId -User $director -AddRole "Edit"

    # - Soporte de Unidad (si aplica): Leer (para poder listar/entrar a la carpeta)
    foreach ($support in $unitSupports) {
        Set-PnPListItemPermission -List $LibraryName -Identity $folderItemId -User $support -AddRole "Read"
    }

    # - Soporte Comercial Específico que apoye a ejecutivos de este grupo: Leer (para poder entrar)
    foreach ($supportEmail in $SoporteComercialEspecifico.Keys) {
        $execsSoportados = $SoporteComercialEspecifico[$supportEmail]
        foreach ($execEmail in $execsSoportados) {
            if ($executives.ContainsKey($execEmail)) {
                Set-PnPListItemPermission -List $LibraryName -Identity $folderItemId -User $supportEmail -AddRole "Read"
                break # Solo necesitamos agregarlo una vez a la carpeta
            }
        }
    }

    # - Ejecutivos de este grupo: Leer (para que puedan ver la estructura y entrar)
    foreach ($execEmail in $executives.Keys) {
        Set-PnPListItemPermission -List $LibraryName -Identity $folderItemId -User $execEmail -AddRole "Read"
    }

    Write-Host "Permisos de la carpeta '$folderName' configurados con éxito." -ForegroundColor DarkGreen

    # 2. Procesar los archivos individuales de los ejecutivos
    foreach ($execEmail in $executives.Keys) {
        $fileName = $executives[$execEmail]
        $fileRelativePath = "$folderRelativePath/$fileName"

        Write-Host "   Procesando archivo ejecutivo: $fileName" -ForegroundColor Gray
        try {
            $file = Get-PnPFile -Url $fileRelativePath -AsListItem -ErrorAction Stop
        } catch {
            Write-Warning "   El archivo '$fileRelativePath' no existe. Saltando permisos de este archivo..."
            continue
        }

        $fileItemId = $file.Id

        # Romper herencia de permisos en el archivo Excel del ejecutivo
        Set-PnPListItemPermission -List $LibraryName -Identity $fileItemId -BreakInheritance -ClearExisting

        # Asignar permisos en el archivo Excel:
        # - Administradores: Editar
        foreach ($admin in $Admins) {
            Set-PnPListItemPermission -List $LibraryName -Identity $fileItemId -User $admin -AddRole "Edit"
        }
        # - Director: Editar
        Set-PnPListItemPermission -List $LibraryName -Identity $fileItemId -User $director -AddRole "Edit"

        # - Ejecutivo dueño del archivo: Editar
        Set-PnPListItemPermission -List $LibraryName -Identity $fileItemId -User $execEmail -AddRole "Edit"

        # - Soporte de Unidad (si aplica): Editar (apoya a toda la unidad)
        foreach ($support in $unitSupports) {
            Set-PnPListItemPermission -List $LibraryName -Identity $fileItemId -User $support -AddRole "Edit"
        }

        # - Soporte Comercial Específico (si este ejecutivo tiene soporte asignado): Editar
        foreach ($supportEmail in $SoporteComercialEspecifico.Keys) {
            $execsSoportados = $SoporteComercialEspecifico[$supportEmail]
            if ($execsSoportados -contains $execEmail) {
                Set-PnPListItemPermission -List $LibraryName -Identity $fileItemId -User $supportEmail -AddRole "Edit"
            }
        }

        Write-Host "   -> Permisos aplicados en '$fileName' (Ejecutivo: $execEmail, Director, Admins y Soportes asignados)." -ForegroundColor DarkGreen
    }
}

Write-Host "`n=== PROCESO FINALIZADO CON ÉXITO ===" -ForegroundColor Green
Write-Host "Todos los permisos fueron configurados en SharePoint de forma silenciosa." -ForegroundColor Cyan
