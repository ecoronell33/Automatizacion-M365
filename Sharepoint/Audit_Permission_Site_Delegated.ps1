$TenantName = "xxxxx.onmicrosoft.com"
$ClientId   = "xxxxxxxxx" 

# URL del Centro de Administración para poder listar todos los sitios
$AdminUrl = "https://xxxxx-admin.sharepoint.com"
$Fecha    = Get-Date -Format "yyyyMMdd-HHmm"

$CsvMiembros = ".\Auditoria-SharePoint-Permisos-$Fecha.csv"
$CsvErrores  = ".\Auditoria-SharePoint-Errores-$Fecha.csv"

$Resultado = [System.Collections.Generic.List[object]]::new()
$Errores   = [System.Collections.Generic.List[object]]::new()

Import-Module PnP.PowerShell

# Conexión solo para consulta en el Admin Center
Write-Host "Conectando al Centro de Administración..." -ForegroundColor Cyan
$AdminConnection = Connect-PnPOnline `
    -Url $AdminUrl `
    -Interactive `
    -ClientId $ClientId `
    -ReturnConnection

# Consulta todas las colecciones de sitios
$Sitios = Get-PnPTenantSite -Connection $AdminConnection

foreach ($Sitio in $Sitios) {

    Write-Host "Auditando colección: $($Sitio.Url)" -ForegroundColor Cyan

    try {
        $RootConnection = Connect-PnPOnline `
            -Url $Sitio.Url `
            -Interactive `
            -ClientId $ClientId `
            -ReturnConnection

        $RootWeb = Get-PnPWeb `
            -Includes Title,Url `
            -Connection $RootConnection

        # Consulta todos los subsitios recursivamente
        $Subsitios = @(
            Get-PnPSubWeb `
                -Recurse `
                -Includes Title,Url `
                -Connection $RootConnection
        )

        $Webs = @($RootWeb) + $Subsitios

        foreach ($Web in $Webs) {

            Write-Host "  Sitio/Subsitio: $($Web.Url)" -ForegroundColor Yellow

            try {
                if ($Web.Url -eq $Sitio.Url) {
                    $WebConnection = $RootConnection
                }
                else {
                    $WebConnection = Connect-PnPOnline `
                        -Url $Web.Url `
                        -Interactive `
                        -ClientId $ClientId `
                        -ReturnConnection
                }

                # Consultamos las asignaciones de roles (Permisos) del sitio actual
                $CurrentWeb = Get-PnPWeb -Connection $WebConnection
                $RoleAssignments = Get-PnPProperty -ClientObject $CurrentWeb -Property RoleAssignments

                foreach ($RoleAssign in $RoleAssignments) {
                    
                    try {
                        # Cargamos las propiedades del miembro y sus permisos
                        Get-PnPProperty -ClientObject $RoleAssign -Property RoleDefinitionBindings, Member
                        
                        $Member = $RoleAssign.Member
                        
                        # Extraemos los roles de SharePoint
                        $Permisos = ($RoleAssign.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join ", "
                        
                        # Traducimos los roles comunes al español (Lector, Editor, etc.)
                        $Permisos = $Permisos -replace '\bRead\b', 'Lector' `
                                              -replace '\bEdit\b', 'Editor' `
                                              -replace '\bFull Control\b', 'Control Total' `
                                              -replace '\bContribute\b', 'Colaborador' `
                                              -replace '\bDesign\b', 'Diseñador'
                        
                        # Si el miembro es un Grupo de SharePoint, buscamos quién está adentro
                        if ($Member.PrincipalType -eq "SharePointGroup") {
                            
                            $Members = @(Get-PnPGroupMember -Identity $Member.LoginName -Connection $WebConnection -ErrorAction Stop)

                            if ($Members.Count -eq 0) {
                                $Resultado.Add([PSCustomObject]@{
                                    ColeccionDeSitios = $Sitio.Url
                                    SitioOSubsitio    = $Web.Url
                                    TituloDelSitio    = $Web.Title
                                    GrupoOUsuario     = $Member.Title
                                    Rol               = $Permisos
                                    Miembro           = ""
                                    Correo            = ""
                                    Usuario           = ""
                                    Tipo              = "Grupo de SharePoint"
                                    GrupoVacio        = "Sí"
                                })
                            }
                            else {
                                foreach ($User in $Members) {
                                    $Resultado.Add([PSCustomObject]@{
                                        ColeccionDeSitios = $Sitio.Url
                                        SitioOSubsitio    = $Web.Url
                                        TituloDelSitio    = $Web.Title
                                        GrupoOUsuario     = $Member.Title
                                        Rol               = $Permisos
                                        Miembro           = $User.Title
                                        Correo            = $User.Email
                                        Usuario           = $User.LoginName
                                        Tipo              = $User.PrincipalType
                                        GrupoVacio        = "No"
                                    })
                                }
                            }
                        }
                        # Si NO es un grupo de SharePoint, significa que es un Usuario o Grupo de Entra ID con permisos DIRECTOS
                        else {
                            $Resultado.Add([PSCustomObject]@{
                                ColeccionDeSitios = $Sitio.Url
                                SitioOSubsitio    = $Web.Url
                                TituloDelSitio    = $Web.Title
                                GrupoOUsuario     = "Permiso Directo"
                                Rol               = $Permisos
                                Miembro           = $Member.Title
                                Correo            = $Member.Email
                                Usuario           = $Member.LoginName
                                Tipo              = $Member.PrincipalType
                                GrupoVacio        = "N/A"
                            })
                        }
                    }
                    catch {
                        $Errores.Add([PSCustomObject]@{
                            Sitio = $Web.Url
                            Grupo = $RoleAssign.Member.Title
                            Etapa = "Consulta de miembros o roles"
                            Error = $_.Exception.Message
                        })
                    }
                }
            }
            catch {
                $Errores.Add([PSCustomObject]@{
                    Sitio = $Web.Url
                    Grupo = ""
                    Etapa = "Consulta del sitio/subsitio"
                    Error = $_.Exception.Message
                })
            }
        }
    }
    catch {
        $Errores.Add([PSCustomObject]@{
            Sitio = $Sitio.Url
            Grupo = ""
            Etapa = "Conexión con la colección"
            Error = $_.Exception.Message
        })
    }
}

# Exportación de resultados a CSV
$Resultado |
    Sort-Object ColeccionDeSitios, SitioOSubsitio, GrupoOUsuario, Miembro |
    Export-Csv `
        -Path $CsvMiembros `
        -Delimiter ";" `
        -Encoding utf8BOM `
        -NoTypeInformation

$Errores |
    Export-Csv `
        -Path $CsvErrores `
        -Delimiter ";" `
        -Encoding utf8BOM `
        -NoTypeInformation

Write-Host "`nAuditoría finalizada con roles y permisos." -ForegroundColor Green
Write-Host "Reporte: $CsvMiembros"
Write-Host "Errores: $CsvErrores"
