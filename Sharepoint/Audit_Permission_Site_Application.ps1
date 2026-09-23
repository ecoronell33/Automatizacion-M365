#Para ejecutar este script, se debe acompañar con la siguiente sentencia:
#.\Report_Sites.ps1 -SiteUrl "https://lantester.sharepoint.com/sites/Cloud" -ClientId "xxxxxxxxxxxxxxxxxx" -Tenant "xxxxx.onmicrosoft.com" -Thumbprint "xxxxxxx" -OutputPath ".\Logs"


[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="URL del sitio de SharePoint a analizar")]
    [string]$SiteUrl,

    [Parameter(Mandatory=$false, HelpMessage="Ruta donde se guardarán los logs. Por defecto crea una carpeta Logs donde se ejecuta el script.")]
    [string]$OutputPath = ".\Logs",

    # --- OPCIONES DE AUTENTICACIÓN DESATENDIDA ---
    
    [Parameter(Mandatory=$false, HelpMessage="ID de la aplicación de Entra ID / Azure AD")]
    [string]$ClientId,

    [Parameter(Mandatory=$false, HelpMessage="Nombre del tenant (ej. contoso.onmicrosoft.com)")]
    [string]$Tenant,

    [Parameter(Mandatory=$false, HelpMessage="Huella digital (Thumbprint) del certificado (Recomendado)")]
    [string]$Thumbprint,

    [Parameter(Mandatory=$false, HelpMessage="Secreto del cliente (Client Secret)")]
    [string]$ClientSecret,

    [Parameter(Mandatory=$false, HelpMessage="Usar Identidad Administrada (Ej. para Azure Automation o Azure Functions)")]
    [switch]$UseManagedIdentity
)

# No usamos Clear-Host en automatización porque borra el log de la consola de la herramienta (ej. Azure Automation)

$properties = @{SiteUrl='';SiteTitle='';ListTitle='';Type='';RelativeUrl='';ParentGroup='';MemberType='';MemberName='';MemberLoginName='';Roles='';}
 
$dateTime = (Get-Date).toString("dd-MM-yyyy-hh-ss")
$excludeLimitedAccess = $true
$includeListsItems = $true

$global:siteTitle = ""
$global:siteUrl = ""

# Exclude certain libraries
$ExcludedLibraries = @("Form Templates", "Preservation Hold Library", "Site Assets", "Images", "Pages", "Settings", "Videos","Timesheet",
  "Site Collection Documents", "Site Collection Images", "Style Library", "AppPages", "Apps for SharePoint", "Apps for Office")

$global:permissions = @()
$global:sharingLinks = @()

function Get-ListItems_WithUniquePermissions {
    param(
        [Parameter(Mandatory)]
        [Microsoft.SharePoint.Client.List]$List
    )
    $selectFields = "ID,HasUniqueRoleAssignments,FileRef,FileLeafRef,FileSystemObjectType"
 
    $Url = $global:siteUrl + '/_api/web/lists/getbytitle(''' + $($list.Title) + ''')/items?$select=' + $($selectFields)
    $nextLink = $Url
    $listItems = @()
    
    while($nextLink){  
        do {
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextLink -Method Get
                $Stoploop = $true
            }
            catch {
                Write-Host "An error occured: $_  : Retrying" -ForegroundColor Red
                $Stoploop = $false
                Start-Sleep -Seconds 30
            }
        }
        While ($Stoploop -eq $false)
  
        $listItems += $response.value | Where-Object {$_.HasUniqueRoleAssignments -eq $true}
        if($response.'odata.nextlink'){
            $nextLink = $response.'odata.nextlink'
        } else {
            $nextLink = $null
        }
    }
    return $listItems
}

Function PermissionObject($_object,$_type,$_relativeUrl,$_siteUrl,$_siteTitle,$_listTitle,$_memberType,$_parentGroup,$_memberName,$_memberLoginName,$_roleDefinitionBindings)
{
    $permission = New-Object -TypeName PSObject -Property $properties
    $permission.SiteUrl = $_siteUrl
    $permission.SiteTitle = $_siteTitle
    $permission.ListTitle = $_listTitle
    $permission.Type = $_Type -eq 1 ? "Folder" : $_Type -eq 0 ? "File" : $_Type
    $permission.RelativeUrl = $_relativeUrl
    $permission.MemberType = $_memberType
    $permission.ParentGroup = $_parentGroup
    $permission.MemberName = $_memberName
    $permission.MemberLoginName = $_memberLoginName
    $permission.Roles = $_roleDefinitionBindings -join ","
    $global:permissions += $permission
}

Function Extract-Guid ($inputString) {
    $splitString = $inputString -split '\|'
    return $splitString[2].TrimEnd('_o')
}

Function QueryUniquePermissionsByObject($_ctx,$_object,$_Type,$_RelativeUrl,$_siteUrl,$_siteTitle,$_listTitle)
{
    $roleAssignments = Get-PnPProperty -ClientObject $_object -Property RoleAssignments

    foreach($roleAssign in $roleAssignments){
        Get-PnPProperty -ClientObject $roleAssign -Property RoleDefinitionBindings,Member
        $PermissionLevels = $roleAssign.RoleDefinitionBindings | Select-Object -ExpandProperty Name
        
        if($excludeLimitedAccess -eq $true){
            $PermissionLevels = ($PermissionLevels | Where-Object { $_ -ne "Limited Access"}) -join ","  
        }
        $Users = Get-PnPProperty -ClientObject ($roleAssign.Member) -Property Users -ErrorAction SilentlyContinue
        
        $AccessType = $roleAssign.RoleDefinitionBindings.Name
        $MemberType = $roleAssign.Member.GetType().Name
        $PermissionType = $roleAssign.Member.PrincipalType  
        
        if($_Type -eq 0){
            $sharingLinks = Get-PnPFileSharingLink -Identity $_object.FieldValues["FileRef"]
        }
        if($_Type -eq 1){
            $sharingLinks = Get-PnPFolderSharingLink -Folder $_object.FieldValues["FileRef"]
        }

        If($PermissionLevels.Length -gt 0) {
            $MemberType = $roleAssign.Member.GetType().Name
            
            If ($roleAssign.Member.Title -like "SharingLinks*")
            {
                if($sharingLinks){
                    $sharingLinks | Where-Object {$roleAssign.Member.Title -match $_.Id } | ForEach-Object {
                        If ($Users.Count -gt 0) 
                        {
                            ForEach ($User in $Users)
                            {
                                PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle $_listTitle "Sharing Links" $roleAssign.Member.LoginName $user.Title $User.LoginName $_.Link.Type
                            }
                        } 
                        else {
                            PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle $_listTitle "Sharing Links" $roleAssign.Member.LoginName $_.Link.Scope "" $_.Link.Type
                        }
                    }  
                }
            }
            ElseIf($MemberType -eq "Group" -or $MemberType -eq "User")
            { 
                $MemberName = $roleAssign.Member.Title
                $MemberLoginName = $roleAssign.Member.LoginName    
                if($MemberType -eq "User")
                {
                    $ParentGroup = "NA"
                }
                else
                {
                    $ParentGroup = $MemberName
                }
                PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle $_listTitle $MemberType $ParentGroup $MemberName $MemberLoginName $PermissionLevels
            }

            if($_Type -eq "Site" -and $MemberType -eq "Group")
            {
                If($PermissionType -eq "SharePointGroup") {  
                    $groupUsers = Get-PnPGroupMember -Identity $roleAssign.Member.LoginName                  
                    $groupUsers | ForEach-Object { 
                        if ($_.LoginName.StartsWith("c:0o.c|federateddirectoryclaimprovider|") -and $_.LoginName.EndsWith("_0")) {
                            $guid = Extract-Guid $_.LoginName
                            Get-PnPMicrosoft365GroupOwners -Identity $guid | ForEach-Object {
                                $user = $_
                                PermissionObject $_object "Site" $_RelativeUrl $_siteUrl $_siteTitle "" "GroupMember" $roleAssign.Member.LoginName $user.DisplayName $user.UserPrincipalName $PermissionLevels
                            }
                        }
                        elseif ($_.LoginName.StartsWith("c:0o.c|federateddirectoryclaimprovider|")) {
                            $guid = Extract-Guid $_.LoginName
                            Get-PnPMicrosoft365GroupMembers -Identity $guid | ForEach-Object {
                                $user = $_
                                PermissionObject $_object "Site" $_RelativeUrl $_siteUrl $_siteTitle "" "GroupMember" $roleAssign.Member.LoginName $user.DisplayName $user.UserPrincipalName $PermissionLevels
                            }
                        }

                        PermissionObject $_object "Site" $_RelativeUrl $_siteUrl $_siteTitle "" "GroupMember" $roleAssign.Member.LoginName $_.Title $_.LoginName $PermissionLevels   
                    }
                }
            } 
        }      
    }
}

Function QueryUniquePermissions($_web)
{
    Write-Host "Querying web $($_web.Title)"
    $global:siteUrl = $_web.Url
    Write-Host $global:siteUrl -ForegroundColor "Red"
    $global:siteTitle = $_web.Title
    
    $ll = Get-PnPList -Includes BaseType, Hidden, Title, HasUniqueRoleAssignments, RootFolder | Where-Object {$_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries }
    Write-Host "Number of lists $($ll.Count)"

    QueryUniquePermissionsByObject $_web $_web "Site" "" $global:siteUrl $siteTitle ""
 
    foreach($list in $ll)
    {      
        $listUrl = $list.RootFolder.ServerRelativeUrl
        if($list.Hidden -ne $True)
        { 
            Write-Host $list.Title -ForegroundColor "Yellow"
            $listTitle = $list.Title
            
            if($list.HasUniqueRoleAssignments -eq $True)
            { 
                $Type = $list.BaseType.ToString()
                QueryUniquePermissionsByObject $_web $list $Type $listUrl $global:siteUrl $siteTitle $listTitle
            }
            
            if($includeListsItems){         
                $collListItem = Get-ListItems_WithUniquePermissions -List $list
                $count = $collListItem.Count
                Write-Host "Number of items with unique permissions: $count within list $listTitle" 
                foreach($item in $collListItem) 
                {
                    $Type = $item.FileSystemObjectType
                    $fileUrl = $item.FileRef  
                    $i = Get-PnPListItem -List $list -Id $item.ID
                    QueryUniquePermissionsByObject $_web $i $Type $fileUrl $global:siteUrl $siteTitle $listTitle
                } 
            }
        }
    }
}

# ----------------- MAIN EJECUCIÓN -----------------

# Validar y crear la carpeta de salida
if(!(Test-Path $OutputPath)){
    New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null
    Write-Host "Created missing output directory at $OutputPath" -ForegroundColor Green
}

# Lógica de Conexión Desatendida
try {
    if ($UseManagedIdentity) {
        Write-Host "Connecting using Managed Identity..." -ForegroundColor Cyan
        Connect-PnPOnline -Url $SiteUrl -ManagedIdentity -ErrorAction Stop
    }
    elseif (![string]::IsNullOrEmpty($ClientId) -and ![string]::IsNullOrEmpty($Tenant) -and ![string]::IsNullOrEmpty($Thumbprint)) {
        Write-Host "Connecting using App Registration (Certificate)..." -ForegroundColor Cyan
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint -ErrorAction Stop
    }
    elseif (![string]::IsNullOrEmpty($ClientId) -and ![string]::IsNullOrEmpty($Tenant) -and ![string]::IsNullOrEmpty($ClientSecret)) {
        Write-Host "Connecting using App Registration (Client Secret)..." -ForegroundColor Cyan
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $Tenant -ClientSecret $ClientSecret -ErrorAction Stop
    }
    else {
        # Fallback por si quieres correrlo tú de forma manual para probar, usa la identidad de tu máquina/usuario actual (Default Azure Credential)
        Write-Warning "No automation credentials provided. Attempting connection with interactive/cached credentials..."
        Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
    }
    
    Write-Host "Successfully connected to $SiteUrl" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to SharePoint: $($_.Exception.Message)"
    Exit
}

# Ejecutar lógica
$web = Get-PnPWeb
QueryUniquePermissions($web)

# Exportar resultados
Write-Host "Permission count: $($global:permissions.Count)"
$exportFilePath = Join-Path -Path $OutputPath -ChildPath $([string]::Concat($siteTitle,"-Permissions_",$dateTime,".csv"))

Write-Host "Export File Path is:" $exportFilePath
Write-Host "Number of lines exported is :" $global:permissions.Count

$global:permissions | Select-Object SiteUrl,SiteTitle,Type,RelativeUrl,ListTitle,MemberType,MemberName,MemberLoginName,ParentGroup,Roles | Export-CSV -Path $exportFilePath -NoTypeInformation -Encoding UTF8
