<#
.SYNOPSIS 
Устанавливает и обновляет параметры веб-приложений 1С:Предприятия.
	
.DESCRIPTION
Устанавливает и обновляет параметры веб-приложений 1С:Предприятия.

.PARAMETER Config
Задает расположение файла конфигурации для скрипта. По умолчанию используется файл EnterpriseConfig.xml расположенный в одной папке со скриптом.

.PARAMETER SetISAPIHandlers
Устанавливает ISAPI обработчики у веб-приложений

.PARAMETER SetISAPIHandler <application,...>
Устанавливает ISAPI обработчики у указанных веб-приложений

.PARAMETER SetApplicationPools
Устанавливает пул приложений для веб-приложений

.PARAMETER SetApplicationPool <application,...>
Устанавливает пул приложений у указанных веб-приложений

.EXAMPLE
Update-EnterpiseWebApplications

.EXAMPLE
Update-EnterpiseWebApplications -SetISAPIHandler Application
  		
.NOTES
Версия 1.0.1.1

Для работы требуется устнановленный пакет WebAdministration
https://technet.microsoft.com/en-us/library/ee790599.aspx

Автор: Гончарук Юрий <yukon39@gmail.com>
Источник: https://github.com/yukon39/EnterpriseConfig
#>

#region CmdletBinding

[CmdletBinding()]
param(
		[parameter(Mandatory=$false)]
		[string]$Config="",
		[switch]$SetISAPIHandlers,
		[parameter(Mandatory=$false)]
		[string[]]$SetISAPIHandler=@(),
		[switch]$SetApplicationPools,
		[parameter(Mandatory=$false)]
		[string[]]$SetApplicationPool=@()
)
#endregion

#region Init

If ($Config -eq "") {
	$Config = Join-Path -Path $PSScriptRoot -ChildPath "EnterpriseConfig.xml"
}

If (-not (Test-Path -Path $Config -PathType Leaf)) {
	Throw "Config file '{0}' not found" -f $Config
}

#endregion

#region CommonFunctions 

Function Read-ScriptConfiguration([string]$ConfigFile) {

    [Xml.XmlDocument](Get-Content -Path $ConfigFile) | 
        Select-Object -ExpandProperty Configuration 
}

Function Get-BinaryPath([string]$id) {

   $script:Configuration.Binaries.ChildNodes | 
		Where-Object { ($_.id -ieq  $id) } | 
		Select-Object -First 1 -ExpandProperty Path 
}
    
#endregion

#region Functions 

Function Enable-ISAPIExtension {
<#
.SYNOPSIS 
Добавляет и разрешает ISAPI-обработчик платформы на локальном IIS-сервере

.LINK
https://www.iis.net/configreference/system.webserver/security/isapicgirestriction
#>
	[CmdletBinding()]
	param(
		[parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[string]$path,
		[parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[string]$id
	)

	Begin {
        
         Set-Variable -Name isapiCgiRestrictionXPath -Value "/system.webServer/security/isapiCgiRestriction" -Option Constant
        
        #Оба значения эквиваленты
        #Set-Variable -Name ServerPSPath -Value "MACHINE/WEBROOT/APPHOST" -Option Constant
        Set-Variable -Name ServerPSPath -Value "IIS:" -Option Constant
	}
	
	Process {

		$binaryPathName = Join-Path -Path $path -ChildPath "wsisapi.dll"
        $filter = "{0}/add[@path='{1}']" -f $isapiCgiRestrictionXPath,$binaryPathName
		
        if ((Get-WebConfiguration -PSPath $ServerPSPath -Filter $filter) -eq $null) {
		
        	Add-WebConfiguration -PSPath $ServerPSPath -Filter $isapiCgiRestrictionXPath -Value @{
                Path=$binaryPathName; 
                Description="1C Enterprise ISAPI extension"; 
                Allowed = $true }
			"ISAPI extension '{0}' added on server" -f $binaryPathName | Write-Host

        } elseif ((Get-WebConfigurationProperty -PSPath $ServerPSPath -Filter  $filter -Name allowed).Value -ne $true) {
            
            Set-WebConfigurationProperty -PSPath $ServerPSPath -Filter $filter -Name allowed –Value $true
			"ISAPI extension '{0}' allowed on server" -f $binaryPathName | Write-Host
        }
    }
}

Function Set-ISAPIExtension {
<#
.SYNOPSIS 
Добавляет и разрешает ISAPI-обработчик платформы для IIS-приложения
#>
	[CmdletBinding()]
	param(
		[parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[string]$id,
		[parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[string]$binary,
		[parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string]$Site="Default Web Site"
	)

	Process {

        $binaryPathName = Get-BinaryPath -id $binary | Join-Path -ChildPath "wsisapi.dll"  		

		[string]$ApplicationPSPath = "IIS:\Sites\{0}\{1}" -f $Site,$id

        Get-WebHandler -PSPath $ApplicationPSPath | 
            Where-Object { ($_.Modules -ieq "IsapiModule") -and ($_.scriptProcessor -ne $binaryPathName) } | 
            ForEach-Object { 
                Remove-WebHandler -Name $_.Name -PSPath $ApplicationPSPath 
                "ISAPI extension '{0}' ({1}) removed" -f $_.Name, $_.scriptProcessor | Write-Host
            }

       If ((Get-WebHandler -PSPath $ApplicationPSPath | 
               Where-Object { ($_.Modules -ieq "IsapiModule") -and ($_.scriptProcessor -ieq $binaryPathName) } |
               Select-Object -First 1) -eq $null) {
            New-WebHandler -PSPath $ApplicationPSPath -Modules "IsapiModule" -ScriptProcessor $binaryPathName -Name "1C Enterprise ISAPI extension" -Path "*" -Verb "*"
            "ISAPI extension '{0}' added to application {1}" -f $binaryPathName, $id | Write-Host
        }
	}
}

Function Set-ApplicationPool {
<#
.SYNOPSIS 
Устанавливает пул приложений для IIS-приложения
#>
	param(
		[parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[string]$id,
		[parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string]$ApplicationPool="DefaultAppPool",
		[parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
		[string]$Site="Default Web Site"
	)
	
    Process {
        
        [string]$ApplicationPSPath = "IIS:\Sites\{0}\{1}" -f $Site,$id
        
        If((Get-ItemProperty -Path $ApplicationPSPath -Name "applicationPool").Value -ne $ApplicationPool) {
            Set-ItemProperty -Path $ApplicationPSPath -Name "applicationPool" -Value $ApplicationPool
            "Application '{0}' moved to pool '{1}'" -f $id, $ApplicationPool | Write-Host
        }
    }
}

#endregion

Import-Module WebAdministration

[boolean]$IsSwitches = $SetISAPIHandlers -or $SetApplicationPools
[boolean]$IsItems = ($SetISAPIHandler.Count -ne 0) -or ($SetApplicationPool.Count -ne 0)
 
$Configuration = Read-ScriptConfiguration -ConfigFile $Config

If (-not $IsSwitches -or $IsItems -or $SetApplicationPools) {
    $Configuration.WebApplications.ChildNodes | 
    	Where-Object { -not $IsItems -or ($SetApplicationPool -icontains $_.id) } | 
    	Set-ApplicationPool
} 

If (-not $IsSwitches -or $IsItems -or $SetISAPIHandlers) {
 
    $Configuration.Binaries.ChildNodes | Enable-ISAPIExtension

    $Configuration.WebApplications.ChildNodes | 
        Where-Object { -not $IsItems -or ($SetISAPIHandler -icontains $_.id) } | 
	    Set-ISAPIExtension
}
