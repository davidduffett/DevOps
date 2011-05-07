. (Join-Path -Path (Split-Path -parent $MyInvocation.MyCommand.Definition) -ChildPath "7-zip.ps1")

$PACKETPRIVACY = 6

function Get-Website ([string]$computerName, [string]$websiteName) {
	return Get-WmiObject -Class IIsWebServerSetting -Namespace "root\microsoftiisv2" -ComputerName $computerName -filter "ServerComment = '$websiteName'" -Authentication $PACKETPRIVACY
}

function Create-AppPool ([string]$appPoolName, [string]$computerName, [string]$accountName, [string]$password) {
	if (!$appPoolName -or !$computerName -or !$accountName -or !$password) {
		$(Throw 'Values are required for $appPoolName, $computerName, $accountName and $password.')
	}

	$existingPool = Get-WmiObject -Class IIsApplicationPoolSetting -Namespace "root\microsoftiisv2" -ComputerName $computerName -filter "name = 'W3SVC/AppPools/$appPoolName'" -Authentication $PACKETPRIVACY
	if ($existingPool -ne $null) {
		Write-Host "Application pool $appPoolName already exists on $computerName"
		return
	}
	
	Write-Host "Creating application pool $appPoolName on $computerName"
	$appPools = [ADSI]"IIS://$computerName/W3SVC/AppPools"
	$newPool = $appPools.Create("IIsApplicationPool", $appPoolName)
	$newPool.Put("AppPoolIdentityType", 3)
	$newPool.Put("WAMUsername", $accountName)
	$newPool.Put("WAMUserPass", $password)
	$newPool.SetInfo()
}

function Create-Website ([string]$name, [string]$computerName, [string]$ipAddress, [string]$localPath, [string]$appPoolName, $port = "80") {
	if (!$name -or !$ipAddress -or !$localPath -or !$appPoolName) {
		$(Throw 'Values are required for $name, $ipAddress, $localPath and $appPoolName.')
	}
	
	$webSite = Get-Website $computerName $name
	if ($webSite -ne $null) {
		Write-Host "Website $name already exists on $computerName"
	} else {
		Write-Host "Creating website $name on $computerName"
		$sites = New-Object System.DirectoryServices.DirectoryEntry("IIS://$computerName/W3SVC")
		$newBinding = "$ipAddress:$port:"
		$sites.CreateNewSite($name, @($newBinding), $localPath)
		$webSite = Get-Website $computerName $name
	}
	
	Write-Host "Updating website settings for $name on $computerName"
	$webSite.AppFriendlyName = $name
	$webSite.AccessScript = $true
	$webSite.AuthAnonymous = $true
	$webSite.AppPoolId = $appPoolName
	[void]$webSite.Put()
}

function Create-VirtualDir ([string]$computerName, [string]$websiteName, [string]$virtualDirectoryName, [string]$localPath) {
	$virtualDir = Get-WebsiteDir $computerName $websiteName $virtualDirectoryName
	if ($virtualDir -ne $null) {
		Write-Host "Virtual directory $virtualDirectoryName already exists under website $websiteName on $computerName"
		return
	}
	
	Write-Host "Creating virtual directory $virtualDirectoryName under website $websiteName on $computerName"
	$webSite = Get-Website $computerName $websiteName
	if ($webSite -eq $null) {
		$(Throw "Could not find or access website $websiteName on $computerName")
	}

	$site = New-Object System.DirectoryServices.DirectoryEntry("IIS://$computerName/" + $webSite.Name + "/root")
	$children = $site.psbase.children
	$virtualDir = $children.add($virtualDirectoryName, $site.psbase.SchemaClassName)
	$virtualDir.Path = $localPath
	$virtualDir.psbase.CommitChanges()
}

function Get-WebsiteDir ([string]$computerName, [string]$websiteName, $virtualDirectoryName = $null) {
	$webSite = Get-Website $computerName $websiteName
	if ($webSite -eq $null) {
		$(Throw "Could not find or access website $websiteName on $computerName")
	}
	
	$nameQuery = "Name = '" + $webSite.Name + "/root'"
	if ($virtualDirectoryName -ne $null -and $virtualDirectoryName -ne "") {
		$nameQuery = "Name = '" + $webSite.Name + "/root/" + $virtualDirectoryName + "'"
	}

	return Get-WMIObject -Class IIsWebVirtualDirSetting -Namespace "root\microsoftiisv2" -Filter $nameQuery -ComputerName $computerName -Authentication $PACKETPRIVACY
}

function Start-Website ([string]$computerName, [string]$websiteName) {
	$webSite = Get-Website $computerName $websiteName
	if ($webSite -eq $null) {
		$(Throw "Could not find or access website $websiteName on $computerName")
	}
	$webServer = Get-WMIObject -Class IIsWebServer -Namespace "root\microsoftiisv2" -ComputerName $computerName -Authentication $PACKETPRIVACY | Where-Object { $_.Name -eq $webSite.Name }
	if ($webServer -eq $null) {
		$(Throw "Could not find or access website $websiteName on $computerName")
	}
	$webServer.Start()
}

function Set-P3PHttpHeader ([string]$computerName, [string]$websiteName, [string]$p3pUrl) {
	$webSite = Get-Website $computerName $websiteName
	if ($webSite -eq $null) {
		$(Throw "Could not find or access website $websiteName on $computerName")
	}

    Write-Host "Setting P3P policy HTTP header for $websiteName on $computerName"
	$path = $website.name + "/root"
    $vdir = Get-WmiObject -Class IIsWebVirtualDirSetting -Namespace "root\microsoftiisv2" -ComputerName $computerName -filter "Name = '$path'" -Authentication $PACKETPRIVACY
    $headers = $vdir.HttpCustomHeaders
    $headers[0].Keyname = 'P3P: policyref="' + $p3pUrl + '", CP="IDC DSP COR CUR DEV PSA IVA IVD CONo HIS TELo OUR DEL UNRo BUS UNI"'
    $headers[0].Value = $null
    $vdir.HttpCustomHeaders = $headers
    [void]$vdir.Put()
}

function Set-404Page ([string]$computerName, [string]$websiteName, [string]$404Url) {
	$webSite = Get-Website $computerName $websiteName
	if ($webSite -eq $null) {
		$(Throw "Could not find or access website $websiteName on $computerName")
	}

	Write-Host "Setting custom 404 error page to $404Url for $websiteName on $computerName"
	$path = $website.name + "/root"
    $vdir = Get-WmiObject -Class IIsWebVirtualDirSetting -Namespace "root\microsoftiisv2" -ComputerName $computerName -filter "Name = '$path'" -Authentication $PACKETPRIVACY
	$httpErrors = $vdir.HttpErrors
	for ($i = 1; $i -le $httpErrors.Count; $i++) {
		if ($httpErrors[$i].HttpErrorCode -eq "404") {
			$httpErrors[$i].HandlerType = "URL"
			$httpErrors[$i].HandlerLocation = $404Url
		}
	}
	$vdir.HttpErrors = $httpErrors
	[void]$vdir.Put()
}

function Set-ServerBindings ([string]$computerName, [string]$websiteName, [string]$bindingList) {
	$webSite = Get-Website $computerName $websiteName
	if ($webSite -eq $null) {
		$(Throw "Could not find or access website $websiteName on $computerName")
	}

    $newBindings = @($bindingList.Split(","))
    $bindings = $webSite.ServerBindings
    for ($i = 0; $i -lt $newBindings.Length; $i++) {
        Write-Host "Setting server binding" $newBindings[$i] "for $websiteName on $computerName"
        if ($bindings.Length -lt $i - 1) {
            $bindings.Insert($newBindings[$i])
        } else {
            $parts = @($newBindings[$i].Split(":"))
            $bindings[$i].IP = $parts[0]
            $bindings[$i].Port = $parts[1]
            $bindings[$i].Hostname = $parts[2]
        }
    }
    $webSite.ServerBindings = $bindings
    [void]$webSite.Put()
}

function Set-SecureBindings ([string]$computerName, [string]$websiteName, [string]$ipAddress, $port = 443) {
	$webSite = Get-Website $computerName $websiteName
	if ($webSite -eq $null) {
		$(Throw "Could not find or access website $websiteName on $computerName")
	}

	Write-Host "Setting secure bindings to $ipAddress:$port for $websiteName on $computerName"
    $bindings = $website.SecureBindings
    $bindings[0].IP = $ipAddress
    $bindings[0].Port = $port
    $website.SecureBindings = $bindings
    [void]$website.Put()
}

function Import-SslCertificate ([string]$computerName, [string]$websiteName, [string]$pfxPath, [string]$pfxPassword) {
	$webSite = Get-Website $computerName $websiteName
	if ($webSite -eq $null) {
		$(Throw "Could not find or access website $websiteName on $computerName")
	}

	$certMgr = New-Object -ComObject IIS.CertObj
	$certMgr.ServerName = $computerName
	$certMgr.InstanceName = $webSite.Name
	if (!$certMgr.IsInstalled()) {
		Write-Host "Importing SSL certificate $pfxPath"
		$certMgr.Import($pfxPath, $pfxPassword, $true, $true)
	} else {
		Write-Host "SSL certificate is already installed for $websiteName on $computerName"
	}
}

function Deploy-Website ([string]$packageFile, [string]$computerName, [string]$websiteName, [string]$blueIISPath, [string]$greenIISPath, [string]$blueNetworkPath, [string]$greenNetworkPath, [string]$virtualDirName = $null) {
	if (!(Test-Path $packageFile)) {
		$(Throw "Package file [$packageFile] could not be found or accessed.")
	}

	# Default to deploying to blue slice
	$newIISPath = $blueIISPath
	$newNetworkPath = $blueNetworkPath

	# Retrieve website directory
	$webDir = Get-WebsiteDir $computerName $websiteName $virtualDirName

	# Are we already on the blue slice? If so, switch to green
	Write-Host "Current site path is" $webDir.Path
	if ($webDir.Path -eq $blueIISPath) {
		$newIISPath = $greenIISPath
		$newNetworkPath = $greenNetworkPath
	}
	
	# Create or clean folder
	if (!(Test-Path $newNetworkPath)) {
		Write-Host "Creating folder $newNetworkPath"
		New-Item $newNetworkPath -type directory
	} else {
		Write-Host "Cleaning $newNetworkPath"
		Remove-Item "$newNetworkPath\*" -recurse -force
	}
	
	# Unzip build file
	Write-Host "Updating $newNetworkPath to new version"
	Unzip-File $packageFile $newNetworkPath

	# Switch IIS path
	Write-Host "Switching eSpares path to $newIISPath"
	$webDir.Path = $newIISPath
	Set-WmiInstance -InputObject $webDir | Out-Null
}
