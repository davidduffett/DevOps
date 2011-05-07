function Create-User ([string]$accountName, [string]$password, [string]$hostName = ".", [string]$groups = "IIS_WPG") {
	if (!$accountName -or !$password) {
		$(Throw 'A value for $accountName and $password is required.')
	}
	
	$computer = [ADSI] "WinNT://$hostName"
	$user = $null
	foreach ($account in $computer.psbase.children) {
		if ($account.Name -eq $accountName) {
			Write-Host "User account $accountName already exists on $hostName"
			$user = $account
		}
	}
	
	if ($user -eq $null) {
		Write-Host "Creating user account $accountName on $hostName"
		$user = $computer.Create("User", "$accountName")
		$user.SetPassword($password)
		$user.SetInfo()
	}
	
	if ($groupList -ne $null) {
		$groupList = $groups.split(",")
		foreach ($group in $groupList) {
			Write-Host "Ensuring $accountName is a member of $group on $hostName"
			$alreadyAMember = $false
			$groupObj = [ADSI] "WinNT://$hostName/$group"
			$members = @($groupObj.psbase.Invoke("Members"))
			foreach ($member in $members) {
				if ($member.GetType().InvokeMember('Name','GetProperty',$null,$member,$null) -eq $accountName) {
					$alreadyAMember = $true
					break
				}
			}
			if (!$alreadyAMember) {
				$groupObj.psbase.Invoke("Add", $user.psbase.Path)
			}
		}
	}
}
