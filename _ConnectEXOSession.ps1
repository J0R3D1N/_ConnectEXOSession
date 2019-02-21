Function Global:_ConnectEXOSession {
	[CmdletBinding()]
	Param(
		[String]$Identity = $(
			$PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.Management.Automation.PSArgumentNullException]::New("Identity"),
                    "RequiredParameterNotDefined",
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    "Identity"
                )
            )
		),
		[ValidateSet("Module","Session")]
		[String]$ImportMethod,
		[Parameter(ParameterSetName="Hidden")]
		[Switch]$Unattended,
		[Parameter(ParameterSetName="Hidden")]
		[String]$Path,
		[Switch]$Clean
	)
    BEGIN {
        TRY {
            #Clean up any existing errors
            $Error.Clear()
			$CurrentSessions = Get-PSSession | ? {$_.ComputerName -eq "outlook.office365.com"}
            If ($EXOSessionStartTime) {$SessionDuration = ((Get-Date) - $EXOSessionStartTime)}
			Write-Verbose ("[{0}] Checking for Active Exchange Online Sessions" -f (Get-Date))
			If ($Clean -and $CurrentSessions.Count -eq 0) {
				Write-Warning "No PowerShell Sessions to Clean up"
				$ValidConnection = $false
			}
			ElseIf ($CurrentSessions.Count -eq 1 -and ($SessionDuration.TotalMinutes -lt 15) -and $Clean -eq $false) {
				Write-Verbose ("[{0}] ACTIVE PowerShell Session connected to: {1} at {2} ({3:N0} Minutes)" -f (Get-Date),$CurrentSessions.ComputerName,$EXOSessionStartTime,$SessionDuration.TotalMinutes)
				$ValidConnection = $true
			}
			ElseIf ($Clean -or ($SessionDuration.TotalMinutes -gt 15 -and $CurrentSessions.Count -eq 1) -or ($CurrentSessions.Count -gt 1)) {
				Write-Warning ("Clean Switch: {0} | Session Duration: {1:N0} Minutes | Active Sessions: {2}" -f $Clean,$SessionDuration.TotalMinutes,$CurrentSessions.Count)
				If ($CurrentSessions.Count -eq 0) {$ValidConnection = $false}
				Else {
					Get-Module | ? {$_.description -like "*outlook.office365.com*"} | Remove-Module -Force -Verbose:$false
					Get-PSSession -Verbose:$false | ? {$_.ComputerName -eq "outlook.office365.com"} | Remove-PSSession -Confirm:$false -Verbose:$false
					[System.GC]::Collect()
					for ($i=0;$i -le 5;$i++) {
						Write-Progress -Activity "PowerShell Session(s) Clean Up Timer" -Status "Waiting 5 seconds..." -SecondsRemaining (5 - $i)
						Start-Sleep -Milliseconds 999
					}
					Write-Progress -Activity "PowerShell Session(s) Clean Up Timer" -Status "CLEAN UP COMPLETE" -Completed
					$ValidConnection = $false
				}
			}
        }
        CATCH {$PSCmdlet.ThrowTerminatingError($PSitem)}
    }
    PROCESS {
        TRY {
			If ($ValidConnection) {Return}
			If ($Unattended) {
				If ([System.String]::IsNullOrEmpty($Path)) {
					$PSCmdlet.ThrowTerminatingError(
						[System.Management.Automation.ErrorRecord]::new(
							[System.Exception]::New("The Path parameter cannot be null or blank, you must provide the file system path to your secure credential files"),
							"RequiredParameterNotDefined",
							[System.Management.Automation.ErrorCategory]::InvalidArgument,
							"Path"
						)
					)
				}
				Write-Verbose ("[{0}] Entered Unattended Mode using Identity: {1}" -f (Get-Date),$Identity)
				$SessionName = $Identity.Split("@")[0]
				If ($Path.Substring($Path.Length - 1) -eq "\") {$Path = $Path.TrimEnd("\")}
				If (Test-Path -Path "$Path\*" -Include *.txt) {
					$PasswordFile = "$Path\$SessionName.txt"
					If (Test-Path $PasswordFile) {
						Write-Verbose ("[{0}] Building Exchange Online Global Credential Object" -f (Get-Date))
						$Password = Get-Content $PasswordFile
						$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
						$Global:EXOCredential = New-Object System.Management.Automation.PSCredential ($Identity, $SecurePassword)
					}
					Else {
						$PSCmdlet.ThrowTerminatingError(
							[System.Management.Automation.ErrorRecord]::new(
								[System.Exception]::New("Credential file for $Identity cannot be found"),
								"CredentialFileNotFound",
								[System.Management.Automation.ErrorCategory]::ObjectNotFound,
								$PasswordFile
							)
						) 
					}
				}
				Else {
					$PSCmdlet.ThrowTerminatingError(
						[System.Management.Automation.ErrorRecord]::new(
							[System.Exception]::New("The Path provided ($Path) is not valid or does not contain any secure credential files"),
							"InvalidPath",
							[System.Management.Automation.ErrorCategory]::ObjectNotFound,
							$Path
						)
					)
				}
			}
			Else {
				$IdentityRegex = "^[\w-\.]+@[\w- ]+\.+[\w-]{2,4}?$"
				$TenantRegex = "^[\w-\.]+@[\w- ]+\.+?onmicrosoft.com$"
				If ($Identity -match $IdentityRegex -or $Identity -match $TenantRegex) {
					Write-Verbose ("[{0}] Building Exchange Online Global Credential Object" -f (Get-Date))
					$SessionName = $Identity.Split("@")[0]
					$Global:EXOCredential = Get-Credential -UserName $Identity -Message "Exchange Online Remote PowerShell Authentication ($Identity)"
				}
				Else {
					$PSCmdlet.ThrowTerminatingError(
						[System.Management.Automation.ErrorRecord]::new(
							[System.FormatException]::New("The Identity provided is not a valid E-mail Address or User Principal Name"),
							"IdentityNotValid",
							[System.Management.Automation.ErrorCategory]::InvalidData,
							$Identity
						)
					)
				}
			}
			
			Write-Verbose ("[{0}] Creating Exchange Online PowerShell Session (EXO-{1})" -f (Get-Date),$SessionName)
			$EXOSession = New-PSSession -Name "EXO-$SessionName" -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -AllowRedirection -Authentication Basic -Credential $Global:EXOCredential -ErrorAction SilentlyContinue
			If ($Error.Count -gt 0) {
				Write-Debug "Error Found during New-PSSession"
				Write-Warning "Error Found during New-PSSession"
				Write-Warning ("[{0}]: {1}" -f $Error.Exception.GetType().FullName,$Error.Exception.Message)
				$ErrorCount++

				If ($ErrorCount -ge 2) {
					$PSCmdlet.ThrowTerminatingError(
						[System.Management.Automation.ErrorRecord]::new(
							[System.Management.Automation.Remoting.PSRemotingTransportException]::New("Unable to create clean PowerShell Session to Exchange Online"),
							"SessionCreationFailed",
							[System.Management.Automation.ErrorCategory]::ConnectionError,
							$ErrorCount
						)
					)
					Exit
				}
				for ($i=0;$i -le 5;$i++) {
					Write-Progress -Activity "PowerShell Session(s) Stall Timer" -Status "Waiting 5 seconds..." -SecondsRemaining (5 - $i)
					Start-Sleep -Milliseconds 999
				}
				Write-Progress -Activity "PowerShell Session(s) Stall Timer" -Status "FINISHED" -Completed
				Write-Verbose ("PowerShell Session Error - ATTEMPT {0} of 2" -f $ErrorCount)
				Write-Debug ("ReLaunch {0}" -f $PSCmdlet.MyInvocation.Line)
				_ConnectEXOSession @PSBoundParameters
			}
			Else {
				If ($ImportMethod -eq "Module") {
					Write-Debug "Import Method MODULE"
					Write-Verbose ("[{0}] PowerShell Session Import Method: {1}" -f (Get-Date),$ImportMethod)
					$ErrorCount = 0
					If (Test-Path -Path "C:\_SessionExport\EXO-$SessionName") {
						Write-Verbose ("[{0}] PowerShell Session Export Path Found" -f (Get-Date))
						$EXOModule = Get-ChildItem -Path "C:\_SessionExport\EXO-$SessionName" -Filter *.psm1
						$EXOManifest = Get-ChildItem -Path "C:\_SessionExport\EXO-$SessionName" -Filter *.psd1
					}
					Else {
						Write-Verbose ("[{0}] PowerShell Session Export Path Not Found - Creating folder(s) - C:\_SessionExport" -f (Get-Date))
						New-Item -ItemType Directory -Path "C:\_SessionExport" -Force | Out-Null
					}

					If ($EXOModule -and $EXOModule.LastWriteTime -gt (Get-Date).AddDays(-30)) {
						Write-Verbose ("[{0}] Exported Session Module Found and is Newer than 30 days" -f (Get-Date))
						$ModuleContent = Get-Content $EXOModule.FullName
						$ManifestContent = Get-Content $EXOManifest.FullName
						If ($ModuleContent) {
							$i = 0
							$Index = $null
							$ModuleContent | % {
								$i++
								$CurrentLine = $_
								If ($CurrentLine -Match "-Credential") {$Index = $i}
							}
							If ($Index -eq $null) {
								$PSCmdlet.ThrowTerminatingError(
									[System.Management.Automation.ErrorRecord]::new(
										[System.NullReferenceException]::New("Unable to determine the index value of '-Credential' in $EXOModule"),
										"IndexIsNullOrEmpty",
										[System.Management.Automation.ErrorCategory]::InvalidResult,
										"Index"
									)
								)
							}
							Else {
								$ModuleContent[$Index-1] = "                    -Credential `$Global:EXOCredential ``"
								$ModuleContent | Set-Content $EXOModule.FullName -Force
							}
						}
						If ($ManifestContent) {
							$i = 0
							$Index = $null
							$ManifestContent | % {
								$i++
								$CurrentLine = $_
								If ($CurrentLine -Match "Description") {$Index = $i}
							}
							If ($Index -eq $null) {
								$PSCmdlet.ThrowTerminatingError(
									[System.Management.Automation.ErrorRecord]::new(
										[System.NullReferenceException]::New("Unable to determine the index value of 'Description' in $EXOManifest"),
										"IndexIsNullOrEmpty",
										[System.Management.Automation.ErrorCategory]::InvalidResult,
										"Index"
									)
								)
							}
							Else {
								$ManifestContent[$Index-1] = "    Description = '[MODULE] Implicit remoting for https://outlook.office365.com/powershell-liveid'"
								$ManifestContent | Set-Content $EXOManifest.FullName -Force
							}
						}
					}
					Else {
						Write-Verbose ("[{0}] Exported Session Not Found - running Export-PSSession to C:\_SessionExport\EXO-{1}" -f (Get-Date),$SessionName)
						Export-PSSession -Session $EXOSession -OutputModule "C:\_SessionExport\EXO-$SessionName" -Encoding ASCII -AllowClobber -Force | Out-Null
						$EXOModule = Get-ChildItem -Path "C:\_SessionExport\EXO-$SessionName" -Filter *.psm1
						$EXOManifest = Get-ChildItem -Path "C:\_SessionExport\EXO-$SessionName" -Filter *.psd1
						$ModuleContent = Get-Content $EXOModule.FullName
						$ManifestContent = Get-Content $EXOManifest.FullName
						
						If ($ModuleContent) {
							$i = 0
							$Index = $null
							$ModuleContent | % {
								$i++
								$CurrentLine = $_
								If ($CurrentLine -Match "-Credential") {$Index = $i}
							}
							If ($Index -eq $null) {
								$PSCmdlet.ThrowTerminatingError(
									[System.Management.Automation.ErrorRecord]::new(
										[System.NullReferenceException]::New("Unable to determine the index value of '-Credential' in $EXOModule"),
										"IndexIsNullOrEmpty",
										[System.Management.Automation.ErrorCategory]::InvalidResult,
										"Index"
									)
								)
							}
							Else {
								$ModuleContent[$Index-1] = "                    -Credential `$Global:EXOCredential ``"
								$ModuleContent | Set-Content $EXOModule.FullName -Force
							}
						}

						If ($ManifestContent) {
							$i = 0
							$Index = $null
							$ManifestContent | % {
								$i++
								$CurrentLine = $_
								If ($CurrentLine -Match "Description") {$Index = $i}
							}
							If ($Index -eq $null) {
								$PSCmdlet.ThrowTerminatingError(
									[System.Management.Automation.ErrorRecord]::new(
										[System.NullReferenceException]::New("Unable to determine the index value of 'Description' in $EXOManifest"),
										"IndexIsNullOrEmpty",
										[System.Management.Automation.ErrorCategory]::InvalidResult,
										"Index"
									)
								)
							}
							Else {
								$ManifestContent[$Index-1] = "    Description = '[MODULE] Implicit remoting for https://outlook.office365.com/powershell-liveid'"
								$ManifestContent | Set-Content $EXOManifest.FullName -Force
							}
						}
					}
					Write-Verbose ("[{0}] Importing Module from PowerShell Session Export ({1})" -f (Get-Date),$EXOModule.FullName)
					$EXOModuleInfo = Import-Module "C:\_SessionExport\EXO-$SessionName" -ArgumentList $EXOSession -Prefix EXO -DisableNameChecking -PassThru -Force -Verbose:$false
					$Global:EXOSessionStartTime = Get-Date
					If ((Get-Command -Module $EXOModuleInfo).Count -gt 1) {
						Write-Verbose ("[{0}] PowerShell Session connected to: {1}" -f $EXOSessionStartTime,$EXOSession.ComputerName)
						Write-Verbose ("[{0}] Found {1} Imported Commands from {2}" -f $EXOSessionStartTime,(Get-Command -Module $EXOModuleInfo).Count,$EXOModuleInfo.Name)
					}
					Else {
						$PSCmdlet.ThrowTerminatingError(
							[System.Management.Automation.ErrorRecord]::new(
								[System.NullReferenceException]::New("The module import of $EXOModule was not successful, check the module to verify it exists and has been properly modified to work with the `$Global:EXOCredential object"),
								"PowerShellModuleNotFound",
								[System.Management.Automation.ErrorCategory]::InvalidResult,
								$EXOModule
							)
						)
					}
				}
				ElseIf ($ImportMethod -eq "Session") {
					Write-Verbose ("[{0}] PowerShell Session Import Method: {1}" -f (Get-Date),$ImportMethod)
					$ErrorCount = 0
					$EXOSessionInfo = Import-PSSession $EXOSession -Prefix EXO -DisableNameChecking -AllowClobber -Verbose:$false
					$Global:EXOSessionStartTime = Get-Date
					If ($EXOSessionInfo.ExportedCommands.Count -gt 1) {
						Write-Verbose ("[{0}] PowerShell Session connected to: {1}" -f $EXOSessionStartTime,$EXOSession.ComputerName)
						Write-Verbose ("[{0}] Found {1} Imported Commands from {2}" -f $EXOSessionStartTime,$EXOSessionInfo.ExportedCommands.Count,$EXOSessionInfo.Name)
					}
					Else {
						$PSCmdlet.ThrowTerminatingError(
							[System.Management.Automation.ErrorRecord]::new(
								[System.NullReferenceException]::New("The PowerShell session import was not successful"),
								"PowerShellSessionNotFound",
								[System.Management.Automation.ErrorCategory]::InvalidResult,
								$EXOSession
							)
						)
					}
				}
				Else {
					$PSCmdlet.ThrowTerminatingError(
						[System.Management.Automation.ErrorRecord]::new(
							[System.Management.Automation.RuntimeException]::New("The ImportMethod parameter cannot be null or blank, you must choose either 'Module' or 'Session'"),
							"RequiredParameterNotDefined",
							[System.Management.Automation.ErrorCategory]::InvalidArgument,
							"ImportMethod"
						)
					)
				}
			}
        }
        CATCH {$PSCmdlet.ThrowTerminatingError($PSitem)}
    }
}