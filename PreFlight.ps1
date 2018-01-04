[string] $version = "1.4"

<#

.DESCRIPTION



###############Disclaimer#####################################################

This is a sample script. The sample scripts are not supported under any 

Microsoft standard support program or service. The sample scripts are 

provided AS IS without warranty of any kind. Microsoft further disclaims all 

implied warranties including, without limitation, any implied warranties of 

merchantability or of fitness for a particular purpose. The entire risk 

arising out of the use or performance of the sample scripts and documentation 

remains with you. In no event shall Microsoft, its authors, or anyone else 

involved in the creation, production, or delivery of the scripts be liable 

for any damages whatsoever (including, without limitation, damages for loss 

of business profits, business interruption, loss of business information, or 

other pecuniary loss) arising out of the use of or inability to use the 

sample scripts or documentation, even if Microsoft has been advised of the 

possibility of such damages.

###############Disclaimer#####################################################


#>

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Text")

#region folders preparation
    mkdir ([System.IO.Path]::GetFullPath("$($Script:MyInvocation.MyCommand.Path)\..\Reports")) -ErrorAction Ignore | Out-Null
    mkdir ([System.IO.Path]::GetFullPath("$($Script:MyInvocation.MyCommand.Path)\..\Scripts")) -ErrorAction Ignore | Out-Null
#endregion

#region custom types
    Add-Type -Language VisualBasic -TypeDefinition @"
    Public Class PreFlightItem
        Public primarySMTPAddress As String
        Public status As String
        Public errorMessage As String
    End Class
"@
#endregion

#region global variables
    [Boolean] $Global:configurationFinished = $False
    [PSCredential] $Global:localCred = New-Object System.Management.Automation.PSCredential ("dummy", (ConvertTo-SecureString "dummy" -AsPlainText -Force))
    [PSCredential] $Global:cloudCred = New-Object System.Management.Automation.PSCredential ("dummy", (ConvertTo-SecureString "dummy" -AsPlainText -Force))
    [bool] $Global:loadCloudMailboxes = $False
    [bool] $Global:isConnected = $False
    [string] $Global:localExchange = ""
    [string] $Global:serviceDomain = ""
    [string] $Global:scheduleStartDateTime = ""
    [string] $Global:scheduleCompleteDateTime = ""
    [string[]] $Global:endPointList = @()
    [string] $Global:migrationEndpoint = ""
    [System.Drawing.Size] $drawingSize = New-Object -TypeName System.Drawing.Size
    [System.Drawing.Point] $drawingPoint = New-Object -TypeName System.Drawing.Point
    [System.Windows.Forms.FormWindowState] $windowState = New-Object System.Windows.Forms.FormWindowState
    [int] $Global:migrationStrategy = 0
#endregion

#region fnConnect
    Function fnConnect {
        [bool] $continue = $True

        $progressBar.Value = 10
        $progressBar.Visible = $True
        $statusLabel.Text = "Connecting to Exchange Online..."
        $cloudSession = Get-PSSession | Where-Object {($_.ComputerName -eq "ps.outlook.com") -and ($_.ConfigurationName -eq "Microsoft.Exchange")}
        if ($CloudSession) {
		    Write-Host "Already connected to Exchange Online" -ForegroundColor Blue
            $Global:isConnected = [bool] ($CloudSession)
        }	
		else {
            if ($Global:cloudCred.UserName -eq "dummy") {
                $result = fnConfigure
                $continue = ($result -eq [System.Windows.Forms.DialogResult]::OK)
            }
            if ($continue) {
			    $cloudSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell" -AllowRedirection -Credential $Global:cloudCred -Authentication Basic
			    Import-PSSession $cloudSession -CommandName Get-Mailbox, Get-MailUser, New-MoveRequest, Get-AcceptedDomain, New-MigrationBatch, Get-MigrationEndpoint

                $cloudSession = Get-PSSession | Where-Object {($_.ComputerName -eq "ps.outlook.com") -and ($_.ConfigurationName -eq "Microsoft.Exchange")}
                $Global:isConnected = [bool] ($CloudSession)

                fnLoad
            }
		}
        $progressBar.Visible = $False
        $statusLabel.Text = ""
        return $continue
    }
#endregion

#region fnLoad
    Function fnLoad {
        $onPremisesTreeView.Nodes.Clear()
        $onlineTreeView.Nodes.Clear()
        $Global:endPointList.Clear()

        if (-not $Global:isConnected) {fnConnect}
        else {
            $progressBar.Value = 25
            $progressBar.Visible = $True
            $statusLabel.Text = "Loading list of users available for migration..."
            Get-MailUser -ResultSize Unlimited | Where-Object {$_.ExchangeGuid -ne "00000000-0000-0000-0000-000000000000"} | ForEach-Object {
                $onPremisesTreeView.Nodes.Add($_.PrimarySmtpAddress.ToString(), $_.Name)
            }
            $progressBar.Value = 75
            $statusLabel.Text = "Discovering service domain..."
            $Global:serviceDomain = (Get-AcceptedDomain | Where-Object {$_.DomainName -like "*.mail.onmicrosoft.com"}).DomainName.ToString()
            $progressBar.Value = 90
            $statusLabel.Text = "Loading list of migration endpoints..."
            Get-MigrationEndpoint | Where-Object {$_.EndpointType -eq "ExchangeRemoteMove"} | ForEach-Object {
                $Global:endPointList += $_.Identity
            }
        }
        $progressBar.Visible = $False
        $statusLabel.Text = ""
    }
#endregion

#region fnDisconnect
    Function fnDisconnect {
        $cloudSession = Get-PSSession | Where-Object {($_.ComputerName -eq "ps.outlook.com") -and ($_.ConfigurationName -eq "Microsoft.Exchange")}
        if ($cloudSession) {
            Remove-PSSession $cloudSession
        }
		else {
			Write-Host "There is no connection to Exchange Online" -ForegroundColor Blue
		}
    }
#endregion

#region fnConfigure
    Function fnConfigure {
        [System.Windows.Forms.Form] $frmConfig = New-Object -TypeName System.Windows.Forms.Form
        [System.Windows.Forms.GroupBox] $grpLocal = New-Object -TypeName System.Windows.Forms.GroupBox
        [System.Windows.Forms.GroupBox] $grpOnline = New-Object -TypeName System.Windows.Forms.GroupBox
        [System.Windows.Forms.TextBox] $txtLocalUser = New-Object -TypeName System.Windows.Forms.TextBox
        [System.Windows.Forms.TextBox] $txtLocalPassword = New-Object -TypeName System.Windows.Forms.TextBox
        [System.Windows.Forms.TextBox] $txtLocalExchange = New-Object -TypeName System.Windows.Forms.TextBox
        [System.Windows.Forms.TextBox] $txtCloudUser = New-Object -TypeName System.Windows.Forms.TextBox
        [System.Windows.Forms.TextBox] $txtCloudPassword = New-Object -TypeName System.Windows.Forms.TextBox
        [System.Windows.Forms.Label] $lblLocalUser = New-Object -TypeName System.Windows.Forms.Label
        [System.Windows.Forms.Label] $lblLocalPassword = New-Object -TypeName System.Windows.Forms.Label
        [System.Windows.Forms.Label] $lblLocalExchange = New-Object -TypeName System.Windows.Forms.Label
        [System.Windows.Forms.Label] $lblCloudUser = New-Object -TypeName System.Windows.Forms.Label
        [System.Windows.Forms.Label] $lblCloudPassword = New-Object -TypeName System.Windows.Forms.Label
        [System.Windows.Forms.Button] $btnCancel = New-Object -TypeName System.Windows.Forms.Button
        [System.Windows.Forms.Button] $btnOk = New-Object -TypeName System.Windows.Forms.Button

        #region txtLocalUser
            $drawingPoint.X = 151
            $drawingPoint.Y = 20
            $drawingSize.Height = 20
            $drawingSize.Width = 179
            $txtLocalUser.Location = $drawingPoint
            $txtLocalUser.Size = $drawingSize
        #endregion

        #region txtLocalPassword
            $drawingPoint.X = 151
            $drawingPoint.Y = 52
            $drawingSize.Height = 20
            $drawingSize.Width = 179
            $txtLocalPassword.Location = $drawingPoint
            $txtLocalPassword.Size = $drawingSize
            $txtLocalPassword.UseSystemPasswordChar = $True
        #endregion

        #region txtLocalExchange
            $drawingPoint.X = 151
            $drawingPoint.Y = 84
            $drawingSize.Height = 20
            $drawingSize.Width = 179
            $txtLocalExchange.Location = $drawingPoint
            $txtLocalExchange.Size = $drawingSize
        #endregion

        #region txtCloudUser
            $drawingPoint.X = 151
            $drawingPoint.Y = 20
            $drawingSize.Height = 20
            $drawingSize.Width = 179
            $txtCloudUser.Location = $drawingPoint
            $txtCloudUser.Size = $drawingSize
        #endregion

        #region txtCloudPassword
            $drawingPoint.X = 151
            $drawingPoint.Y = 52
            $drawingSize.Height = 20
            $drawingSize.Width = 179
            $txtCloudPassword.Location = $drawingPoint
            $txtCloudPassword.Size = $drawingSize
            $txtCloudPassword.UseSystemPasswordChar = $True
        #endregion

        #region lblLocalUser
            $drawingPoint.X = 8
            $drawingPoint.Y = 23
            $drawingSize.Height = 13
            $drawingSize.Width = 137
            $lblLocalUser.AutoSize = $True
            $lblLocalUser.Location = $drawingPoint
            $lblLocalUser.Name = "lblLocalUser"
            $lblLocalUser.Size = $drawingSize
            $lblLocalUser.Text = "User name (DOMAIN\User)"
        #endregion

        #region lblLocalPassword
            $drawingPoint.X = 92
            $drawingPoint.Y = 55
            $drawingSize.Height = 13
            $drawingSize.Width = 53
            $lblLocalPassword.AutoSize = $True
            $lblLocalPassword.Location = $drawingPoint
            $lblLocalPassword.Name = "lblLocalPassword"
            $lblLocalPassword.Size = $drawingSize
            $lblLocalPassword.Text = "Password"
        #endregion

        #region lblLocalExchange
            $drawingPoint.X = 53
            $drawingPoint.Y = 87
            $drawingSize.Height = 13
            $drawingSize.Width = 92
            $lblLocalExchange.AutoSize = $True
            $lblLocalExchange.Location = $drawingPoint
            $lblLocalExchange.Name = "lblLocalExchange"
            $lblLocalExchange.Size = $drawingSize
            $lblLocalExchange.TabIndex = 0
            $lblLocalExchange.Text = "MRS proxy FQDN"
        #endregion

        #region lblCloudUser
            $drawingPoint.X = 8
            $drawingPoint.Y = 23
            $drawingSize.Height = 13
            $drawingSize.Width = 132
            $lblCloudUser.AutoSize = $True
            $lblCloudUser.Location = $drawingPoint
            $lblCloudUser.Name = "lblCloudUser"
            $lblCloudUser.Size = $drawingSize
            $lblCloudUser.TabIndex = 0
            $lblCloudUser.Text = "User name (user@domain)"
        #endregion

        #region lblCloudPassword
            $drawingPoint.X = 92
            $drawingPoint.Y = 55
            $drawingSize.Height = 13
            $drawingSize.Width = 53
            $lblCloudPassword.AutoSize = $True
            $lblCloudPassword.Location = $drawingPoint
            $lblCloudPassword.Name = "lblCloudPassword"
            $lblCloudPassword.Size = $drawingSize
            $lblCloudPassword.TabIndex = 0
            $lblCloudPassword.Text = "Password"
        #endregion

        #region btnOk
            $drawingPoint.X = 533
            $drawingPoint.Y = 133
            $drawingSize.Height = 23
            $drawingSize.Width = 75
            $btnOk.Location = $drawingPoint
            $btnOk.Name = "btnOk"
            $btnOk.Size = $drawingSize
            $btnOk.Text = "Ok"
            $btnOk.UseVisualStyleBackColor = $True
            $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $btnOk.Add_Click({
                $Global:localCred = New-Object System.Management.Automation.PSCredential ($txtLocalUser.Text, (ConvertTo-SecureString $txtLocalPassword.Text -AsPlainText -Force))
                $Global:cloudCred = New-Object System.Management.Automation.PSCredential ($txtcloudUser.Text, (ConvertTo-SecureString $txtcloudPassword.Text -AsPlainText -Force))
                $Global:localExchange = $txtLocalExchange.Text
                $Global:configurationFinished = $True
                $frmConfig.Close()
            })
        #endregion

        #region btnCancel
            $drawingPoint.X = 615
            $drawingPoint.Y = 133
            $drawingSize.Height = 23
            $drawingSize.Width = 75
            $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $btnCancel.Location = $drawingPoint
            $btnCancel.Name = "btnCancel"
            $btnCancel.Size = $drawingSize
            $btnCancel.Text = "Cancel"
            $btnCancel.UseVisualStyleBackColor = $True
            $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        #endregion

        #region grpLocal
            $drawingPoint.X = 12
            $drawingPoint.Y = 12
            $drawingSize.Height = 115
            $drawingSize.Width = 336
            $grpLocal.Controls.Add($txtLocalUser)
            $grpLocal.Controls.Add($txtLocalPassword)
            $grpLocal.Controls.Add($txtLocalExchange)
            $grpLocal.Controls.Add($lblLocalUser)
            $grpLocal.Controls.Add($lblLocalPassword)
            $grpLocal.Controls.Add($lblLocalExchange)
            $grpLocal.Location = $drawingPoint
            $grpLocal.Name = "grpLocal"
            $grpLocal.Size = $drawingSize
            $grpLocal.TabStop = $False
            $grpLocal.Text = "Exchange on-premises"
        #endregion

        #region grpOnline
            $drawingPoint.X = 354
            $drawingPoint.Y = 12
            $drawingSize.Height = 115
            $drawingSize.Width = 336
            $grpOnline.Controls.Add($txtCloudUser)
            $grpOnline.Controls.Add($txtCloudPassword)
            $grpOnline.Controls.Add($lblCloudUser)
            $grpOnline.Controls.Add($lblCloudPassword)
            $grpOnline.Location = $drawingPoint
            $grpOnline.Name = "grpOnline"
            $grpOnline.Size = $drawingSize
            $grpOnline.TabStop = $False
            $grpOnline.Text = "Exchange Online"
        #endregion

        #region frmConfig
            $drawingSize.Height = 169
            $drawingSize.Width = 700
            $frmConfig.AcceptButton = $btnOk
            $frmConfig.CancelButton = $btnCancel
            $frmConfig.ClientSize = $drawingSize
            $frmConfig.ControlBox = $False
            $frmConfig.Controls.Add($grpLocal)
            $frmConfig.Controls.Add($grpOnline)
            $frmConfig.Controls.Add($btnOk)
            $frmConfig.Controls.Add($btnCancel)
            $frmConfig.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
            $frmConfig.MaximizeBox = $False
            $frmConfig.MinimizeBox = $False
            $frmConfig.Name = "frmConfig"
            $frmConfig.Text = "Configuration"
            $frmConfig.Add_Closed({$frmConfig = $null})
            $frmConfig.Add_Load({
                If ($Global:configurationFinished) {
                    $txtLocalUser.Text = $Global:localCred.UserName
                    $txtLocalPassword.Text = "***************"
                    $txtCloudUser.Text = $Global:cloudCred.UserName
                    $txtCloudPassword.Text = "***************"
                    $txtLocalExchange.Text = $Global:localExchange
                }
            })
        #endregion

        $grpLocal.ResumeLayout($False)
        $grpLocal.PerformLayout()
        $grpOnline.ResumeLayout($False)
        $grpOnline.PerformLayout()
        $frmConfig.ResumeLayout($False)
            
        $frmConfig.WindowState = $windowState
        #[void]
        return $frmConfig.ShowDialog()
    }
#endregion

#region fnAbout
    Function fnAbout {
        [System.Windows.Forms.Form] $frmAbout = New-Object -TypeName System.Windows.Forms.Form
        [System.Windows.Forms.Panel] $panelAbout = New-Object -TypeName System.Windows.Forms.Panel
        [System.Windows.Forms.Label] $lblTitle = New-Object -TypeName System.Windows.Forms.Label
        [System.Windows.Forms.Label] $lblName = New-Object -TypeName System.Windows.Forms.Label
        [System.Windows.Forms.Label] $lblVersion = New-Object -TypeName System.Windows.Forms.Label
        [System.Windows.Forms.Label] $lblDisclaimer = New-Object -TypeName System.Windows.Forms.Label
        [System.Windows.Forms.Button] $btnAboutOk = New-Object -TypeName System.Windows.Forms.Button
        [System.Windows.Forms.LinkLabel] $lnkBlog = New-Object -TypeName System.Windows.Forms.LinkLabel
        [System.Drawing.Font] $formFont = New-Object -TypeName System.Drawing.Font("Century Gothic",36,[System.Drawing.FontStyle]::Regular)

        #region panelAbout
            $drawingPoint.X = 0
            $drawingPoint.Y = 0
            $drawingSize.Height = 84
            $drawingSize.Width = 439
            $panelAbout.Location = $drawingPoint
            $panelAbout.Size = $drawingSize
            $panelAbout.BackColor = [System.Drawing.Color]::White
            $panelAbout.Controls.Add($lblTitle)
            $panelAbout.Name = "panelAbout"
        #endregion

        #region lblTitle
            $drawingPoint.X = 9
            $drawingPoint.Y = 9
            $drawingSize.Height = 58
            $drawingSize.Width = 326
            $lblTitle.AutoSize = $True
            $lblTitle.Location = $drawingPoint
            $lblTitle.Name = "lblTitle"
            $lblTitle.Size = $drawingSize
            $lblTitle.Text = "Pre-flight tool"
            $lblTitle.Font = $formFont
        #endregion

        #region lblName
            $drawingPoint.X = 16
            $drawingPoint.Y = 101
            $drawingSize.Height = 13
            $drawingSize.Width = 106
            $lblName.AutoSize = $True
            $lblName.Location = $drawingPoint
            $lblName.Name = "lblName"
            $lblName.Text = "Pre-flight tool"
        #endregion

        #region lblVersion
            $drawingPoint.X = 16
            $drawingPoint.Y = 119
            $drawingSize.Height = 13
            $drawingSize.Width = 60
            $lblVersion.AutoSize = $True
            $lblVersion.Location = $drawingPoint
            $lblVersion.Name = "lblVersion"
            $lblVersion.Text = "Version $version"
        #endregion

        #region lblDisclaimer
            $drawingPoint.X = 18
            $drawingPoint.Y = 154
            $drawingSize.Height = 163
            $drawingSize.Width = 397
            $lblDisclaimer.AutoSize = $False
            $lblDisclaimer.Location = $drawingPoint
            $lblDisclaimer.Size = $drawingSize
            $lblDisclaimer.Name = "lblDisclaimer"
            $lblDisclaimer.Text = "Version $version"
            $lblDisclaimer.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $lblDisclaimer.Text = "The Pre-flight tool is a sample script. The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages."
        #endregion

        #region lnkBlog
            $drawingPoint.X = 16
            $drawingPoint.Y = 138
            $drawingSize.Height = 13
            $drawingSize.Width = 102
            $lnkBlog.AutoSize = $True
            $lnkBlog.Location = $drawingPoint
            $lnkBlog.Name = "lnkBlog"
            $lnkBlog.Text = "FastTrack Tips Blog"
            $lnkBlog.Add_Click({Start-Process "http://aka.ms/ftctips"})
        #endregion

        #region btnAboutOk
            $drawingPoint.X = 340
            $drawingPoint.Y = 320
            $drawingSize.Height = 23
            $drawingSize.Width = 75
            $btnAboutOk.Location = $drawingPoint
            $btnAboutOk.Name = "btnAboutOk"
            $btnAboutOk.Size = $drawingSize
            $btnAboutOk.Text = "&Ok"
            $btnAboutOk.UseVisualStyleBackColor = $True
            $btnAboutOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $btnAboutOk.Add_Click({$frmAbout.Close()})
        #endregion

        #region frmAbout
            $drawingSize.Height = 356
            $drawingSize.Width = 440
            $frmAbout.AcceptButton = $btnAboutOk
            $frmAbout.ClientSize = $drawingSize
            $frmAbout.ControlBox = $False
            $frmAbout.Controls.Add($lblDisclaimer)
            $frmAbout.Controls.Add($lnkBlog)
            $frmAbout.Controls.Add($lblVersion)
            $frmAbout.Controls.Add($lblName)
            $frmAbout.Controls.Add($panelAbout)
            $frmAbout.Controls.Add($btnAboutOk)
            $frmAbout.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
            $frmAbout.MaximizeBox = $False
            $frmAbout.MinimizeBox = $False
            $frmAbout.Name = "frmAbout"
            $frmAbout.Text = "About Pre-flight"
            $frmAbout.ShowIcon = $False
            $frmAbout.ShowInTaskbar = $False
            $frmAbout.Add_Closed({$frmAbout = $null})
        #endregion

        $frmAbout.ResumeLayout($False)
        $frmAbout.ResumeLayout()
        $frmAbout.WindowState = $windowState
        [void] $frmAbout.ShowDialog()
    }
#endregion

#region fnRunPreFlight
    Function fnRunPreFlight {
        [PreFlightItem[]] $preFlightReport = @()
        [Int] $totalMailboxes = 0
        [Int] $currentMailbox = 0

        if (-not $Global:isConnected) {fnConnect}
        $totalMailboxes = $onlineTreeView.Nodes.Count

        if ($totalMailboxes -gt 0) {
            $progressBar.Value = 0
            $progressBar.Visible = $True
            $onlineTreeView.Nodes | ForEach-Object {
                $reportItem = New-Object -TypeName PreFlightItem
                $reportItem.primarySMTPAddress = $_.Name
                $Error.Clear()
                Write-Progress -Activity "Running pre-flight" -Status "Checking $($_.Name) - $([math]::Round($progressBar.Value))% complete" -PercentComplete ($progressBar.Value)
                $statusLabel.Text = "Checking $($_.Name)"
                try {
                    New-MoveRequest -Remote -RemoteHostName $Global:localExchange -RemoteCredential $Global:localCred -Identity $_.Name -TargetDeliveryDomain $Global:serviceDomain -BatchName "PreFlight" -ErrorAction Stop -WhatIf
                    $Message= $Error[0].Exception.Message

                    if($Message -eq $null) {
                        $reportItem.status = "Pass"
                        $reportItem.errorMessage = ""
                    }
                    else {
                        $reportItem.status = "Fail"
                        $reportItem.errorMessage = $Error[0].Exception.Message
                    }
                } catch {
                    $reportItem.status = "Fail"
                    $reportItem.errorMessage = $Error[0].Exception.Message
                }
                $preFlightReport += $reportItem
                $currentMailbox++
                $progressBar.Value = (100*($currentMailbox))/$totalMailboxes
                Write-Progress -Activity "Running pre-flight" -Status "Checked $($_.Name) - $([math]::Round($progressBar.Value))% complete" -PercentComplete ($progressBar.Value)
                $statusLabel.Text = "Checked $($_.Name)"
            }
            Write-Progress -Activity "Running pre-flight" -Completed
            fnWriteReport -ReportData $preFlightReport
            $statusLabel.Text = ""
            $progressBar.Visible = $False
            $progressBar.Value = 0
        }
    }
#endregion

#region fnWriteReport
    Function fnWriteReport {
        Param ([PreFlightItem[]] $ReportData)

        [string] $reportFilePath = [System.IO.Path]::GetFullPath("$($Script:MyInvocation.MyCommand.Path)\..\Reports\$(Get-Date -Format "yyyymmdd-HHmmss").csv")

        "primarySMTPAddress,status,errorMessage" | Out-File -FilePath $reportfilePath -Encoding ascii -Force

        $ReportData | ForEach-Object {
            [string] $reportLine = "$($_.primarySMTPAddress),$($_.status),$($_.errorMessage)"
            $reportLine | Out-File -FilePath $reportFilePath -Encoding ascii -Append
        }
    }
#endregion

#region fnWriteScript
    Function fnWriteScript {
        [Int] $totalMailboxes = $onlineTreeView.Nodes.Count
        [string] $scriptFileName = "$(Get-Date -Format "yyyymmdd-HHmmss").ps1"
        [System.Windows.Forms.SaveFileDialog] $saveDialog = New-Object -TypeName System.Windows.Forms.SaveFileDialog

        if ($totalMailboxes -gt 0) {
            $saveDialog.InitialDirectory = [System.IO.Path]::GetFullPath("$($Script:MyInvocation.MyCommand.Path)\..\Scripts")
            $saveDialog.Filter = "Windows PowerShell Script (*.ps1)|*.ps1|All files (*.*)|*.*"
            $saveDialog.FileName = $scriptFileName
            if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                [string] $scriptFilePath = $saveDialog.FileName

                "$('$')localExchange = '$Global:localExchange'" | Out-File -FilePath $scriptFilePath -Encoding ascii -Force
                '$localCred = Get-Credential -Message "Enter your local credential"' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                '' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                '$cloudSession = Get-PSSession | Where-Object {($_.ComputerName -eq "ps.outlook.com") -and ($_.ConfigurationName -eq "Microsoft.Exchange")}' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                'if ($cloudSession) {' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                '    $disconnectAtTheEnd = $False' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
	    	    '    Write-Host "Already connected to Exchange Online" -ForegroundColor Blue' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                '}' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
    		    'else {' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                '    $disconnectAtTheEnd = $True' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                '    $cloudCred = Get-Credential -Message "Enter your cloud credential"' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
	    	    '    $cloudSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell" -AllowRedirection -Credential $cloudCred -Authentication Basic' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
		        '    Import-PSSession $cloudSession -CommandName New-MoveRequest' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                '}' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                '' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                $onlineTreeView.Nodes | ForEach-Object {
                    "New-MoveRequest -Remote -RemoteHostName $('$')localExchange -RemoteCredential $('$')localCred -Identity $($_.Name) -TargetDeliveryDomain $Global:serviceDomain" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                }
                '' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                'if ($disconnectAtTheEnd) {Remove-PSSession $cloudSession}' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
            }
        }
    }
#endregion

#region fnSchedule
    Function fnSchedule {
        [Int] $currentMailbox = 0
        [Int] $totalMailboxes = $onlineTreeView.Nodes.Count
        [string[]] $migrationList = @("EmailAddress")
        [string] $batchName = "$(Get-Date -Format "yyyyMMdd-HHmmss")"
        [string] $format = [System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.FullDateTimePattern
        [System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.DialogResult]::OK
        $Global:migrationStrategy = 0
        $Global:scheduleStartDateTime = ""
        $Global:scheduleCompleteDateTime = ""

        if ($totalMailboxes -gt 0) {
            if ($Global:endPointList.Count -gt 0) {
                #region form
                    [System.Windows.Forms.Form] $frmSchedule = New-Object -TypeName System.Windows.Forms.Form
                    [System.Windows.Forms.Button] $btnScheduleOk = New-Object -TypeName System.Windows.Forms.Button
                    [System.Windows.Forms.Button] $btnScheduleCancel = New-Object -TypeName System.Windows.Forms.Button
                    [System.Windows.Forms.DateTimePicker] $startSchedulePicker = New-Object -TypeName System.Windows.Forms.DateTimePicker
                    [System.Windows.Forms.DateTimePicker] $completeSchedulePicker = New-Object -TypeName System.Windows.Forms.DateTimePicker
                    [System.Windows.Forms.GroupBox] $grpStart = New-Object -TypeName System.Windows.Forms.GroupBox
                    [System.Windows.Forms.GroupBox] $grpComplete = New-Object -TypeName System.Windows.Forms.GroupBox
                    [System.Windows.Forms.RadioButton] $radioStartManual = New-Object -TypeName System.Windows.Forms.RadioButton
                    [System.Windows.Forms.RadioButton] $radioStartAutomatic = New-Object -TypeName System.Windows.Forms.RadioButton
                    [System.Windows.Forms.RadioButton] $radioStartSchedule = New-Object -TypeName System.Windows.Forms.RadioButton
                    [System.Windows.Forms.RadioButton] $radioCompleteManual = New-Object -TypeName System.Windows.Forms.RadioButton
                    [System.Windows.Forms.RadioButton] $radioCompleteAutomatic = New-Object -TypeName System.Windows.Forms.RadioButton
                    [System.Windows.Forms.RadioButton] $radioCompleteSchedule = New-Object -TypeName System.Windows.Forms.RadioButton
                    [System.Windows.Forms.Label] $lblSelectEndpoint = New-Object -TypeName System.Windows.Forms.Label
                    [System.Windows.Forms.ComboBox] $endpointBox = New-Object -TypeName System.Windows.Forms.ComboBox

                    $grpStart.SuspendLayout()
                    $grpComplete.SuspendLayout()
                    $frmSchedule.SuspendLayout()

                    #region btnScheduleOk
                        $drawingPoint.X = 538
                        $drawingPoint.Y = 166
                        $drawingSize.Height = 23
                        $drawingSize.Width = 75
                        $btnScheduleOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $btnScheduleOk.Location = $drawingPoint
                        $btnScheduleOk.Name = "btnScheduleOk"
                        $btnScheduleOk.Size = $drawingSize
                        $btnScheduleOk.TabIndex = 1
                        $btnScheduleOk.Text = "Ok"
                        $btnScheduleOk.Add_Click({
                            if ($radioStartAutomatic.Checked) {$Global:migrationStrategy = 10}
                            elseif ($radioStartSchedule.Checked) {$Global:migrationStrategy = 20}
                            if ($radioCompleteAutomatic.Checked) {$Global:migrationStrategy += 1}
                            elseif ($radioCompleteSchedule.Checked) {$Global:migrationStrategy += 2}

                            $Global:migrationEndpoint = $endpointBox.SelectedItem.ToString()
                            $Global:scheduleStartDateTime = $startSchedulePicker.Value.GetDateTimeFormats('u')
                            $Global:scheduleCompleteDateTime = $completeSchedulePicker.Value.GetDateTimeFormats('u')
                            $frmSchedule.Close()
                        })
                    #endregion
        
                    #region btnScheduleCancel
                        $drawingPoint.X = 619
                        $drawingPoint.Y = 166
                        $drawingSize.Height = 23
                        $drawingSize.Width = 75
                        $btnScheduleCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                        $btnScheduleCancel.Location = $drawingPoint
                        $btnScheduleCancel.Name = "btnScheduleCancel"
                        $btnScheduleCancel.Size = $drawingSize
                        $btnScheduleCancel.TabIndex = 2
                        $btnScheduleCancel.Text = "Cancel"
                    #endregion

                    #region lblSelectEndpoint
                        $drawingPoint.X = 113
                        $drawingPoint.Y = 139
                        $drawingSize.Height = 13
                        $drawingSize.Width = 228
                        $lblSelectEndpoint.AutoSize = $True
                        $lblSelectEndpoint.Location = $drawingPoint
                        $lblSelectEndpoint.Name = "lblSelectEndpoint"
                        $lblSelectEndpoint.Size = $drawingSize
                        $lblSelectEndpoint.Text = "Select the migrantion endpoint you wan to use:"
                    #endregion

                    #region endpointBox
                        $drawingPoint.X = 366
                        $drawingPoint.Y = 136
                        $drawingSize.Height = 31
                        $drawingSize.Width = 328
                        $endpointBox.Location = $drawingPoint
                        $endpointBox.Size = $drawingSize
                        $endpointBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
                        $endpointBox.FormattingEnabled = $True
                        $endpointBox.Name = "endpointBox"
                        $endpointBox.TabIndex = 2
                    #endregion

                    #region startSchedulePicker
                        $drawingPoint.X = 43
                        $drawingPoint.Y = 88
                        $drawingSize.Height = 20
                        $drawingSize.Width = 279
                        $startSchedulePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
                        $startSchedulePicker.Location = $drawingPoint
                        $startSchedulePicker.Name = "startSchedulePicker"
                        $startSchedulePicker.Size = $drawingSize
                        $startSchedulePicker.TabIndex = 0
                        $startSchedulePicker.CustomFormat = $format
                    #endregion

                    #region completeSchedulePicker
                        $drawingPoint.X = 43
                        $drawingPoint.Y = 88
                        $drawingSize.Height = 20
                        $drawingSize.Width = 279
                        $completeSchedulePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
                        $completeSchedulePicker.Location = $drawingPoint
                        $completeSchedulePicker.Name = "completeSchedulePicker"
                        $completeSchedulePicker.Size = $drawingSize
                        $completeSchedulePicker.TabIndex = 0
                        $completeSchedulePicker.CustomFormat = $format
                    #endregion

                    #region radioStartManual
                        $drawingPoint.X = 18
                        $drawingPoint.Y = 19
                        $drawingSize.Height = 17
                        $drawingSize.Width = 70
                        $radioStartManual.Location = $drawingPoint
                        $radioStartManual.Name = "radioStartManual"
                        $radioStartManual.Size = $drawingSize
                        $radioStartManual.TabIndex = 4
                        $radioStartManual.TabStop = $True
                        $radioStartManual.Text = "Manually"
                    #endregion

                    #region radioStartAutomatic
                        $drawingPoint.X = 18
                        $drawingPoint.Y = 42
                        $drawingSize.Height = 17
                        $drawingSize.Width = 90
                        $radioStartAutomatic.Location = $drawingPoint
                        $radioStartAutomatic.Name = "radioStartAutomatic"
                        $radioStartAutomatic.Size = $drawingSize
                        $radioStartAutomatic.TabIndex = 4
                        $radioStartAutomatic.TabStop = $True
                        $radioStartAutomatic.Text = "Automatically"
                    #endregion

                    #region radioStartSchedule
                        $drawingPoint.X = 18
                        $drawingPoint.Y = 65
                        $drawingSize.Height = 17
                        $drawingSize.Width = 218
                        $radioStartSchedule.Location = $drawingPoint
                        $radioStartSchedule.Name = "radioStartSchedule"
                        $radioStartSchedule.Size = $drawingSize
                        $radioStartSchedule.TabIndex = 4
                        $radioStartSchedule.TabStop = $True
                        $radioStartSchedule.Text = "Start the batch automatically after time:"
                        $radioStartSchedule.Checked = $True
                        $radioStartSchedule.Add_CheckedChanged({
                            if ($radioStartSchedule.Checked) {
                                $radioCompleteManual.Checked = $False
                                $radioCompleteManual.Enabled = $False
                            }
                            else {
                                $radioCompleteManual.Enabled = $True
                            }
                        })
                    #endregion

                    #region radioCompleteManual
                        $drawingPoint.X = 18
                        $drawingPoint.Y = 19
                        $drawingSize.Height = 17
                        $drawingSize.Width = 70
                        $radioCompleteManual.Enabled = $False
                        $radioCompleteManual.Checked = $False
                        $radioCompleteManual.Location = $drawingPoint
                        $radioCompleteManual.Name = "radioCompleteManual"
                        $radioCompleteManual.Size = $drawingSize
                        $radioCompleteManual.TabIndex = 4
                        $radioCompleteManual.TabStop = $True
                        $radioCompleteManual.Text = "Manually"
                    #endregion

                    #region radioCompleteAutomatic
                        $drawingPoint.X = 18
                        $drawingPoint.Y = 42
                        $drawingSize.Height = 17
                        $drawingSize.Width = 90
                        $radioCompleteAutomatic.Location = $drawingPoint
                        $radioCompleteAutomatic.Name = "radioCompleteAutomatic"
                        $radioCompleteAutomatic.Size = $drawingSize
                        $radioCompleteAutomatic.TabIndex = 4
                        $radioCompleteAutomatic.TabStop = $True
                        $radioCompleteAutomatic.Text = "Automatically"
                    #endregion

                    #region radioCompleteSchedule
                        $drawingPoint.X = 18
                        $drawingPoint.Y = 65
                        $drawingSize.Height = 17
                        $drawingSize.Width = 245
                        $radioCompleteSchedule.Location = $drawingPoint
                        $radioCompleteSchedule.Name = "radioCompleteSchedule"
                        $radioCompleteSchedule.Size = $drawingSize
                        $radioCompleteSchedule.TabIndex = 4
                        $radioCompleteSchedule.TabStop = $True
                        $radioCompleteSchedule.Text = "Complete the batch automatically after time:"
                        $radioCompleteSchedule.Checked = $True
                    #endregion

                    #region grpStart
                        $drawingPoint.X = 13
                        $drawingPoint.Y = 13
                        $drawingSize.Height = 115
                        $drawingSize.Width = 328
                        $grpStart.Controls.Add($radioStartManual)
                        $grpStart.Controls.Add($radioStartAutomatic)
                        $grpStart.Controls.Add($radioStartSchedule)
                        $grpStart.Controls.Add($startSchedulePicker)
                        $grpStart.Location = $drawingPoint
                        $grpStart.Name = "grpStart"
                        $grpStart.Size = $drawingSize
                        $grpStart.TabStop = $False
                        $grpStart.Text = "Please select the preferred option to start the batch"
                    #endregion

                    #region grpComplete
                        $drawingPoint.X = 366
                        $drawingPoint.Y = 13
                        $drawingSize.Height = 115
                        $drawingSize.Width = 328
                        $grpComplete.Controls.Add($radioCompleteManual)
                        $grpComplete.Controls.Add($radioCompleteAutomatic)
                        $grpComplete.Controls.Add($radioCompleteSchedule)
                        $grpComplete.Controls.Add($completeSchedulePicker)
                        $grpComplete.Location = $drawingPoint
                        $grpComplete.Name = "grpComplete"
                        $grpComplete.Size = $drawingSize
                        $grpComplete.TabStop = $False
                        $grpComplete.Text = "Please select the preferred option to start the batch"
                    #endregion

                    #region frmSchedule
                        $drawingSize.Height = 240
                        $drawingSize.Width = 722
                        $frmSchedule.Size = $drawingSize
                        $frmSchedule.AcceptButton = $btnScheduleOk
                        $frmSchedule.CancelButton = $btnScheduleCancel
                        $frmSchedule.Controls.Add($grpStart)
                        $frmSchedule.Controls.Add($grpComplete)
                        $frmSchedule.Controls.Add($lblSelectEndpoint)
                        $frmSchedule.Controls.Add($endpointBox)
                        $frmSchedule.Controls.Add($btnScheduleOk)
                        $frmSchedule.Controls.Add($btnScheduleCancel)
                        $frmSchedule.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
                        $frmSchedule.ControlBox = $False
                        $frmSchedule.MaximizeBox = $False
                        $frmSchedule.MinimizeBox = $False
                        $frmSchedule.Name = "frmSchedule"
                        $frmSchedule.Text = "Schedule migration"
                        $frmSchedule.Add_Load({
                            $startSchedulePicker.Value = [System.DateTime]::Now.AddDays(1)
                            $completeSchedulePicker.Value = [System.DateTime]::Now.AddDays(2)
                            $Global:endPointList | ForEach-Object {
                                if ($_ -ne "") {$endpointBox.Items.Add($_)}
                            }
                            $endpointBox.SelectedIndex = 0
                        })
                        $frmSchedule.ResumeLayout($False)
                        $frmSchedule.PerformLayout()

                        $result = $frmSchedule.ShowDialog()
                    #endregion
                #endregion

                #region schedule migration
                    if ($result = [System.Windows.Forms.DialogResult]::OK) {
                        $progressBar.Value = 0
                        $progressBar.Visible = $True
                        Write-Progress -Activity "Creating migration batch $batchName" -Status "Preparing list of users..." -PercentComplete ($progressBar.Value)
                        $statusLabel.Text = "Preparing list of users..."

                        $onlineTreeView.Nodes | ForEach-Object {
                            $migrationList += "`n$($_.Name)"
                        }

                        $progressBar.Value = 10
                        Write-Progress -Activity "Creating migration batch $batchName" -Status "Creating migration batch with strategy $Global:migrationStrategy..." -PercentComplete ($progressBar.Value)
                        $statusLabel.Text = "Creating migration batch with strategy $Global:migrationStrategy..."

                        $csvData = [System.Text.Encoding]::ASCII.GetBytes($migrationList)
                        switch($Global:migrationStrategy) {
                            0  {New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $csvData -AutoStart:$False -AutoComplete:$False}
                            1  {New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $csvData -AutoStart:$False -AutoComplete:$True}
                            2  {New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $csvData -AutoStart:$False -CompleteAfter $Global:scheduleCompleteDateTime}
                            10 {New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $csvData -AutoStart:$True  -AutoComplete:$False}
                            11 {New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $csvData -AutoStart:$True  -AutoComplete:$True}
                            12 {New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $csvData -AutoStart:$True  -CompleteAfter $Global:scheduleCompleteDateTime}
                            21 {New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $csvData -AutoComplete:$True  -StartAfter $Global:scheduleStartDateTime}
                            22 {New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $csvData -StartAfter $Global:scheduleStartDateTime -CompleteAfter $Global:scheduleCompleteDateTime}
                        }

                        $progressBar.Value = 100
                        Write-Progress -Activity "Creating migration batch $batchName" -Status "Done!" -PercentComplete ($progressBar.Value)
                        $statusLabel.Text = "Done creating migration batch $batchName!"

                        $progressBar.Visible = $False
                        $progressBar.Value = 0
                        Write-Progress -Activity "Creating migration batch $batchName" -Completed
                        $statusLabel.Text = ""
                        $frmSchedule = $null
                    }
                #endregion
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("You need to register a migration`nendpoint before continuing.", "Missing migration endpoint", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Start-Process "https://technet.microsoft.com/library/jj218611(v=exchg.160).aspx"
            }
        }
    }
#endregion

#region Declaring form objects
    [string] $folderPath = "$(Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)\Images"
    [string] $filePath = ""
    [System.Windows.Forms.Form] $frmMain = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.StatusStrip] $mainStatusStrip = New-Object -TypeName System.Windows.Forms.StatusStrip
    [System.Windows.Forms.ToolStripProgressBar] $progressBar = New-Object -TypeName System.Windows.Forms.ToolStripProgressBar
    [System.Windows.Forms.ToolStripStatusLabel] $statusLabel = New-Object -TypeName System.Windows.Forms.ToolStripStatusLabel
    [System.Windows.Forms.TreeView] $onPremisesTreeView = New-Object -TypeName System.Windows.Forms.TreeView
    [System.Windows.Forms.TreeView] $onlineTreeView = New-Object -TypeName System.Windows.Forms.TreeView
    [System.Windows.Forms.Button] $btnAdd = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.Button] $btnAddAll = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.Button] $btnRemove = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.Button] $btnRemoveAll = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.MenuStrip] $mainMenuStrip = New-Object -TypeName System.Windows.Forms.MenuStrip
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFile = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileConfigure = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileConnect = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileReload = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFilePreFlight = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileExport = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileMigrate = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileExit = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemHelp = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemHelpBlog = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemHelpAbout = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripSeparator] $menuItemFileSpace1 = New-Object -TypeName System.Windows.Forms.ToolStripSeparator
    [System.Windows.Forms.ToolStripSeparator] $menuItemFileSpace2 = New-Object -TypeName System.Windows.Forms.ToolStripSeparator
    [System.Windows.Forms.ToolStrip] $toolBar = New-Object -TypeName System.Windows.Forms.ToolStrip
    [System.Windows.Forms.ToolStripButton] $toolbarBtnConfiguration = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripButton] $toolbarBtnConnect = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripButton] $toolbarBtnReload = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripButton] $toolbarBtnPreFlight = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripButton] $toolbarBtnExport = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripButton] $toolbarBtnMigrate = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripSeparator] $toolbarSeparator1 = New-Object -TypeName System.Windows.Forms.ToolStripSeparator
    [System.Windows.Forms.Label] $lblAvailable = New-Object -TypeName System.Windows.Forms.Label
    [System.Windows.Forms.Label] $lblSelected = New-Object -TypeName System.Windows.Forms.Label
#endregion

#region setting form objects
    #region lblAvailable
        $drawingPoint.X = 9
        $drawingPoint.Y = 79
        $drawingSize.Height = 13
        $drawingSize.Width = 100
        $lblAvailable.AutoSize = $True
        $lblAvailable.Location = $drawingPoint
        $lblAvailable.Name = "Label1"
        $lblAvailable.Size = $drawingSize
        $lblAvailable.Text = "Available mailboxes"
    #endregion

    #region lblSelected
        $drawingPoint.X = 623
        $drawingPoint.Y = 79
        $drawingSize.Height = 13
        $drawingSize.Width = 100
        $lblSelected.AutoSize = $True
        $lblSelected.Location = $drawingPoint
        $lblSelected.Name = "Label1"
        $lblSelected.Size = $drawingSize
        $lblSelected.Text = "Selected mailboxes"
    #endregion

    #region progressBar
        $drawingSize.Height = 16
        $drawingSize.Width = 100
        $progressBar.Name = "progressBar"
        $progressBar.Size = $drawingSize
        $progressBar.Visible = $False
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    #endregion

    #region statusLabel
        $drawingSize.Height = 17
        $drawingSize.Width = 10
        $statusLabel.Name = "statusLabel"
        $statusLabel.Size = $drawingSize
        $statusLabel.Text = ""
        $statusLabel.RightToLeft = [System.Windows.Forms.RightToLeft]::No
        $statusLabel.Alignment = "Right"
    #endregion

    #region mainStatusStrip
        $drawingPoint.X = 0
        $drawingPoint.Y = 524
        $drawingSize.Height = 22
        $drawingSize.Width = 1173
        $mainStatusStrip.Location = $drawingPoint
        $mainStatusStrip.Name = "mainStatusStrip"
        $mainStatusStrip.Size = $drawingSize
        $mainStatusStrip.TabStop = $False
        $mainStatusStrip.SizingGrip = $False
        $mainStatusStrip.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
        #$mainStatusStrip.LayoutStyle = [System.Windows.Forms.ToolStripLayoutStyle]::Flow
        $mainStatusStrip.Items.AddRange(@($progressBar, $statusLabel))
    #endregion

    #region onPremisesTreeView
        $drawingPoint.X = 12
        $drawingPoint.Y = 99
        $drawingSize.Height = 412
        $drawingSize.Width = 535
        $onPremisesTreeView.Location = $drawingPoint
        $onPremisesTreeView.Name = "onPremisesTreeView"
        $onPremisesTreeView.Size = $drawingSize
        $onPremisesTreeView.CheckBoxes = $True
        $onPremisesTreeView.HideSelection = $False
        $onPremisesTreeView.HotTracking = $True
        $onPremisesTreeView.ShowLines = $False
        $onPremisesTreeView.ShowPlusMinus = $False
        $onPremisesTreeView.TabIndex = 1
    #endregion

    #region onlineTreeView
        $drawingPoint.X = 626
        $drawingPoint.Y = 99
        $drawingSize.Height = 412
        $drawingSize.Width = 535
        $onlineTreeView.Location = $drawingPoint
        $onlineTreeView.Name = "onlineTreeView"
        $onlineTreeView.Size = $drawingSize
        $onlineTreeView.CheckBoxes = $True
        $onlineTreeView.HideSelection = $False
        $onlineTreeView.HotTracking = $True
        $onlineTreeView.ShowLines = $False
        $onlineTreeView.ShowPlusMinus = $False
        $onlineTreeView.TabIndex = 6
    #endregion

    #region btnAdd
        $drawingPoint.X = 563
        $drawingPoint.Y = 211
        $drawingSize.Height = 23
        $drawingSize.Width = 48
        $btnAdd.Location = $drawingPoint
        $btnAdd.Name = "btnAdd"
        $btnAdd.Size = $drawingSize
        $btnAdd.TabIndex = 2
        $btnAdd.Text = ">"
        $btnAdd.UseVisualStyleBackColor = $True
        $btnAdd.Add_Click({
            if ($onPremisesTreeView.Nodes.Count -gt 0) {
                ($onPremisesTreeView.Nodes.Count - 1)..0 | ForEach-Object {
                    if ($onPremisesTreeView.Nodes[$_].Checked) {
                        $node = $onPremisesTreeView.Nodes[$_]
                        $onPremisesTreeView.Nodes[$_].Remove()
                        $onlineTreeView.Nodes.Add($node)
                    }
                }
            }
        })
    #endregion

    #region btnAddAll
        $drawingPoint.X = 563
        $drawingPoint.Y = 240
        $drawingSize.Height = 23
        $drawingSize.Width = 48
        $btnAddAll.Location = $drawingPoint
        $btnAddAll.Name = "btnAddAll"
        $btnAddAll.Size = $drawingSize
        $btnAddAll.TabIndex = 3
        $btnAddAll.Text = ">>"
        $btnAddAll.UseVisualStyleBackColor = $True
        $btnAddAll.Add_Click({
            if ($onPremisesTreeView.Nodes.Count -gt 0) {
                ($onPremisesTreeView.Nodes.Count - 1)..0 | ForEach-Object {
                    $node = $onPremisesTreeView.Nodes[$_]
                    $onPremisesTreeView.Nodes[$_].Remove()
                    $onlineTreeView.Nodes.Add($node)
                }
            }
        })
    #endregion

    #region btnRemove
        $drawingPoint.X = 563
        $drawingPoint.Y = 269
        $drawingSize.Height = 23
        $drawingSize.Width = 48
        $btnRemove.Location = $drawingPoint
        $btnRemove.Name = "btnRemove"
        $btnRemove.Size = $drawingSize
        $btnRemove.TabIndex = 4
        $btnRemove.Text = "<"
        $btnRemove.UseVisualStyleBackColor = $True
        $btnRemove.Add_Click({
            if ($onlineTreeView.Nodes.Count -gt 0) {
                ($onlineTreeView.Nodes.Count - 1)..0 | ForEach-Object {
                    if ($onlineTreeView.Nodes[$_].Checked) {
                        $node = $onlineTreeView.Nodes[$_]
                        $onlineTreeView.Nodes[$_].Remove()
                        $onPremisesTreeView.Nodes.Add($node)
                    }
                }
            }
        })
    #endregion

    #region btnRemoveAll
        $drawingPoint.X = 563
        $drawingPoint.Y = 298
        $drawingSize.Height = 23
        $drawingSize.Width = 48
        $btnRemoveAll.Location = $drawingPoint
        $btnRemoveAll.Name = "btnRemoveAll"
        $btnRemoveAll.Size = $drawingSize
        $btnRemoveAll.TabIndex = 5
        $btnRemoveAll.Text = "<<"
        $btnRemoveAll.UseVisualStyleBackColor = $True
        $btnRemoveAll.Add_Click({
            if ($onlineTreeView.Nodes.Count -gt 0) {
                ($onlineTreeView.Nodes.Count - 1)..0 | ForEach-Object {
                    $node = $onlineTreeView.Nodes[$_]
                    $onlineTreeView.Nodes[$_].Remove()
                    $onPremisesTreeView.Nodes.Add($node)
                }
            }
        })
    #endregion

    #region menuItemFileConfigure
        $drawingSize.Height = 22
        $drawingSize.Width = 152
        $menuItemFileConfigure.Name = "menuItemFileConfigure"
        $menuItemFileConfigure.Size = $drawingSize
        $menuItemFileConfigure.Text = "C&onfigure..."
        $menuItemFileConfigure.Add_Click({fnConfigure})
    #endregion

    #region menuItemFileConnect
        $drawingSize.Height = 22
        $drawingSize.Width = 152
        $menuItemFileConnect.Name = "menuItemFileConnect"
        $menuItemFileConnect.Size = $drawingSize
        $menuItemFileConnect.Text = "&Connect..."
        $menuItemFileConnect.Add_Click({fnConnect})
    #endregion

    #region menuItemFileReload
        $drawingSize.Height = 22
        $drawingSize.Width = 152
        $menuItemFileReload.Name = "menuItemFileReload"
        $menuItemFileReload.Size = $drawingSize
        $menuItemFileReload.Text = "&Reload"
        $menuItemFileReload.Add_Click({fnLoad})
    #endregion

    #region menuItemFilePreFlight
        $drawingSize.Height = 22
        $drawingSize.Width = 152
        $menuItemFilePreFlight.Name = "menuItemFilePreFlight"
        $menuItemFilePreFlight.Size = $drawingSize
        $menuItemFilePreFlight.Text = "Run &pre-flight"
        $menuItemFilePreFlight.Add_Click({fnRunPreFlight})
    #endregion

    #region menuItemFileExport
        $drawingSize.Height = 22
        $drawingSize.Width = 179
        $menuItemFileExport.Name = "menuItemFileExport"
        $menuItemFileExport.Size = $drawingSize
        $menuItemFileExport.Text = "&Export commands..."
        $menuItemFileExport.Add_Click({fnWriteScript})
    #endregion

    #region menuItemFileMigrate
        $drawingSize.Height = 22
        $drawingSize.Width = 152
        $menuItemFileMigrate.Name = "menuItemFileMigrate"
        $menuItemFileMigrate.Size = $drawingSize
        $menuItemFileMigrate.Text = "&Migrate"
        $menuItemFileMigrate.Add_Click({fnSchedule})
    #endregion

    #region menuItemFileSpace1
        $drawingSize.Height = 6
        $drawingSize.Width = 149
        $menuItemFileSpace1.Name = "menuItemFileSpace1"
        $menuItemFileSpace1.Size = $drawingSize
    #endregion

    #region menuItemFileSpace2
        $drawingSize.Height = 6
        $drawingSize.Width = 149
        $menuItemFileSpace1.Name = "menuItemFileSpace2"
        $menuItemFileSpace1.Size = $drawingSize
    #endregion

    #region menuItemFileExit
        $drawingSize.Height = 22
        $drawingSize.Width = 152
        $menuItemFileExit.Name = "menuItemFileExit"
        $menuItemFileExit.Size = $drawingSize
        $menuItemFileExit.Text = "E&xit"
        $menuItemFileExit.Add_Click({$frmMain.Close()})
    #endregion

    #region menuItemFile
        $drawingSize.Height = 20
        $drawingSize.Width = 37
        $menuItemFile.DropDownItems.AddRange(@($menuItemFileConfigure, $menuItemFileConnect, $menuItemFileReload, $menuItemFileSpace1, $menuItemFilePreFlight, $menuItemFileExport, $menuItemFileMigrate, $menuItemFileSpace2, $menuItemFileExit))
        $menuItemFile.Name = "menuItemFile"
        $menuItemFile.Size = $drawingSize
        $menuItemFile.Text = "&File"
    #endregion

    #region menuItemHelpBlog
        $drawingSize.Height = 22
        $drawingSize.Width = 207
        $menuItemHelpBlog.Name = "menuItemHelpBlog"
        $menuItemHelpBlog.Size = $drawingSize
        $menuItemHelpBlog.Text = "Open &FastTrack Tips Blog"
        $menuItemHelpBlog.Add_Click({Start-Process "http://aka.ms/ftctips"})
    #endregion

    #region menuItemHelpAbout
        $drawingSize.Height = 22
        $drawingSize.Width = 207
        $menuItemHelpAbout.Name = "menuItemHelpAbout"
        $menuItemHelpAbout.Size = $drawingSize
        $menuItemHelpAbout.Text = "&About"
        $menuItemHelpAbout.Add_Click({fnAbout})
    #endregion

    #region menuItemHelp
        $drawingSize.Height = 20
        $drawingSize.Width = 44
        $menuItemHelp.DropDownItems.AddRange(@($menuItemHelpBlog, $menuItemHelpAbout))
        $menuItemHelp.Name = "menuItemHelp"
        $menuItemHelp.Size = $drawingSize
        $menuItemHelp.Text = "&Help"
    #endregion

    #region mainMenuStrip
        $drawingPoint.X = 0
        $drawingPoint.Y = 0
        $drawingSize.Height = 24
        $drawingSize.Width = 1173
        $mainMenuStrip.Items.AddRange(@($menuItemFile, $menuItemHelp))
        $mainMenuStrip.Location = $drawingPoint
        $mainMenuStrip.Name = "mainMenuStrip"
        $mainMenuStrip.Size = $drawingSize
        $mainMenuStrip.TabIndex = 7
        $mainMenuStrip.Text = "mainMenuStrip"
    #endregion

    #region toolbarBtnConfiguration
        $filePath = "$folderPath\Configuration48.png"
        $drawingSize.Height = 52
        $drawingSize.Width = 52
        $toolbarBtnConfiguration.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Image
        $toolbarBtnConfiguration.Image = [System.Drawing.Image]::FromFile($filePath)
        $toolbarBtnConfiguration.ImageTransparentColor = [System.Drawing.Color]::Magenta
        $toolbarBtnConfiguration.Name = "toolbarBtnConfiguration"
        $toolbarBtnConfiguration.Size = $drawingSize
        $toolbarBtnConfiguration.Text = "Configuration"
        $toolbarBtnConfiguration.Add_Click({fnConfigure})
    #endregion

    #region toolbarBtnConnect
        $filePath = "$folderPath\Connect48.png"
        $drawingSize.Height = 52
        $drawingSize.Width = 52
        $toolbarBtnConnect.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Image
        $toolbarBtnConnect.Image = [System.Drawing.Image]::FromFile($filePath)
        $toolbarBtnConnect.ImageTransparentColor = [System.Drawing.Color]::Magenta
        $toolbarBtnConnect.Name = "toolbarBtnConnect"
        $toolbarBtnConnect.Size = $drawingSize
        $toolbarBtnConnect.Text = "Connect"
        $toolbarBtnConnect.Add_Click({fnConnect})
    #endregion

    #region toolbarBtnReload
        $filePath = "$folderPath\Reload48.png"
        $drawingSize.Height = 52
        $drawingSize.Width = 52
        $toolbarBtnReload.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Image
        $toolbarBtnReload.Image = [System.Drawing.Image]::Fromfile($filePath)
        $toolbarBtnReload.ImageTransparentColor = [System.Drawing.Color]::Magenta
        $toolbarBtnReload.Name = "toolbarBtnReload"
        $toolbarBtnReload.Size = $drawingSize
        $toolbarBtnReload.Text = "Reload"
        $toolbarBtnReload.Add_Click({fnLoad})

    #endregion

    #region toolbarBtnPreFlight
        $filePath = "$folderPath\Pre-flight48.png"
        $drawingSize.Height = 52
        $drawingSize.Width = 52
        $toolbarBtnPreFlight.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Image
        $toolbarBtnPreFlight.Image = [System.Drawing.Image]::Fromfile($filePath)
        $toolbarBtnPreFlight.ImageTransparentColor = [System.Drawing.Color]::Magenta
        $toolbarBtnPreFlight.Name = "toolbarBtnPreFlight"
        $toolbarBtnPreFlight.Size = $drawingSize
        $toolbarBtnPreFlight.Text = "Run pre-flight"
        $toolbarBtnPreFlight.Add_Click({fnRunPreFlight})
    #endregion

    #region toolbarBtnExport
        $filePath = "$folderPath\Export48.png"
        $drawingSize.Height = 52
        $drawingSize.Width = 52
        $toolbarBtnExport.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Image
        $toolbarBtnExport.Image = [System.Drawing.Image]::Fromfile($filePath)
        $toolbarBtnExport.ImageTransparentColor = [System.Drawing.Color]::Magenta
        $toolbarBtnExport.Name = "toolbarBtnExport"
        $toolbarBtnExport.Size = $drawingSize
        $toolbarBtnExport.Text = "Export commands"
        $toolbarBtnExport.Add_Click({fnWriteScript})
    #endregion

    #region toolbarBtnMigrate
        $filePath = "$folderPath\Migrate48.png"
        $drawingSize.Height = 52
        $drawingSize.Width = 52
        $toolbarBtnMigrate.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Image
        $toolbarBtnMigrate.Image = [System.Drawing.Image]::Fromfile($filePath)
        $toolbarBtnMigrate.ImageTransparentColor = [System.Drawing.Color]::Magenta
        $toolbarBtnMigrate.Name = "toolbarBtnMigrate"
        $toolbarBtnMigrate.Size = $drawingSize
        $toolbarBtnMigrate.Text = "Migrate"
        $toolbarBtnMigrate.Add_Click({fnSchedule})
    #endregion

    #region toolbarSeparator1
        $drawingSize.Height = 55
        $drawingSize.Width = 6
        $toolbarSeparator1.Name = "toolbarSeparator1"
        $toolbarSeparator1.Size = $drawingSize
    #endregion

    #region toolBar
        $drawingPoint.X = 0
        $drawingPoint.Y = 24
        $drawingSize.Height = 48
        $drawingSize.Width = 48
        $toolBar.ImageScalingSize = $drawingSize
        $toolBar.Items.AddRange(@($toolbarBtnConfiguration, $toolbarBtnConnect, $toolbarBtnReload, $toolbarSeparator1, $toolbarBtnPreFlight, $toolbarBtnExport, $toolbarBtnMigrate))
        $toolBar.Location = $drawingPoint
        $toolBar.Name = "toolBar"
        $drawingSize.Height = 55
        $drawingSize.Width = 1173
        $toolBar.Size = $drawingSize
        $toolBar.Text = "toolBar"
    #endregion

    #region frmMain
        $drawingSize.Height = 585
        $drawingSize.Width = 1189
        $frmMain.Size = $drawingSize
        $frmMain.MinimumSize = $drawingSize
        $frmMain.MaximumSize = $drawingSize
        $drawingSize.Height = 546
        $drawingSize.Width = 1173
        $frmMain.ClientSize = $drawingSize
        $frmMain.MaximizeBox = $False
        $frmMain.Name = "frmMain"
        $frmMain.Text = "Mailbox migration"
        $frmMain.Icon = [System.Drawing.Icon]::FromHandle($toolbarBtnConnect.Image.GetHicon())
        $frmMain.Add_Closed({fnDisconnect})
    #endregion

#endregion

#region Loading form
    $frmMain.ResumeLayout($False)
    $frmMain.PerformLayout()
    $frmMain.Controls.Add($toolBar)
    $frmMain.Controls.Add($btnRemoveAll)
    $frmMain.Controls.Add($btnRemove)
    $frmMain.Controls.Add($btnAddAll)
    $frmMain.Controls.Add($btnAdd)
    $frmMain.Controls.Add($onlineTreeView)
    $frmMain.Controls.Add($onPremisesTreeView)
    $frmMain.Controls.Add($mainStatusStrip)
    $frmMain.Controls.Add($mainMenuStrip)
    $frmMain.Controls.Add($lblAvailable)
    $frmMain.Controls.Add($lblSelected)
    $frmMain.WindowState = $windowState
    [void] $frmMain.ShowDialog()
#endregion
