Param ([switch] $AutoLoad)
[string] $version = "1.8.5"

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
    Add-Type -Language VisualBasic -ReferencedAssemblies ("System.Collections", "System.Windows.Forms") -TypeDefinition @"
    Imports System.Collections
    Imports System.Windows.Forms

    Public Class PreFlightItem
        Public primarySMTPAddress As String
        Public status As String
        Public errorMessage As String
    End Class

    Public Class ListViewColumnSorter
        Implements System.Collections.IComparer

        Private ColumnToSort As Integer
        Private OrderOfSort As SortOrder
        Private ObjectCompare As CaseInsensitiveComparer

        Public Sub New()
            ' Initialize the column to '0'.
            ColumnToSort = 0

            ' Initialize the sort order to 'none'.
            OrderOfSort = SortOrder.None

            ' Initialize the CaseInsensitiveComparer object.
            ObjectCompare = New CaseInsensitiveComparer()
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            Dim compareResult As Integer
            Dim listviewX As ListViewItem
            Dim listviewY As ListViewItem

            ' Cast the objects to be compared to ListViewItem objects.
            listviewX = CType(x, ListViewItem)
            listviewY = CType(y, ListViewItem)

            ' Compare the two items.
            compareResult = ObjectCompare.Compare(listviewX.SubItems(ColumnToSort).Text, listviewY.SubItems(ColumnToSort).Text)

            ' Calculate the correct return value based on the object 
            ' comparison.
            If (OrderOfSort = SortOrder.Ascending) Then
                ' Ascending sort is selected, return typical result of 
                ' compare operation.
                Return compareResult
            ElseIf (OrderOfSort = SortOrder.Descending) Then
                ' Descending sort is selected, return negative result of 
                ' compare operation.
                Return (-compareResult)
            Else
                ' Return '0' to indicate that they are equal.
                Return 0
            End If
        End Function

        Public Property SortColumn() As Integer
            Set(ByVal Value As Integer)
                ColumnToSort = Value
            End Set

            Get
                Return ColumnToSort
            End Get
        End Property

        Public Property Order() As SortOrder
            Set(ByVal Value As SortOrder)
                OrderOfSort = Value
            End Set

            Get
                Return OrderOfSort
            End Get
        End Property
    End Class
"@
#endregion

#region global variables
    [Boolean] $Global:configurationFinished = $False
    [Boolean] $Global:localCredentialChanged = $False
    [Boolean] $Global:cloudCredentialChanged = $False
    [PSCredential] $Global:localCred = New-Object System.Management.Automation.PSCredential ("dummy", (ConvertTo-SecureString "dummy" -AsPlainText -Force))
    [PSCredential] $Global:cloudCred = New-Object System.Management.Automation.PSCredential ("dummy", (ConvertTo-SecureString "dummy" -AsPlainText -Force))
    [Boolean] $Global:isConnected = $False
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
        [Boolean] $continue = $True
        

        $progressBar.Value = 10
        $progressBar.Visible = $True
        $statusLabel.Text = "Connecting to Exchange Online..."

        $cloudSession = Get-PSSession | Where-Object {(($_.ComputerName -eq "outlook.office365.com") -or ($_.ComputerName -eq "ps.outlook.com")) -and ($_.ConfigurationName -eq "Microsoft.Exchange")}
        if ($CloudSession) {
		    Write-Host "Already connected to Exchange Online" -ForegroundColor Blue
            $Global:isConnected = [Boolean] ($CloudSession)
	    if ($Global:serviceDomain -eq "") {
	    	$Global:serviceDomain = (Get-AcceptedDomain | Where-Object {$_.DomainName -like "*.mail.onmicrosoft.com"}).DomainName.ToString()
	    }
        }
        else {
            if ($Global:cloudCred.UserName -eq "dummy") {
                $result = fnConfigure
                $continue = ($result -eq [System.Windows.Forms.DialogResult]::OK)
            }
            if ($continue) {
                if ((Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue).Count -eq 0) {
			        $cloudSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell" -AllowRedirection -Credential $Global:cloudCred -Authentication Basic
			        Import-PSSession $cloudSession -CommandName Get-Mailbox, Get-MailUser, New-MoveRequest, Get-AcceptedDomain, New-MigrationBatch, Get-MigrationEndpoint

                    $cloudSession = Get-PSSession | Where-Object {($_.ComputerName -eq "ps.outlook.com") -and ($_.ConfigurationName -eq "Microsoft.Exchange")}
                    $Global:isConnected = [Boolean] ($CloudSession)
                }
                else {
                    Connect-ExchangeOnline -UserPrincipalName $Global:cloudCred.UserName

                    $cloudSession = Get-PSSession | Where-Object {($_.ComputerName -eq "outlook.office365.com") -and ($_.ConfigurationName -eq "Microsoft.Exchange")}
                    $Global:isConnected = [Boolean] ($CloudSession)
                }
            }
            if ($Global:isConnected) {
                fnLoad -LoadUsers:$AutoLoad
            }
        }

        $progressBar.Visible = $False
        $statusLabel.Text = ""
        return $continue
    }
#endregion

#region fnLoad
    Function fnLoad {
        Param ([switch] $LoadUsers)
        $onPremisesListView.Items.Clear()
        $onlineListView.Items.Clear()
        $Global:endPointList.Clear()

        if (-not $Global:isConnected) {
            Write-Host "You are not connected to Exchange yet." -ForegroundColor Red
        }
        else {
            $progressBar.Value = 25
            $progressBar.Visible = $True
            If ($LoadUsers) {
                $onPremisesListView.BackgroundImage = $null
                $statusLabel.Text = "Loading list of users available for migration..."
                Get-MailUser -ResultSize Unlimited | Sort-Object Name | Where-Object {$_.ExchangeGuid -ne "00000000-0000-0000-0000-000000000000"} | ForEach-Object {
                    $listViewItem = New-Object -TypeName System.Windows.Forms.ListViewItem([System.String[]](@($_.DisplayName, $_.PrimarySmtpAddress)), -1)
                    $ListViewItem.Checked = $False
                    $onPremisesListView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($listViewItem)))
                }
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
    }
#endregion

#region fnConfigure
    Function fnConfigure {
        $Global:localCredentialChanged = $False
        $Global:cloudCredentialChanged = $False
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
            $txtLocalUser.add_TextChanged({
                $Global:cloudCredentialChanged = $True
                if (($txtcloudUser.Text -ne "") -and ($txtcloudPassword.Text -ne "")) {$btnOk.Enabled = $True}
                else {$btnOk.Enabled = $False}
            })
        #endregion

        #region txtLocalPassword
            $drawingPoint.X = 151
            $drawingPoint.Y = 52
            $drawingSize.Height = 20
            $drawingSize.Width = 179
            $txtLocalPassword.Location = $drawingPoint
            $txtLocalPassword.Size = $drawingSize
            $txtLocalPassword.UseSystemPasswordChar = $True
            $txtLocalPassword.Add_TextChanged({$Global:localCredentialChanged = $True})
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
            $txtCloudUser.Add_TextChanged({
                $Global:cloudCredentialChanged = $True
                if (($txtcloudUser.Text -ne "") -and ($txtcloudPassword.Text -ne "")) {$btnOk.Enabled = $True}
                else {$btnOk.Enabled = $False}
            })
        #endregion

        #region txtCloudPassword
            $drawingPoint.X = 151
            $drawingPoint.Y = 52
            $drawingSize.Height = 20
            $drawingSize.Width = 179
            $txtCloudPassword.Location = $drawingPoint
            $txtCloudPassword.Size = $drawingSize
            $txtCloudPassword.UseSystemPasswordChar = $True
            $txtCloudPassword.Add_TextChanged({
                $Global:cloudCredentialChanged = $True
                if (($txtcloudUser.Text -ne "") -and ($txtcloudPassword.Text -ne "")) {$btnOk.Enabled = $True}
                else {$btnOk.Enabled = $False}
            })
            if ((Get-Command Connect-EXOPSSession -ErrorAction SilentlyContinue).Count -gt 0) {
                $txtCloudPassword.Enabled = $False
                $txtCloudPassword.Visible = $False
                $txtCloudPassword.Text = "dummy"
            }
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
            $lblCloudPassword.Text = "Password"
            if ((Get-Command Connect-EXOPSSession -ErrorAction SilentlyContinue).Count -gt 0) {
                $lblCloudPassword.Enabled = $False
                $lblCloudPassword.Visible = $False
            }
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
            $btnOk.Enabled = $False
            $btnOk.Add_Click({
                if ($Global:localCredentialChanged -and ($txtLocalUser.Text -ne "") -and ($txtLocalPassword.Text -ne "")) {$Global:localCred = New-Object System.Management.Automation.PSCredential ($txtLocalUser.Text, (ConvertTo-SecureString $txtLocalPassword.Text -AsPlainText -Force))}
                if ($Global:cloudCredentialChanged -and ($txtcloudUser.Text -ne "") -and ($txtcloudPassword.Text -ne "")) {$Global:cloudCred = New-Object System.Management.Automation.PSCredential ($txtcloudUser.Text, (ConvertTo-SecureString $txtcloudPassword.Text -AsPlainText -Force))}
                if ($Global:cloudCred.UserName -ne "dummy") {$Global:configurationFinished = $True}
                $Global:localExchange = $txtLocalExchange.Text
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
            $grpLocal.Text = "Exchange on-premises (optional - used only for pre-flight)"
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
        $totalMailboxes = $onlineListView.Items.Count

        if ($totalMailboxes -gt 0) {
            $progressBar.Value = 0
            $progressBar.Visible = $True
            $onlineListView.Items | ForEach-Object {
                $reportItem = New-Object -TypeName PreFlightItem
                $reportItem.primarySMTPAddress = $_.SubItems.Item(1).Text
                $Error.Clear()
                Write-Progress -Activity "Running pre-flight" -Status "Checking $($_.SubItems.Item(0).Text) - $([math]::Round($progressBar.Value))% complete" -PercentComplete ($progressBar.Value)
                $statusLabel.Text = "Checking $($_.SubItems.Item(0).Text)"
                try {
                    New-MoveRequest -Remote -RemoteHostName $Global:localExchange -RemoteCredential $Global:localCred -Identity $_.SubItems.Item(1).Text -TargetDeliveryDomain $Global:serviceDomain -BatchName "PreFlight" -ErrorAction Stop -WhatIf
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
                Write-Progress -Activity "Running pre-flight" -Status "Checked $($_.SubItems.Item(0).Text) - $([math]::Round($progressBar.Value))% complete" -PercentComplete ($progressBar.Value)
                $statusLabel.Text = "Checked $($_.SubItems.Item(0).Text)"
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

        [string] $reportFilePath = [System.IO.Path]::GetFullPath("$($Script:MyInvocation.MyCommand.Path)\..\Reports\$(Get-Date -Format "yyyyMMdd-HHmmss").csv")

        "primarySMTPAddress,status,errorMessage" | Out-File -FilePath $reportfilePath -Encoding ascii -Force

        $ReportData | ForEach-Object {
            [string] $reportLine = "$($_.primarySMTPAddress),$($_.status),$($_.errorMessage)"
            $reportLine | Out-File -FilePath $reportFilePath -Encoding ascii -Append
        }
    }
#endregion

#region fnGetSchedule
    Function fnGetSchedule {
        $Global:migrationStrategy = 0
        $Global:scheduleStartDateTime = ""
        $Global:scheduleCompleteDateTime = ""

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
            $btnScheduleOk.Text = "Ok"
            $btnScheduleOk.Add_Click({
                if ($radioStartAutomatic.Checked) {$Global:migrationStrategy = 10}
                elseif ($radioStartSchedule.Checked) {$Global:migrationStrategy = 20}
                if ($radioCompleteAutomatic.Checked) {$Global:migrationStrategy += 1}
                elseif ($radioCompleteSchedule.Checked) {$Global:migrationStrategy += 2}

                $Global:migrationEndpoint = $endpointBox.SelectedItem.ToString()
                $Global:scheduleStartDateTime = [System.TimeZoneInfo]::ConvertTimeToUtc($startSchedulePicker.Value).GetDateTimeFormats('u')
                $Global:scheduleCompleteDateTime = [System.TimeZoneInfo]::ConvertTimeToUtc($completeSchedulePicker.Value).GetDateTimeFormats('u')
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
            $frmSchedule = $null
        #endregion
        return $result
    }
#endregion

#region fnWriteScript
    Function fnWriteScript {
        [Int] $currentMailbox = 0
        [Int] $totalMailboxes = $onlineListView.Items.Count
        [string] $batchName = "$(Get-Date -Format "yyyyMMdd-HHmmss")"
        [string] $scriptFileName = "$batchName.ps1"
        [string] $format = [System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.FullDateTimePattern
        [string] $scriptFilePath = ""
        [System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.DialogResult]::OK
        [System.Windows.Forms.SaveFileDialog] $saveDialog = New-Object -TypeName System.Windows.Forms.SaveFileDialog

        if ($totalMailboxes -gt 0) {
            if ($Global:endPointList.Count -gt 0) {
                $result = fnGetSchedule
                $saveDialog.InitialDirectory = [System.IO.Path]::GetFullPath("$($Script:MyInvocation.MyCommand.Path)\..\Scripts")
                $saveDialog.Filter = "Windows PowerShell Script (*.ps1)|*.ps1|All files (*.*)|*.*"
                $saveDialog.FileName = $scriptFileName
                if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                        $scriptFilePath = $saveDialog.FileName
                        '$cloudSession = Get-PSSession | Where-Object {($_.ComputerName -eq "ps.outlook.com") -and ($_.ConfigurationName -eq "Microsoft.Exchange")}' | Out-File -FilePath $scriptFilePath -Encoding ascii -Force
                        'if ($cloudSession) {' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        '    $disconnectAtTheEnd = $False' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
	    	            '    Write-Host "Already connected to Exchange Online" -ForegroundColor Blue' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        '}' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
    		            'else {' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        '    $disconnectAtTheEnd = $True' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        '    $cloudCred = Get-Credential -Message "Enter your cloud credential"' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
	    	            '    $cloudSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell" -AllowRedirection -Credential $cloudCred -Authentication Basic' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
		                '    Import-PSSession $cloudSession -CommandName New-MigrationBatch' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        '}' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        '' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        '$migrationList = "EmailAddress"' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        $onlineListView.Items | ForEach-Object {
                            "$('$migrationList += "`r`n')$($_.SubItems.Item(1).Text)""" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        }
                        '' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        '$csvData = [System.Text.Encoding]::ASCII.GetBytes($migrationList)' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        '' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        switch($Global:migrationStrategy) {
                            0  {"New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $('$csvData') -AutoStart:$('$False') -AutoComplete:$('$False')" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append}
                            1  {"New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $('$csvData') -AutoStart:$('$False') -AutoComplete:$('$True')" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append}
                            2  {"New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $('$csvData') -AutoStart:$('$False') -CompleteAfter '$Global:scheduleCompleteDateTime'" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append}
                            10 {"New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $('$csvData') -AutoStart:$('$True')  -AutoComplete:$('$False')" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append}
                            11 {"New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $('$csvData') -AutoStart:$('$True')  -AutoComplete:$('$True')" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append}
                            12 {"New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $('$csvData') -AutoStart:$('$True')  -CompleteAfter '$Global:scheduleCompleteDateTime'" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append}
                            21 {"New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $('$csvData') -AutoComplete:$('$True')  -StartAfter '$Global:scheduleStartDateTime'" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append}
                            22 {"New-MigrationBatch -Name $batchName -SourceEndpoint $Global:migrationEndpoint -TargetDeliveryDomain $Global:serviceDomain -TimeZone 'UTC' -CsvData $('$csvData') -StartAfter '$Global:scheduleStartDateTime' -CompleteAfter '$Global:scheduleCompleteDateTime'" | Out-File -FilePath $scriptFilePath -Encoding ascii -Append}
                        }
                        '' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append
                        'if ($disconnectAtTheEnd) {Remove-PSSession $cloudSession}' | Out-File -FilePath $scriptFilePath -Encoding ascii -Append -NoNewline
                    }
                }
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("You need to register a migration`r`nendpoint before continuing.", "Missing migration endpoint", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Start-Process "https://technet.microsoft.com/library/jj218611(v=exchg.160).aspx"
            }
        }
    }
#endregion

#region fnMigrate
    Function fnMigrate {
        [Int] $currentMailbox = 0
        [Int] $totalMailboxes = $onlineListView.Items.Count
        [string] $migrationList = "EmailAddress"
        [string] $batchName = "$(Get-Date -Format "yyyyMMdd-HHmmss")"
        [string] $format = [System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.FullDateTimePattern
        [System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.DialogResult]::OK

        if ($totalMailboxes -gt 0) {
            if ($Global:endPointList.Count -gt 0) {
                $result = fnGetSchedule

                if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                    $progressBar.Value = 0
                    $progressBar.Visible = $True
                    Write-Progress -Activity "Creating migration batch $batchName" -Status "Preparing list of users..." -PercentComplete ($progressBar.Value)
                    $statusLabel.Text = "Preparing list of users..."

                    $onlineListView.Items | ForEach-Object {
                        $migrationList += "`r`n$($_.SubItems.Item(1).Text)"
                    }

                    $progressBar.Value = 10
                    Write-Progress -Activity "Creating migration batch $batchName" -Status "Creating migration batch with strategy $Global:migrationStrategy..." -PercentComplete ($progressBar.Value)
                    $statusLabel.Text = "Creating migration batch $batchName with strategy $Global:migrationStrategy..."

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
                }
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("You need to register a migration`r`nendpoint before continuing.", "Missing migration endpoint", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Start-Process "https://technet.microsoft.com/library/jj218611(v=exchg.160).aspx"
            }
        }
    }
#endregion

#region fnImport
    Function fnImport {
        [System.Windows.Forms.OpenFileDialog] $openDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
        $openDialog.CheckFileExists = $True
        $openDialog.InitialDirectory = [System.IO.Path]::GetFullPath("$($Script:MyInvocation.MyCommand.Path)\..\Scripts")
        $openDialog.Filter = "Comma Separated Values file (*.csv)|*.csv|All files (*.*)|*.*"
        $openDialog.DefaultExt = "csv"
        $openDialog.Title = "Import csv file"
        $openDialog.Multiselect = $False

        if ($Global:isConnected){
            if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                fnRemoveAll
                Import-Csv -Path $openDialog.FileName | ForEach-Object {
                    If ($onPremisesListView.Items.Count -gt 0) {
                        $itemFound = $onPremisesListView.FindItemWithText($_.PrimarySMTPAddress, $True, 0)
                    }
                    Else {
                        $itemFound = $null
                    }

                    if ($null -ne $itemFound) {
                        $onPremisesListView.Items[$itemFound.Index].Remove()
                        $onlineListView.Items.Add($itemFound)
                    }
                    else {
                        $User = Get-MailUser -Identity $_.PrimarySMTPAddress -ErrorAction SilentlyContinue | Where-Object {$_.ExchangeGuid -ne "00000000-0000-0000-0000-000000000000"}
                        If ($null -ne $User) {
                            Write-Host "Object loaded: $($_.PrimarySMTPAddress)" -ForegroundColor Yellow
                            $listViewItem = New-Object -TypeName System.Windows.Forms.ListViewItem([System.String[]](@($User.DisplayName, $User.PrimarySmtpAddress)), -1)
                            $listViewItem.Checked = $False
                            $onlineListView.Items.Add($listViewItem)
                        }
                        else {
                            Write-Host "Object not found in Exchange Online: $($_.PrimarySMTPAddress)" -ForegroundColor Red
                            $listViewItem = New-Object -TypeName System.Windows.Forms.ListViewItem([System.String[]](@($_.PrimarySmtpAddress, $_.PrimarySmtpAddress)), -1)
                            $listViewItem.Checked = $False
                            $listViewItem.Font = New-Object -TypeName System.Drawing.Font ($listViewItem.Font, [System.Drawing.FontStyle]::Italic)
                            $onlineListView.Items.Add($listViewItem)
                        }
                    }
                }
            }
        }
    }
#endregion

#region fnRemoveAll
    Function fnRemoveAll {
        if ($onlineListView.Items.Count -gt 0) {
            ($onlineListView.Items.Count - 1)..0 | ForEach-Object {
                $itemFound = $onlineListView.Items[$_]
                $onlineListView.Items[$_].Remove()
                $onPremisesListView.Items.Add($itemFound)
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
    [System.Windows.Forms.ColumnHeader] $columnOnPremises1 = New-Object -TypeName System.Windows.Forms.ColumnHeader
    [System.Windows.Forms.ColumnHeader] $columnOnPremises2 = New-Object -TypeName System.Windows.Forms.ColumnHeader
    [System.Windows.Forms.ColumnHeader] $columnOnline1 = New-Object -TypeName System.Windows.Forms.ColumnHeader
    [System.Windows.Forms.ColumnHeader] $columnOnline2 = New-Object -TypeName System.Windows.Forms.ColumnHeader
    [System.Windows.Forms.ListView] $onPremisesListView = New-Object -TypeName System.Windows.Forms.ListView
    [System.Windows.Forms.ListView] $onlineListView = New-Object -TypeName System.Windows.Forms.ListView
    [System.Windows.Forms.Button] $btnAdd = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.Button] $btnAddAll = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.Button] $btnRemove = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.Button] $btnRemoveAll = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.MenuStrip] $mainMenuStrip = New-Object -TypeName System.Windows.Forms.MenuStrip
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFile = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileConfigure = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileConnect = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileReload = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripMenuItem] $menuItemFileImport = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
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
    [System.Windows.Forms.ToolStripButton] $toolbarBtnImport = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripButton] $toolbarBtnPreFlight = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripButton] $toolbarBtnExport = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripButton] $toolbarBtnMigrate = New-Object -TypeName System.Windows.Forms.ToolStripButton
    [System.Windows.Forms.ToolStripSeparator] $toolbarSeparator1 = New-Object -TypeName System.Windows.Forms.ToolStripSeparator
    [System.Windows.Forms.Label] $lblAvailable = New-Object -TypeName System.Windows.Forms.Label
    [System.Windows.Forms.Label] $lblSelected = New-Object -TypeName System.Windows.Forms.Label
    [System.Windows.Forms.Label] $lblSearchOnPremises = New-Object -TypeName System.Windows.Forms.Label
    [System.Windows.Forms.Label] $lblSearchOnline = New-Object -TypeName System.Windows.Forms.Label
    [System.Windows.Forms.TextBox] $txtSearchOnPremises = New-Object -TypeName System.Windows.Forms.TextBox
    [System.Windows.Forms.TextBox] $txtSearchOnline = New-Object -TypeName System.Windows.Forms.TextBox
    [ListViewColumnSorter] $localSorter = New-Object -TypeName ListViewColumnSorter
    [ListViewColumnSorter] $onlineSorter = New-Object -TypeName ListViewColumnSorter
#endregion

#region setting form objects
    #region lblAvailable
        $drawingPoint.X = 9
        $drawingPoint.Y = 79
        $drawingSize.Height = 13
        $drawingSize.Width = 100
        $lblAvailable.AutoSize = $True
        $lblAvailable.Location = $drawingPoint
        $lblAvailable.Name = "lblAvailable"
        $lblAvailable.Size = $drawingSize
        $lblAvailable.Text = "Available mailboxes"
        $lblAvailable.TabStop = $False
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
        $lblSelected.TabStop = $False
    #endregion

    #region lblSearchOnPremises
        $drawingPoint.X = 12
        $drawingPoint.Y = 102
        $drawingSize.Height = 13
        $drawingSize.Width = 41
        $lblSearchOnPremises.AutoSize = $True
        $lblSearchOnPremises.Location = $drawingPoint
        $lblSearchOnPremises.Name = "lblSearchOnPremises"
        $lblSearchOnPremises.Size = $drawingSize
        $lblSearchOnPremises.Text = "Search:"
        $lblSearchOnPremises.TabStop = $False
    #endregion

    #region lblSearchOnline
        $drawingPoint.X = 623
        $drawingPoint.Y = 102
        $drawingSize.Height = 13
        $drawingSize.Width = 41
        $lblSearchOnline.AutoSize = $True
        $lblSearchOnline.Location = $drawingPoint
        $lblSearchOnline.Name = "lblSearchOnline"
        $lblSearchOnline.Size = $drawingSize
        $lblSearchOnline.Text = "Search:"
        $lblSearchOnline.TabStop = $False
    #endregion

    #region txtSearchOnPremises
        $drawingPoint.X = 58
        $drawingPoint.Y = 99
        $drawingSize.Height = 20
        $drawingSize.Width = 489
        $txtSearchOnPremises.Location = $drawingPoint
        $txtSearchOnPremises.Size = $drawingSize
        $txtSearchOnPremises.Add_TextChanged({
            $itemFound = $onPremisesListView.FindItemWithText($txtSearchOnPremises.Text)

            If ($itemFound -ne $null) {
                $onPremisesListView.TopItem = $itemFound
            }
        })
    #endregion
    
    #region txtSearchOnline
        $drawingPoint.X = 669
        $drawingPoint.Y = 99
        $drawingSize.Height = 20
        $drawingSize.Width = 489
        $txtSearchOnline.Location = $drawingPoint
        $txtSearchOnline.Size = $drawingSize
        $txtSearchOnline.Add_TextChanged({
            $itemFound = $onlineListView.FindItemWithText($txtSearchOnline.Text)

            If ($itemFound -ne $null) {
                $onlineListView.TopItem = $itemFound
            }
        })
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
        $mainStatusStrip.Items.AddRange(@($progressBar, $statusLabel))
        $mainStatusStrip.TabStop = $False
    #endregion

    #region columnOnPremises1
        $columnOnPremises1.Tag = "columnOnPremises1"
        $columnOnPremises1.Name = "columnOnPremises1"
        $columnOnPremises1.Text = "Display name"
        $columnOnPremises1.Width = 265
    #endregion

    #region columnOnPremises2
        $columnOnPremises2.Tag = "columnOnPremises1"
        $columnOnPremises2.Name = "columnOnPremises1"
        $columnOnPremises2.Text = "E-mail address"
        $columnOnPremises2.Width = 249
    #endregion

    #region columnOnline1
        $columnOnline1.Tag = "columnOnline1"
        $columnOnline1.Name = "columnOnline1"
        $columnOnline1.Text = "Display name"
        $columnOnline1.Width = 265
    #endregion

    #region columnOnline2
        $columnOnline2.Tag = "columnOnline2"
        $columnOnline2.Name = "columnOnline2"
        $columnOnline2.Text = "E-mail address"
        $columnOnline2.Width = 249
    #endregion

    #region onPremisesListView
        $drawingPoint.X = 12
        $drawingPoint.Y = 125
        $drawingSize.Height = 386
        $drawingSize.Width = 535

        $filePath = "$folderPath\Background.png"
        $onPremisesListView.BackgroundImage = [System.Drawing.Image]::FromFile($filePath)
        $onPremisesListView.CheckBoxes = $True
        $onPremisesListView.Columns.AddRange(@($columnOnPremises1, $columnOnPremises2))
        $onPremisesListView.HideSelection = $False
        $onPremisesListView.Location = $drawingPoint
        $onPremisesListView.Size = $drawingSize
        $onPremisesListView.Name = "onPremisesListView"
        $onPremisesListView.ListViewItemSorter = $localSorter
        $onPremisesListView.Sorting = [System.Windows.Forms.SortOrder]::Ascending
        $onPremisesListView.View = [System.Windows.Forms.View]::Details
        $onPremisesListView.FullRowSelect = $True
        $onPremisesListView.Add_ColumnClick({
            Param($sender, $e)
            If ($onPremisesListView.AccessibleName -eq $e.Column.ToString()) {
                If ($onPremisesListView.Sorting -eq [System.Windows.Forms.SortOrder]::Ascending) {
                    $onPremisesListView.Sorting = [System.Windows.Forms.SortOrder]::Descending
                }
                Else {
                    $onPremisesListView.Sorting = [System.Windows.Forms.SortOrder]::Ascending
                }
            }
            Else {
                $onPremisesListView.Sorting = [System.Windows.Forms.SortOrder]::Ascending
                $onPremisesListView.AccessibleName = $e.Column.ToString()
            }
            $onPremisesListView.ListViewItemSorter.Order = $onPremisesListView.Sorting
            $onPremisesListView.ListViewItemSorter.SortColumn = $e.Column
            $onPremisesListView.Sort()
        })
    #endregion

    #region onlineListView
        $drawingPoint.X = 626
        $drawingPoint.Y = 125
        $drawingSize.Height = 386
        $drawingSize.Width = 535

        $onlineListView.CheckBoxes = $True
        $onlineListView.Columns.AddRange(@($columnOnline1, $columnOnline2))
        $onlineListView.HideSelection = $False
        $onlineListView.Location = $drawingPoint
        $onlineListView.Size = $drawingSize
        $onlineListView.Name = "onPremisesListView"
        $onlineListView.ListViewItemSorter = $onlineSorter
        $onlineListView.Sorting = [System.Windows.Forms.SortOrder]::Ascending
        $onlineListView.View = [System.Windows.Forms.View]::Details
        $onlineListView.FullRowSelect = $True
        $onlineListView.Add_ColumnClick({
            Param($sender, $e)
            If ($onlineListView.AccessibleName -eq $e.Column.ToString()) {
                If ($onlineListView.Sorting -eq [System.Windows.Forms.SortOrder]::Ascending) {
                    $onlineListView.Sorting = [System.Windows.Forms.SortOrder]::Descending
                }
                Else {
                    $onlineListView.Sorting = [System.Windows.Forms.SortOrder]::Ascending
                }
            }
            Else {
                $onlineListView.Sorting = [System.Windows.Forms.SortOrder]::Ascending
                $onlineListView.AccessibleName = $e.Column.ToString()
            }
            $onlineListView.ListViewItemSorter.Order = $onlineListView.Sorting
            $onlineListView.ListViewItemSorter.SortColumn = $e.Column
            $onlineListView.Sort()
        })
    #endregion

    #region btnAdd
        $drawingPoint.X = 563
        $drawingPoint.Y = 211
        $drawingSize.Height = 23
        $drawingSize.Width = 48
        $btnAdd.Location = $drawingPoint
        $btnAdd.Name = "btnAdd"
        $btnAdd.Size = $drawingSize
        $btnAdd.Text = ">"
        $btnAdd.UseVisualStyleBackColor = $True
        $btnAdd.Add_Click({
            if ($onPremisesListView.Items.Count-gt 0) {
                ($onPremisesListView.Items.Count - 1)..0 | ForEach-Object {
                    if ($onPremisesListView.Items[$_].Checked) {
                        $node = $onPremisesListView.Items[$_]
                        $onPremisesListView.Items[$_].Remove()
                        $onlineListView.Items.Add($node)
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
        $btnAddAll.Text = ">>"
        $btnAddAll.UseVisualStyleBackColor = $True
        $btnAddAll.Add_Click({
            if ($onPremisesListView.Items.Count -gt 0) {
                ($onPremisesListView.Items.Count - 1)..0 | ForEach-Object {
                    $node = $onPremisesListView.Items[$_]
                    $onPremisesListView.Items[$_].Remove()
                    $onlineListView.Items.Add($node)
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
        $btnRemove.Text = "<"
        $btnRemove.UseVisualStyleBackColor = $True
        $btnRemove.Add_Click({
            if ($onlineListView.Items.Count -gt 0) {
                ($onlineListView.Items.Count - 1)..0 | ForEach-Object {
                    if ($onlineListView.Items[$_].Checked) {
                        $node = $onlineListView.Items[$_]
                        $onlineListView.Items[$_].Remove()
                        $onPremisesListView.Items.Add($node)
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
        $btnRemoveAll.Text = "<<"
        $btnRemoveAll.UseVisualStyleBackColor = $True
        $btnRemoveAll.Add_Click({fnRemoveAll})
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
        $menuItemFileReload.Add_Click({fnLoad -LoadUsers})
    #endregion

    #region menuItemFileImport
        $drawingSize.Height = 22
        $drawingSize.Width = 152
        $menuItemFileImport.Name = "menuItemFileImport"
        $menuItemFileImport.Size = $drawingSize
        $menuItemFileImport.Text = "&Import..."
        $menuItemFileImport.Add_Click({fnImport})
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
        $menuItemFileMigrate.Add_Click({fnMigrate})
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
        $menuItemFile.DropDownItems.AddRange(@($menuItemFileConfigure, $menuItemFileConnect, $menuItemFileReload, $menuItemFileImport, $menuItemFileSpace1, $menuItemFilePreFlight, $menuItemFileExport, $menuItemFileMigrate, $menuItemFileSpace2, $menuItemFileExit))
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
        $mainMenuStrip.Text = "mainMenuStrip"
        $mainMenuStrip.TabStop = $False
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

    #region toolbarBtnImport
        $filePath = "$folderPath\Import48.png"
        $drawingSize.Height = 52
        $drawingSize.Width = 52
        $toolbarBtnImport.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Image
        $toolbarBtnImport.Image = [System.Drawing.Image]::Fromfile($filePath)
        $toolbarBtnImport.ImageTransparentColor = [System.Drawing.Color]::Magenta
        $toolbarBtnImport.Name = "toolbarBtnImport"
        $toolbarBtnImport.Size = $drawingSize
        $toolbarBtnImport.Text = "Import"
        $toolbarBtnImport.Add_Click({fnImport})

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
        $toolbarBtnReload.Add_Click({fnLoad -LoadUsers})

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
        $toolbarBtnMigrate.Add_Click({fnMigrate})
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
        $toolBar.Items.AddRange(@($toolbarBtnConfiguration, $toolbarBtnConnect, $toolbarBtnReload, $toolbarBtnImport, $toolbarSeparator1, $toolbarBtnPreFlight, $toolbarBtnExport, $toolbarBtnMigrate))
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
    $frmMain.Controls.Add($txtSearchOnPremises)
    $frmMain.Controls.Add($txtSearchOnline)
    $frmMain.Controls.Add($onPremisesListView)
    $frmMain.Controls.Add($btnAdd)
    $frmMain.Controls.Add($btnAddAll)
    $frmMain.Controls.Add($btnRemove)
    $frmMain.Controls.Add($btnRemoveAll)
    $frmMain.Controls.Add($onlineListView)
    $frmMain.Controls.Add($mainStatusStrip)
    $frmMain.Controls.Add($mainMenuStrip)
    $frmMain.Controls.Add($lblAvailable)
    $frmMain.Controls.Add($lblSelected)
    $frmMain.Controls.Add($lblSearchOnPremises)
    $frmMain.Controls.Add($lblSearchOnline)
    $frmMain.WindowState = $windowState
    [void] $frmMain.ShowDialog()
#endregion
