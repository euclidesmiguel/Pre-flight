#requires -PSEdition "Desktop"
[CmdletBinding()]
param ([PSCredential] $Credential = (New-Object System.Management.Automation.PSCredential ("dummy", (ConvertTo-SecureString "dummy" -AsPlainText -Force))))

[string] $version = "0.9"

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
#endregion

#region custom types
Add-Type -Language VisualBasic -ReferencedAssemblies ("System.Collections", "System.Windows.Forms") -TypeDefinition @"
    Imports System.Collections
    Imports System.Windows.Forms

    Public Class MigrationDetail
        Public ArchiveSizeMB As Integer
        Public CompletionTimestamp As String
        Public ErrorSummary As String
        Public ItemsTransferred As Integer
        Public MailboxEmailAddress As String
        Public MailboxSizeMB As Integer
        Public MBTransferred As Integer
        Public OverallDuration As System.TimeSpan
        Public RecipientTypeDetails As String
        Public StartTimestamp As String
        Public Status As String
        Public StatusDetail As String

        Public ReadOnly Property TotalMailboxSizeMB() As Integer
            Get
                Return Me.ArchiveSizeMB + Me.MailboxSizeMB
            End Get
        End Property

        Public Sub New()
            Me.ArchiveSizeMB = 0
            Me.CompletionTimestamp = ""
            Me.ErrorSummary = ""
            Me.ItemsTransferred = 0
            Me.MailboxEmailAddress = ""
            Me.MailboxSizeMB = 0
            Me.MBTransferred = 0
            Me.OverallDuration = System.TimeSpan.Zero
            Me.RecipientTypeDetails = ""
            Me.StartTimestamp = ""
            Me.Status = ""
            Me.StatusDetail = ""
        End Sub
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
[Boolean] $Global:cloudCredentialChanged = $False
[Boolean] $Global:isConnected = $False
[PSCredential] $Global:cloudCred = $Credential
[System.Drawing.Size] $drawingSize = New-Object -TypeName System.Drawing.Size
[System.Drawing.Point] $drawingPoint = New-Object -TypeName System.Drawing.Point
[System.Windows.Forms.FormWindowState] $windowState = New-Object System.Windows.Forms.FormWindowState
#endregion

#region fnConnect
Function fnConnect {
    [Boolean] $continue = $True

    $progressBar.Value = 10
    $progressBar.Visible = $True
    $statusLabel.Text = "Connecting to Exchange Online..."

    $cloudSession = Get-PSSession | Where-Object { (($_.ComputerName -eq "outlook.office365.com") -or ($_.ComputerName -eq "ps.outlook.com")) -and ($_.ConfigurationName -eq "Microsoft.Exchange") }
    if ($CloudSession) {
        Write-Host "Already connected to Exchange Online" -ForegroundColor Cyan
        $Global:isConnected = [Boolean] ($CloudSession)
    }
    else {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        if ($Global:cloudCred.UserName -eq "dummy") {
            Write-Host "Updating Username to connect..." -ForegroundColor Cyan
            $result = fnConfigure
            $continue = ($result -eq [System.Windows.Forms.DialogResult]::OK)
            Write-Host "Username updated successfully..." -ForegroundColor Green
        }
        if ($continue) {
            Write-Host "Checking if EXO module is installed..." -ForegroundColor Yellow
            if ( -not(Get-Module ExchangeOnlineManagement -ListAvailable) -and -not(Get-Module ExchangeOnlineManagement) ) {
                Write-Host "Installing EXO module..." -ForegroundColor Yellow
                Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
            }
            Write-Host "EXO module is installed..." -ForegroundColor Yellow
            Import-Module ExchangeOnlineManagement
            Write-Host "Attempting to connect using Modern Auth..." -ForegroundColor Yellow
            Connect-ExchangeOnline -UserPrincipalName $Global:cloudCred.UserName -ShowBanner:$false -Verbose
            Write-Host "Successfully connected using Modern Auth..." -ForegroundColor Yellow

            $cloudSession = Get-PSSession | Where-Object { ($_.ComputerName -eq "outlook.office365.com") -and ($_.ConfigurationName -eq "Microsoft.Exchange") }
            $Global:isConnected = [Boolean] ($CloudSession)
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

#region fnDisconnect
Function fnDisconnect {
    $cloudSession = Get-PSSession | Where-Object { ($_.ComputerName -eq "ps.outlook.com") -and ($_.ConfigurationName -eq "Microsoft.Exchange") }
    if ($cloudSession) {
        Disconnect-ExchangeOnline -Confirm:$false
    }
}
#endregion

#region fnLoad
Function fnLoad {
    $listMigrationInProgress.Items.Clear()

    if (-not $Global:isConnected) {
        Write-Host "You are not connected to Exchange yet." -ForegroundColor Red
    }
    else {
        $progressBar.Value = 35
        $progressBar.Visible = $True
        $statusLabel.Text = "Loading migration batches..."

        Get-MigrationBatch | ForEach-Object {
            $listViewItem = New-Object -TypeName System.Windows.Forms.ListViewItem([System.String[]](@($_.Identity.Name, $_.Status, $_.TotalCount, $_.Identity.Id, $_.StartDateTimeUTC.ToString(), $_.ActiveCount, $_.StoppedCount, $_.FinalizedCount, $_.FailedCount, $_.PendingCount, $_.SyncedCount)), -1)
            $ListViewItem.Checked = $True
            $listMigrationInProgress.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($listViewItem)))
        }
    }
    $listMigrationInProgress.Sort()
    $progressBar.Visible = $False
    $statusLabel.Text = ""
}
#endregion

#region fnConfigure
Function fnConfigure {

    $Global:cloudCredentialChanged = $False
    [System.Windows.Forms.Form] $frmConfig = New-Object -TypeName System.Windows.Forms.Form
    [System.Windows.Forms.Button] $btnOk = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.Button] $btnCancel = New-Object -TypeName System.Windows.Forms.Button
    [System.Windows.Forms.GroupBox] $grpOnline = New-Object -TypeName System.Windows.Forms.GroupBox
    [System.Windows.Forms.Label] $lblCloudPassword = New-Object -TypeName System.Windows.Forms.Label
    [System.Windows.Forms.Label] $lblCloudUser = New-Object -TypeName System.Windows.Forms.Label
    [System.Windows.Forms.TextBox] $txtCloudPassword = New-Object -TypeName System.Windows.Forms.TextBox
    [System.Windows.Forms.TextBox] $txtCloudUser = New-Object -TypeName System.Windows.Forms.TextBox

    #region txtCloudUser
    $drawingPoint.X = 151
    $drawingPoint.Y = 20
    $drawingSize.Height = 20
    $drawingSize.Width = 179
    $txtCloudUser.Location = $drawingPoint
    $txtCloudUser.Size = $drawingSize
    $txtCloudUser.Add_TextChanged({
            $Global:cloudCredentialChanged = $True
            if (($txtcloudUser.Text -ne "") -and ($txtcloudPassword.Text -ne "")) { $btnOk.Enabled = $True }
            else { $btnOk.Enabled = $False }
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
            if (($txtcloudUser.Text -ne "") -and ($txtcloudPassword.Text -ne "")) { $btnOk.Enabled = $True }
            else { $btnOk.Enabled = $False }
        })
    if ((Get-Command Connect-EXOPSSession -ErrorAction SilentlyContinue).Count -gt 0) {
        $txtCloudPassword.Enabled = $False
        $txtCloudPassword.Visible = $False
        $txtCloudPassword.Text = "dummy"
    }
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
    $drawingPoint.X = 191
    $drawingPoint.Y = 103
    $drawingSize.Height = 23
    $drawingSize.Width = 75
    $btnOk.Location = $drawingPoint
    $btnOk.Name = "btnOk"
    $btnOk.Size = $drawingSize
    $btnOk.Text = "Ok"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOk.Enabled = $False
    $btnOk.Add_Click({
            if ($Global:cloudCredentialChanged -and ($txtcloudUser.Text -ne "") -and ($txtcloudPassword.Text -ne "")) { $Global:cloudCred = New-Object System.Management.Automation.PSCredential ($txtcloudUser.Text, (ConvertTo-SecureString $txtcloudPassword.Text -AsPlainText -Force)) }
            if ($Global:cloudCred.UserName -ne "dummy") { $Global:configurationFinished = $True }
            $Global:configurationFinished = $True
            $frmConfig.Close()
        })
    #endregion

    #region btnCancel
    $drawingPoint.X = 273
    $drawingPoint.Y = 103
    $drawingSize.Height = 23
    $drawingSize.Width = 75
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.Location = $drawingPoint
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = $drawingSize
    $btnCancel.Text = "Cancel"
    #endregion

    #region grpOnline
    $drawingPoint.X = 12
    $drawingPoint.Y = 12
    $drawingSize.Height = 85
    $drawingSize.Width = 336
    $grpOnline.Controls.Add($lblCloudUser)
    $grpOnline.Controls.Add($txtCloudUser)
    $grpOnline.Controls.Add($lblCloudPassword)
    $grpOnline.Controls.Add($txtCloudPassword)
    $grpOnline.Location = $drawingPoint
    $grpOnline.Name = "grpOnline"
    $grpOnline.Size = $drawingSize
    $grpOnline.Text = "Exchange Online"
    #endregion

    #region frmConfig
    $drawingSize.Height = 141
    $drawingSize.Width = 360
    $frmConfig.ClientSize = $drawingSize
    $frmConfig.ControlBox = $False
    $frmConfig.Controls.Add($grpOnline)
    $frmConfig.Controls.Add($btnOk)
    $frmConfig.Controls.Add($btnCancel)
    $frmConfig.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $frmConfig.MaximizeBox = $False
    $frmConfig.MinimizeBox = $False
    $frmConfig.Name = "frmConfig"
    $frmConfig.Text = "Configuration"
    $frmConfig.ShowIcon = $False
    $frmConfig.ShowInTaskbar = $False
    $frmConfig.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $frmConfig.AcceptButton = $btnOk
    $frmConfig.CancelButton = $btnCancel
    $frmConfig.Add_Closed({ $frmConfig = $null })
    $frmConfig.Add_Load({
            If ($Global:configurationFinished) {
                $txtCloudUser.Text = $Global:cloudCred.UserName
                $txtCloudPassword.Text = "***************"
            }
        })
    #endregion

    $grpOnline.ResumeLayout($False)
    $grpOnline.PerformLayout()
    $frmConfig.ResumeLayout($False)
            
    $frmConfig.WindowState = $windowState

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
    [System.Drawing.Font] $formFont = New-Object -TypeName System.Drawing.Font("Century Gothic", 28, [System.Drawing.FontStyle]::Regular)

    #region lblTitle
    $drawingPoint.X = 0
    $drawingPoint.Y = 18
    $drawingSize.Height = 44
    $drawingSize.Width = 436
    $lblTitle.AutoSize = $True
    $lblTitle.Location = $drawingPoint
    $lblTitle.Name = "lblTitle"
    $lblTitle.Size = $drawingSize
    $lblTitle.Text = "Migration reporting tool"
    $lblTitle.Font = $formFont
    #endregion

    #region panelAbout
    $drawingPoint.X = 0
    $drawingPoint.Y = 0
    $drawingSize.Height = 84
    $drawingSize.Width = 438
    $panelAbout.Location = $drawingPoint
    $panelAbout.Size = $drawingSize
    $panelAbout.BackColor = [System.Drawing.Color]::White
    $panelAbout.Controls.Add($lblTitle)
    $panelAbout.Name = "panelAbout"
    #endregion

    #region lblName
    $drawingPoint.X = 15
    $drawingPoint.Y = 95
    $drawingSize.Height = 13
    $drawingSize.Width = 106
    $lblName.AutoSize = $True
    $lblName.Location = $drawingPoint
    $lblName.Name = "lblName"
    $lblName.Text = "Pre-flight tool"
    #endregion

    #region lblVersion
    $drawingPoint.X = 15
    $drawingPoint.Y = 113
    $drawingSize.Height = 13
    $drawingSize.Width = 60
    $lblVersion.AutoSize = $True
    $lblVersion.Location = $drawingPoint
    $lblVersion.Name = "lblVersion"
    $lblVersion.Text = "Version $version"
    #endregion

    #region lblDisclaimer
    $drawingPoint.X = 17
    $drawingPoint.Y = 153
    $drawingSize.Height = 163
    $drawingSize.Width = 397
    $lblDisclaimer.AutoSize = $False
    $lblDisclaimer.Location = $drawingPoint
    $lblDisclaimer.Size = $drawingSize
    $lblDisclaimer.Name = "lblDisclaimer"
    $lblDisclaimer.Text = "Version $version"
    $lblDisclaimer.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $lblDisclaimer.Text = "The Migration reporting tool is a sample script. The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages."
    #endregion

    #region lnkBlog
    $drawingPoint.X = 15
    $drawingPoint.Y = 133
    $drawingSize.Height = 13
    $drawingSize.Width = 102
    $lnkBlog.AutoSize = $True
    $lnkBlog.Location = $drawingPoint
    $lnkBlog.Name = "lnkBlog"
    $lnkBlog.Text = "FastTrack Tips Blog"
    $lnkBlog.Add_Click({ Start-Process "http://aka.ms/ftctips" })
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
    $btnAboutOk.Add_Click({ $frmAbout.Close() })
    #endregion

    #region frmAbout
    $drawingSize.Height = 356
    $drawingSize.Width = 440
    $frmAbout.AcceptButton = $btnAboutOk
    $frmAbout.CancelButton = $btnAboutOk
    $frmAbout.ClientSize = $drawingSize
    $frmAbout.ControlBox = $False
    $frmAbout.Controls.Add($panelAbout)
    $frmAbout.Controls.Add($lblName)
    $frmAbout.Controls.Add($lblVersion)
    $frmAbout.Controls.Add($lnkBlog)
    $frmAbout.Controls.Add($lblDisclaimer)
    $frmAbout.Controls.Add($btnAboutOk)
    $frmAbout.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $frmAbout.MaximizeBox = $False
    $frmAbout.MinimizeBox = $False
    $frmAbout.Name = "frmAbout"
    $frmAbout.Text = "About Migration reporting"
    $frmAbout.ShowIcon = $False
    $frmAbout.ShowInTaskbar = $False
    $frmAbout.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $frmAbout.Add_Closed({ $frmAbout = $null })
    #endregion

    $panelAbout.ResumeLayout($False)
    $panelAbout.PerformLayout()
    $frmAbout.ResumeLayout($False)
    $frmAbout.PerformLayout()
    $frmAbout.WindowState = $windowState
    [void] $frmAbout.ShowDialog()
}
#endregion

#region fnRunReport
Function fnRunReport {
    [System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.DialogResult]::OK
    [System.Windows.Forms.SaveFileDialog] $saveDialog = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    [string] $badCharsRE = "[ $([System.IO.Path]::GetInvalidFileNameChars() -join '')]"
    [string] $htmlHeader = ""
    [string] $htmlBody = ""
    [string] $reportFileName = "$(Get-Date -Format "yyyyMMdd-HHmmss")"
    [string] $detailFolderName = ""
    [string] $detailFileName = ""
    [int] $totalMailboxes = 0
    [int] $currentMailbox = 0
    [MigrationDetail] $itemDetail = New-Object -TypeName MigrationDetail

    if ($listMigrationInProgress.CheckedItems.Count -eq 1) {
        $htmlHeader = "<title>Migration report for $($listMigrationInProgress.CheckedItems.Item(0).SubItems.Item(0).Text)</title>"
    }
    elseif ($listMigrationInProgress.CheckedItems.Count -gt 1) {
        $htmlHeader = "<title>Migration report for $($listMigrationInProgress.CheckedItems.Count) migration batches</title>"
    }
    else { return }

    $saveDialog.InitialDirectory = [System.IO.Path]::GetFullPath("$($Script:MyInvocation.MyCommand.Path)\..\Reports")
    $saveDialog.Filter = "HTML file (*.html, *.htm)|*.html;*.htm|All files (*.*)|*.*"
    $saveDialog.FileName = "$reportFileName.html"
    $result = $saveDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $statusLabel.Text = "Generating general migration report $([System.IO.Path]::GetFileName($reportFileName))"
        $progressBar.Value = 0
        $progressBar.Visible = $True
        Write-Progress -Activity "Generating migration report" -Status $statusLabel.Text -PercentComplete $progressBar.Value

        $reportFileName = $saveDialog.FileName
        $detailFolderName = ([System.IO.Path]::GetDirectoryName($reportFileName), "\", [System.IO.Path]::GetFileNameWithoutExtension($reportFileName)) -join ""
            
        $htmlHeader += "`r`n<style>"
        $htmlHeader += "body {font-family: Calibri, Arial;}"
        $htmlHeader += "table {border-width: 1px;border-collapse: collapse;}"
        $htmlHeader += "td {padding: 3px}"
        $htmlHeader += "</style>"

        $CheckedItems = $listMigrationInProgress.CheckedItems
        $CheckedItems | ForEach-Object {
            $htmlBody += "<table>`r`n"
            $htmlBody += "  <tr><th colspan=2 style=""float:left""><h1>$($_.SubItems.Item(0).Text)</h1></th></tr>`r`n"
            $htmlBody += "  <tr><td><b>Batch name:</b></td><td><a href="".\$([System.IO.Path]::GetFileName($detailFolderName))\$([System.IO.Path]::GetFileName(($_.SubItems.Item(0).Text -replace $badCharsRE))).csv"">$($_.SubItems.Item(0).Text)</a></td></tr>`r`n"
            $htmlBody += "  <tr><td><b>Status:</b></td><td>$($_.SubItems.Item(1).Text)</td></tr>`r`n"
            $htmlBody += "  <tr><td><b>Migration time:</b></td><td>$($_.SubItems.Item(4).Text)</td></tr>`r`n"
            $htmlBody += "  <tr><td><b>Total mailboxes scheduled for migration:</b></td><td>$($_.SubItems.Item(2).Text)</td></tr>`r`n"
            $htmlBody += "  <tr><td><b>Migration in progress:</b></td><td>$($_.SubItems.Item(5).Text)</td></tr>`r`n"
            $htmlBody += "  <tr><td><b>Migration stopped:</b></td><td>$($_.SubItems.Item(6).Text)</td></tr>`r`n"
            $htmlBody += "  <tr><td><b>Migration finished:</b></td><td>$($_.SubItems.Item(7).Text)</td></tr>`r`n"
            $htmlBody += "  <tr><td><b>Migration failed:</b></td><td>$($_.SubItems.Item(8).Text)</td></tr>`r`n"
            $htmlBody += "  <tr><td><b>Migration pending:</b></td><td>$($_.SubItems.Item(9).Text)</td></tr>`r`n"
            $htmlBody += "  <tr><td><b>Synced mailboxes:</b></td><td>$($_.SubItems.Item(10).Text)</td></tr>`r`n"
            $htmlBody += "</table>`r`n"
            $htmlBody += "<br>`r`n"

            $totalMailboxes += ([int]$($_.SubItems.Item(2).Text))
        }
        $(ConvertTo-Html -Head $htmlHeader -Body $htmlBody) | Out-File -FilePath $reportFileName -Force

        $CheckedItems | ForEach-Object {
            $detailFileName = ($detailFolderName, "\", ($_.SubItems.Item(0).Text -replace $badCharsRE), ".csv") -join ""
            mkdir $detailFolderName -Force -ErrorAction SilentlyContinue
            Remove-Item -Path $detailFileName -Force -ErrorAction SilentlyContinue

            $itemDetail = New-Object -TypeName MigrationDetail

            $MigrationUsers = Get-MigrationUser -BatchId $($_.SubItems.Item(3).Text) 
            $MigrationUsers | ForEach-Object {
                $statusLabel.Text = "Getting migration status for $($_.MailboxEmailAddress)..."
                $progressBar.Value = ($currentMailbox / $totalMailboxes) * 100
                Write-Progress -Activity "Generating migration report" -Status $statusLabel.Text -PercentComplete $progressBar.Value
                $currentMailbox++

                $itemDetail.MailboxEmailAddress = $_.MailboxEmailAddress
                $itemDetail.Status = $_.Status
                $itemDetail.ErrorSummary = $_.ErrorSummary

                $MoveRequestStatistics = $_ | Get-MoveRequestStatistics -ErrorAction SilentlyContinue
                $MoveRequestStatistics | ForEach-Object {
                    try { $itemDetail.ArchiveSizeMB = [math]::Round([int]($_.TotalArchiveSize.ToString().Split("(")[1] -replace "[ bytes),\,]", "") / 1MB) }catch {}
                    try { $itemDetail.CompletionTimestamp = $_.CompletionTimestamp.ToString() }catch {}
                    try { $itemDetail.ItemsTransferred = $_.ItemsTransferred }catch {}
                    try { $itemDetail.MailboxSizeMB = [math]::Round(($_.TotalMailboxSize.ToString().Split("(")[1] -replace "[ bytes),\,]", "") / 1MB) }catch {}
                    try { $itemDetail.MBTransferred = [math]::Round(($_.BytesTransferred.ToString().Split("(")[1] -replace "[ bytes),\,]", "") / 1MB) }catch {}
                    try { $itemDetail.OverallDuration = $_.OverallDuration.TimeSpan }catch {}
                    try { $itemDetail.RecipientTypeDetails = $_.RecipientTypeDetails }catch {}
                    try { $itemDetail.RecipientTypeDetails = $_.RecipientTypeDetails }catch {}
                    try { $itemDetail.StatusDetail = $_.StatusDetail }catch {}
                }
                $itemDetail | Export-Csv -Path $detailFileName -Append -NoTypeInformation
            }
        }
        $statusLabel.Text = ""
        $progressBar.Value = 0
        $progressBar.Visible = $False
        Write-Progress -Activity "Generating migration report" -Completed
    }
}
#endregion

#region Declaring form objects
[string] $folderPath = "$(Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)\Images"
[string] $filePath = ""
[System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.DialogResult]::OK
[System.Windows.Forms.Form] $frmMain = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.StatusStrip] $mainStatusStrip = New-Object -TypeName System.Windows.Forms.StatusStrip
[System.Windows.Forms.ToolStripProgressBar] $progressBar = New-Object -TypeName System.Windows.Forms.ToolStripProgressBar
[System.Windows.Forms.ToolStripStatusLabel] $statusLabel = New-Object -TypeName System.Windows.Forms.ToolStripStatusLabel
[System.Windows.Forms.MenuStrip] $mainMenuStrip = New-Object -TypeName System.Windows.Forms.MenuStrip
[System.Windows.Forms.ToolStripMenuItem] $menuItemFile = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
[System.Windows.Forms.ToolStripMenuItem] $menuItemFileConfigure = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
[System.Windows.Forms.ToolStripMenuItem] $menuItemFileConnect = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
[System.Windows.Forms.ToolStripMenuItem] $menuItemFileReload = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
[System.Windows.Forms.ToolStripMenuItem] $menuItemFileReport = New-Object -TypeName System.Windows.Forms.ToolStripMenuItem
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
[System.Windows.Forms.ToolStripButton] $toolbarBtnReport = New-Object -TypeName System.Windows.Forms.ToolStripButton
[System.Windows.Forms.ToolStripSeparator] $toolbarSeparator = New-Object -TypeName System.Windows.Forms.ToolStripSeparator
[System.Windows.Forms.ListView] $listMigrationInProgress = New-Object -TypeName System.Windows.Forms.ListView
[System.Windows.Forms.ColumnHeader] $column0 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column1 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column2 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column3 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column4 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column5 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column6 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column7 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column8 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column9 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[System.Windows.Forms.ColumnHeader] $column10 = New-Object -TypeName System.Windows.Forms.ColumnHeader
[ListViewColumnSorter] $sorter = New-Object -TypeName ListViewColumnSorter
#endregion

#region setting form objects
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
$drawingSize.Width = 591
$mainStatusStrip.Location = $drawingPoint
$mainStatusStrip.Name = "mainStatusStrip"
$mainStatusStrip.Size = $drawingSize
$mainStatusStrip.SizingGrip = $False
$mainStatusStrip.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$mainStatusStrip.Items.AddRange(@($progressBar, $statusLabel))
$mainStatusStrip.ResumeLayout($False)
$mainStatusStrip.PerformLayout()
#endregion

#region column0
$column0.Tag = "column0"
$column0.Name = "column0"
$column0.Text = "Batch"
$column0.Width = 400
#endregion

#region column1
$column1.Tag = "column1"
$column1.Name = "column1"
$column1.Text = "Status"
$column1.Width = 85
#endregion

#region column2
$column2.Tag = "column2"
$column2.Name = "column2"
$column2.Text = "Mailboxes"
$column2.Width = 85
#endregion

#region column3
$column3.Tag = "column3"
$column3.Name = "column3"
$column3.Text = "BatchGuid"
$column3.Width = 0
#endregion

#region column4
$column4.Tag = "column4"
$column4.Name = "column4"
$column4.Text = "StartDateTimeUTC"
$column4.Width = 0
#endregion

#region column5
$column5.Tag = "column5"
$column5.Name = "column5"
$column5.Text = "ActiveCount"
$column5.Width = 0
#endregion

#region column6
$column6.Tag = "column6"
$column6.Name = "column6"
$column6.Text = "StoppedCount"
$column6.Width = 0
#endregion

#region column7
$column7.Tag = "column7"
$column7.Name = "column7"
$column7.Text = "FinalizedCount"
$column7.Width = 0
#endregion

#region column8
$column8.Tag = "column8"
$column8.Name = "column8"
$column8.Text = "FailedCount"
$column8.Width = 0
#endregion

#region column9
$column9.Tag = "column9"
$column9.Name = "column9"
$column9.Text = "PendingCount"
$column9.Width = 0
#endregion

#region column10
$column10.Tag = "column10"
$column10.Name = "column10"
$column10.Text = "SyncedCount"
$column10.Width = 0
#endregion

#region listMigrationInProgress
$drawingPoint.X = 0
$drawingPoint.Y = 79
$drawingSize.Height = 445
$drawingSize.Width = 591
$listMigrationInProgress.CheckBoxes = $True
$listMigrationInProgress.Columns.AddRange(@($column0, $column1, $column2, $column3, $column4, $column5, $column6, $column7, $column8, $column9, $column10))
$listMigrationInProgress.HideSelection = $False
$listMigrationInProgress.Location = $drawingPoint
$listMigrationInProgress.Size = $drawingSize
$listMigrationInProgress.Name = "listMigrationInProgress"
$listMigrationInProgress.ListViewItemSorter = $sorter
$listMigrationInProgress.Sorting = [System.Windows.Forms.SortOrder]::Ascending
$listMigrationInProgress.View = [System.Windows.Forms.View]::Details
$listMigrationInProgress.FullRowSelect = $True
$listMigrationInProgress.Add_ColumnClick({
        Param($sender, $e)
        If ($listMigrationInProgress.AccessibleName -eq $e.Column.ToString()) {
            If ($listMigrationInProgress.Sorting -eq [System.Windows.Forms.SortOrder]::Ascending) {
                $listMigrationInProgress.Sorting = [System.Windows.Forms.SortOrder]::Descending
            }
            Else {
                $listMigrationInProgress.Sorting = [System.Windows.Forms.SortOrder]::Ascending
            }
        }
        Else {
            $listMigrationInProgress.Sorting = [System.Windows.Forms.SortOrder]::Ascending
            $listMigrationInProgress.AccessibleName = $e.Column.ToString()
        }
        $listMigrationInProgress.ListViewItemSorter.Order = $listMigrationInProgress.Sorting
        $listMigrationInProgress.ListViewItemSorter.SortColumn = $e.Column
        $listMigrationInProgress.Sort()
    })
#endregion

#region menuItemFileConfigure
$drawingSize.Height = 22
$drawingSize.Width = 152
$menuItemFileConfigure.Name = "menuItemFileConfigure"
$menuItemFileConfigure.Size = $drawingSize
$menuItemFileConfigure.Text = "C&onfigure..."
$menuItemFileConfigure.Add_Click({ $result = fnConfigure })
#endregion

#region menuItemFileConnect
$drawingSize.Height = 22
$drawingSize.Width = 152
$menuItemFileConnect.Name = "menuItemFileConnect"
$menuItemFileConnect.Size = $drawingSize
$menuItemFileConnect.Text = "&Connect..."
$menuItemFileConnect.Add_Click({ fnConnect -Credential $Global:cloudCred })
#endregion

#region menuItemFileReload
$drawingSize.Height = 22
$drawingSize.Width = 152
$menuItemFileReload.Name = "menuItemFileReload"
$menuItemFileReload.Size = $drawingSize
$menuItemFileReload.Text = "&Reload"
$menuItemFileReload.Add_Click({ fnLoad })
#endregion

#region menuItemFileReport
$drawingSize.Height = 22
$drawingSize.Width = 152
$menuItemFileReport.Name = "menuItemFilePreFlight"
$menuItemFileReport.Size = $drawingSize
$menuItemFileReport.Text = "Repor&t..."
$menuItemFileReport.Add_Click({ fnRunReport })
#endregion

#region menuItemFileExit
$drawingSize.Height = 22
$drawingSize.Width = 152
$menuItemFileExit.Name = "menuItemFileExit"
$menuItemFileExit.Size = $drawingSize
$menuItemFileExit.Text = "E&xit"
$menuItemFileExit.Add_Click({ $frmMain.Close() })
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

#region menuItemFile
$drawingSize.Height = 20
$drawingSize.Width = 37
$menuItemFile.DropDownItems.AddRange(@($menuItemFileConfigure, $menuItemFileConnect, $menuItemFileReload, $menuItemFileSpace1, $menuItemFileReport, $menuItemFileSpace2, $menuItemFileExit))
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
$menuItemHelpBlog.Add_Click({ Start-Process "http://aka.ms/ftctips" })
#endregion

#region menuItemHelpAbout
$drawingSize.Height = 22
$drawingSize.Width = 207
$menuItemHelpAbout.Name = "menuItemHelpAbout"
$menuItemHelpAbout.Size = $drawingSize
$menuItemHelpAbout.Text = "&About"
$menuItemHelpAbout.Add_Click({ fnAbout })
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
$drawingSize.Width = 591
$mainMenuStrip.Items.AddRange(@($menuItemFile, $menuItemHelp))
$mainMenuStrip.Location = $drawingPoint
$mainMenuStrip.Name = "mainMenuStrip"
$mainMenuStrip.Size = $drawingSize
$mainMenuStrip.Text = "mainMenuStrip"
$mainMenuStrip.ResumeLayout($False)
$mainMenuStrip.PerformLayout()
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
$toolbarBtnConfiguration.Text = "Configure"
$toolbarBtnConfiguration.Add_Click({ $result = fnConfigure })
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
$toolbarBtnConnect.Add_Click({ fnConnect -Credential $Global:cloudCred })
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
$toolbarBtnReload.Add_Click({ fnLoad })

#endregion

#region toolbarBtnReport
$filePath = "$folderPath\Report48.png"
$drawingSize.Height = 52
$drawingSize.Width = 52
$toolbarBtnReport.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Image
$toolbarBtnReport.Image = [System.Drawing.Image]::Fromfile($filePath)
$toolbarBtnReport.ImageTransparentColor = [System.Drawing.Color]::Magenta
$toolbarBtnReport.Name = "toolbarBtnReport"
$toolbarBtnReport.Size = $drawingSize
$toolbarBtnReport.Text = "Generate report"
$toolbarBtnReport.Add_Click({ fnRunReport })
#endregion

#region toolbarSeparator
$drawingSize.Height = 55
$drawingSize.Width = 6
$toolbarSeparator.Name = "toolbarSeparator"
$toolbarSeparator.Size = $drawingSize
#endregion

#region toolBar
$drawingPoint.X = 0
$drawingPoint.Y = 24
$drawingSize.Height = 48
$drawingSize.Width = 48
$toolBar.ImageScalingSize = $drawingSize
$toolBar.Items.AddRange(@($toolbarBtnConfiguration, $toolbarBtnConnect, $toolbarBtnReload, $toolbarSeparator, $toolbarBtnReport))
$toolBar.Location = $drawingPoint
$drawingSize.Height = 55
$drawingSize.Width = 591
$toolBar.Size = $drawingSize
$toolBar.Name = "toolBar"
$toolBar.Text = "toolBar"
$toolBar.ResumeLayout($False)
$toolBar.PerformLayout()
#endregion

#region frmMain
$filePath = "$folderPath\ReportIcon.ico"
$drawingSize.Height = 585
$drawingSize.Width = 607
$frmMain.Size = $drawingSize
$frmMain.MinimumSize = $drawingSize
$frmMain.MaximumSize = $drawingSize
$drawingSize.Height = 546
$drawingSize.Width = 591
$frmMain.MainMenuStrip = $mainMenuStrip
$frmMain.ClientSize = $drawingSize
$frmMain.MaximizeBox = $False
$frmMain.Name = "frmMain"
$frmMain.Text = "Migration reporting"
$frmMain.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$frmMain.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($filePath)
$frmMain.Add_Closed({ fnDisconnect })
$frmMain.Controls.Add($toolBar)
$frmMain.Controls.Add($mainMenuStrip)
$frmMain.Controls.Add($listMigrationInProgress)
$frmMain.Controls.Add($mainStatusStrip)
$frmMain.WindowState = $windowState
$frmMain.ResumeLayout($False)
$frmMain.PerformLayout()
#endregion

#endregion

if ($Credential.UserName -ne "dummy") {
    $Global:configurationFinished = $True
    fnConnect -Credential $Credential | Out-Null
}

[void] $frmMain.ShowDialog()
