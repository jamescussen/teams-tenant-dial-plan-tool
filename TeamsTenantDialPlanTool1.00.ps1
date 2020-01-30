########################################################################
# Name: Teams Tenant Dial Plan Tool 
# Version: v1.00 (1/9/2019)
# Date: 1/9/2019
# Created By: James Cussen
# Web Site: http://www.myskypelab.com (formally http://www.mylynclab.com)
# 
# Notes: This is a PowerShell tool. To run the tool, open it from the PowerShell command line on a Lync server.
#		 For more information on the requirements for setting up and using this tool please visit http://www.myskypelab.com.
#
# Copyright: Copyright (c) 2019, James Cussen (www.myskypelab.com) All rights reserved.
# Licence: 	Redistribution and use of script, source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#				1) Redistributions of script code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				2) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				3) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#				4) This license does not include any resale or commercial use of this software.
#				5) Any portion of this software may not be reproduced, duplicated, copied, sold, resold, or otherwise exploited for any commercial purpose without express written consent of James Cussen.
#			THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; LOSS OF GOODWILL OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Release Notes:
# 1.00 Initial Release.
########################################################################

[cmdletbinding()]
Param()

$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

Write-Host ""
Write-Host "--------------------------------------------------------------"
Write-Host "PowerShell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 PowerShell installed.  This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has Version 2 PowerShell installed. This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 5 PowerShell installed. CHECK PASSED!" -foreground "green"
}
else
{
	Write-Host "This machine has version $MajorVersion PowerShell installed. Unknown level of support for this version." -foreground "yellow"
}
Write-Host "--------------------------------------------------------------"
Write-Host ""


$script:OnlineUsername = ""
if($OnlineUsernameInput -ne $null -and $OnlineUsernameInput -ne "")
{
	Write-Host "INFO: Using command line AdminPasswordInput setting = $OnlineUsernameInput" -foreground "Yellow"
	$script:OnlineUsername = $OnlineUsernameInput
}

$script:OnlinePassword = ""
if($OnlinePasswordInput -ne $null -and $OnlinePasswordInput -ne "")
{
	Write-Host "INFO: Using command line OnlinePasswordInput setting = $OnlinePasswordInput" -foreground "Yellow"
	$script:OnlinePassword = $OnlinePasswordInput
}


#Office 365 reconnect variables
$Script:O365Creds = $null
$Script:O365ReconnectAttempts = 0
$Script:UserConnectedToSfBOnline = $false


# Set up the form  ============================================================

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$mainForm = New-Object System.Windows.Forms.Form 
$mainForm.Text = "Teams Tenant Dial Plan Tool 1.00"
$mainForm.Size = New-Object System.Drawing.Size(525,680) 
$mainForm.MinimumSize = New-Object System.Drawing.Size(520,450) 
$mainForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$mainForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$mainForm.KeyPreview = $True
$mainForm.TabStop = $false


$global:SFBOsession = $null
#ConnectButton
$ConnectOnlineButton = New-Object System.Windows.Forms.Button
$ConnectOnlineButton.Location = New-Object System.Drawing.Size(20,7)
$ConnectOnlineButton.Size = New-Object System.Drawing.Size(100,20)
$ConnectOnlineButton.Text = "Connect SfBO"
$ConnectTooltip = New-Object System.Windows.Forms.ToolTip
$ConnectToolTip.SetToolTip($ConnectOnlineButton, "Connect/Disconnect from Skype for Business Online")
#$ConnectButton.tabIndex = 1
$ConnectOnlineButton.Enabled = $true
$ConnectOnlineButton.Add_Click({	

	$ConnectOnlineButton.Enabled = $false
	
	$StatusLabel.Text = "STATUS: Connecting to O365..."
	
	if($ConnectOnlineButton.Text -eq "Connect SfBO")
	{
		ConnectSkypeForBusinessOnline
		[System.Windows.Forms.Application]::DoEvents()
		CheckSkypeForBusinessOnline
	}
	elseif($ConnectOnlineButton.Text -eq "Disconnect SfBO")
	{	
		$ConnectOnlineButton.Text = "Disconnecting..."
		$StatusLabel.Text = "STATUS: Disconnecting from O365..."
		$Script:UserConnectedToSfBOnline = $false
		DisconnectSkypeForBusinessOnline
		CheckSkypeForBusinessOnline
		$Script:O365Creds = $null
	}
	
	$ConnectOnlineButton.Enabled = $true
	
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($ConnectOnlineButton)


$ConnectedOnlineLabel = New-Object System.Windows.Forms.Label
$ConnectedOnlineLabel.Location = New-Object System.Drawing.Size(125,10) 
$ConnectedOnlineLabel.Size = New-Object System.Drawing.Size(100,15) 
$ConnectedOnlineLabel.Text = "Connected"
$ConnectedOnlineLabel.TabStop = $false
$ConnectedOnlineLabel.forecolor = "green"
$ConnectedOnlineLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::left
$ConnectedOnlineLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$mainForm.Controls.Add($ConnectedOnlineLabel)
$ConnectedOnlineLabel.Visible = $false



$MyLinkLabel = New-Object System.Windows.Forms.LinkLabel
$MyLinkLabel.Location = New-Object System.Drawing.Size(380,5)
$MyLinkLabel.Size = New-Object System.Drawing.Size(120,15)
$MyLinkLabel.DisabledLinkColor = [System.Drawing.Color]::Red
$MyLinkLabel.VisitedLinkColor = [System.Drawing.Color]::Blue
$MyLinkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$MyLinkLabel.LinkColor = [System.Drawing.Color]::Navy
$MyLinkLabel.TabStop = $False
$MyLinkLabel.Text = "www.myskypelab.com"
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("http://www.myskypelab.com")
})
$mainForm.Controls.Add($MyLinkLabel)



#Policy Label ============================================================
$policyLabel = New-Object System.Windows.Forms.Label
$policyLabel.Location = New-Object System.Drawing.Size(22,43) 
$policyLabel.Size = New-Object System.Drawing.Size(58,15) 
$policyLabel.Text = "Dial Plans: "
$policyLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$policyLabel.TabStop = $false
$mainForm.Controls.Add($policyLabel)


# Add Client Policy Dropdown box ============================================================
$policyDropDownBox = New-Object System.Windows.Forms.ComboBox 
$policyDropDownBox.Location = New-Object System.Drawing.Size(80,40) 
$policyDropDownBox.Size = New-Object System.Drawing.Size(255,20) 
$policyDropDownBox.DropDownHeight = 200 
$policyDropDownBox.DropDownWidth = 300
$policyDropDownBox.tabIndex = 1
$policyDropDownBox.DropDownStyle = "DropDownList"
$policyDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$mainForm.Controls.Add($policyDropDownBox) 


$policyDropDownBox.add_SelectedValueChanged(
{
	$StatusLabel.Text = "STATUS: Getting dial plan..."
	[System.Windows.Forms.Application]::DoEvents()
	GetNormalisationPolicy
	UpdateListViewSettings
	$TestPhonePatternTextLabel.Text = "Matched Pattern:"
	$TestPhoneTranslationTextLabel.Text = "Matched Translation:"
	$TestPhoneResultTextLabel.Text = "Test Result:"
	$StatusLabel.Text = ""
})


#NewPolicy button
$NewPolicyButton = New-Object System.Windows.Forms.Button
$NewPolicyButton.Location = New-Object System.Drawing.Size(340,39)
$NewPolicyButton.Size = New-Object System.Drawing.Size(50,20)
$NewPolicyButton.Text = "New.."
$NewPolicyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$NewPolicyButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Opening new dial plan dialog..."
	[System.Windows.Forms.Application]::DoEvents()
	$Result = New-Policy -Message "Enter the new dial plan name:" -WindowTitle "New Tenant Dial Plan" -DefaultText "Dial Plan Name"
	
	if($Result -ne $null)
	{
		$ExistingDialplans = Get-CsTenantDialPlan
		$foundPolicy = $false
		foreach($ExistingDialplan in $ExistingDialplans)
		{
			if($ExistingDialplan.Identity -eq $Result.NewPolicy)
			{
				$foundPolicy = $true
			}
		}
		
		if(!$foundPolicy)
		{
			if($Result.ExistingChecked)
			{
				Write-host "INFO: Creating from existing Dial Plan:" $Result.ExistingPolicy -foreground "yellow"
				$CopyDialPlan = Get-CsTenantDialPlan -Identity $Result.ExistingPolicy
				
				$Name = $Result.NewPolicy
				$Description = $CopyDialPlan.Description
				$NormalizationRules = $CopyDialPlan.NormalizationRules
				$ExternalAccessPrefix = $CopyDialPlan.ExternalAccessPrefix
				$OptimizeDeviceDialing = $CopyDialPlan.OptimizeDeviceDialing
				$AccessPrefix = $Result.AccessPrefix
				
				Write-Verbose "$Name $Description $NormalizationRules $ExternalAccessPrefix $OptimizeDeviceDialing $AccessPrefix"
				
				if($AccessPrefix -eq "" -or $AccessPrefix -eq $null)
				{
					Write-Host "RUNNING: New-CsTenantDialPlan -Identity `"$Name`" -NormalizationRules @{Replace=$NormalizationRules}" -foreground "green"
					New-CsTenantDialPlan -Identity "$Name" -NormalizationRules @{Replace=$NormalizationRules}
				}
				else
				{
					Write-Host "RUNNING: New-CsTenantDialPlan -Identity `"$Name`" -NormalizationRules @{Replace=$NormalizationRules} -ExternalAccessPrefix $AccessPrefix -OptimizeDeviceDialing `$true" -foreground "green"
					New-CsTenantDialPlan -Identity "$Name" -NormalizationRules @{Replace=$NormalizationRules} -ExternalAccessPrefix $AccessPrefix -OptimizeDeviceDialing $true
				}
				$policyDropDownBox.Items.Clear()
				Get-CsTenantDialPlan | select-object identity | ForEach-Object {[void] $policyDropDownBox.Items.Add(($_.identity).Replace("Tag:",""))}
				$policyDropDownBox.SelectedIndex = $policyDropDownBox.Items.IndexOf("$Name")
			}
			else
			{
				$NewPolicyName = $Result.NewPolicy
				$AccessPrefix = $Result.AccessPrefix
				#$AccessPrefix = $AccessPrefixTextBox.Text.ToString()
				Write-host "INFO: New Dial Plan - $NewPolicyName" -foreground "yellow"
				
				if($AccessPrefix -eq "" -or $AccessPrefix -eq $null)
				{
					Write-Host "RUNNING: New-CsTenantDialPlan -Identity `"$NewPolicyName`"" -foreground "green"
					New-CsTenantDialPlan -Identity "$NewPolicyName"
				}
				else
				{
					Write-Host "RUNNING: New-CsTenantDialPlan -Identity `"$NewPolicyName`" -ExternalAccessPrefix $AccessPrefix -OptimizeDeviceDialing `$true" -foreground "green"
					New-CsTenantDialPlan -Identity "$NewPolicyName" -ExternalAccessPrefix $AccessPrefix -OptimizeDeviceDialing $true
				}
				
				#REMOVE ALL THE DEFAULT VALUES###############
				$policyItem = (Get-CsTenantDialPlan -identity "Global").NormalizationRules
				Set-CsTenantDialPlan -Identity $NewPolicyName -NormalizationRules @{Remove=$policyItem}
				#############################################
				
				$policyDropDownBox.Items.Clear()
				Get-CsTenantDialPlan | select-object identity | ForEach-Object {[void] $policyDropDownBox.Items.Add(($_.identity).Replace("Tag:",""))}
				$policyDropDownBox.SelectedIndex = $policyDropDownBox.Items.IndexOf("$NewPolicyName")
				
			}
		}
		else
		{
			Write-Host "ERROR: Dial Plan with this name already exists" -foreground "red"
		}

	}
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($NewPolicyButton)

#NewPolicy button
$EditPolicyButton = New-Object System.Windows.Forms.Button
$EditPolicyButton.Location = New-Object System.Drawing.Size(392,39)
$EditPolicyButton.Size = New-Object System.Drawing.Size(50,20)
$EditPolicyButton.Text = "Edit.."
$EditPolicyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$EditPolicyButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Opening edit dial plan dialog..."
	[System.Windows.Forms.Application]::DoEvents()
	$PolicyName = $policyDropDownBox.Text.ToString()
	
	$Result = Edit-Policy -Message "Dial Plan name:" -WindowTitle "Edit Tenant Dial Plan" -PolicyName $PolicyName
		
	if($Result -ne $null)
	{
		if($Result.ExistingChecked)
		{
			Write-host "INFO: Creating from existing Dial Plan:" $Result.ExistingPolicy -foreground "yellow"
			$CopyDialPlan = Get-CsTenantDialPlan -Identity $Result.ExistingPolicy
			
			#$Name = $Result.NewPolicy
			$NewPolicyName = $Result.NewPolicy
			$Description = $CopyDialPlan.Description
			$NormalizationRules = $CopyDialPlan.NormalizationRules
			$ExternalAccessPrefix = $CopyDialPlan.ExternalAccessPrefix
			$OptimizeDeviceDialing = $CopyDialPlan.OptimizeDeviceDialing
			$AccessPrefix = $Result.AccessPrefix
			
			if($AccessPrefix -eq "" -or $AccessPrefix -eq $null)
			{
				Write-Host "RUNNING: Set-CsTenantDialPlan -Identity $NewPolicyName -NormalizationRules @{Replace=$NormalizationRules} -OptimizeDeviceDialing `$false" -foreground "green"
				Write-Host "INFO: External Access number value cannot be left blank if it's already set. Value will not be changed but OptimizeDeviceDialing will be disabled." -foreground "yellow"
				Set-CsTenantDialPlan -Identity $NewPolicyName -NormalizationRules @{Replace=$NormalizationRules} -OptimizeDeviceDialing $false
			}
			else
			{
				Write-Host "RUNNING: Set-CsTenantDialPlan -Identity $NewPolicyName -NormalizationRules @{Replace=$NormalizationRules} -ExternalAccessPrefix $AccessPrefix -OptimizeDeviceDialing `$true" -foreground "green"
				Set-CsTenantDialPlan -Identity $NewPolicyName -NormalizationRules @{Replace=$NormalizationRules} -ExternalAccessPrefix $AccessPrefix -OptimizeDeviceDialing $true
			}
			$policyDropDownBox.SelectedIndex = $policyDropDownBox.Items.IndexOf("$PolicyName")
		}
		else
		{
			$NewPolicyName = $Result.NewPolicy
			$AccessPrefix = $Result.AccessPrefix
			#$AccessPrefix = $AccessPrefixTextBox.Text.ToString()
			Write-host "INFO: New Policy - $NewPolicyName" -foreground "yellow"
			#Set-CsTenantDialPlan -Identity $PolicyName
			
			if($AccessPrefix -eq "" -or $AccessPrefix -eq $null)
			{
				
				Write-Host "RUNNING: Set-CsTenantDialPlan -Identity $NewPolicyName -OptimizeDeviceDialing `$false" -foreground "green"
				Write-Host "INFO: External Access number value cannot be left blank if it's already set. Value will not be changed but OptimizeDeviceDialing will be disabled." -foreground "yellow"
				Set-CsTenantDialPlan -Identity $NewPolicyName -OptimizeDeviceDialing $false
			}
			else
			{
				Write-Host "RUNNING: Set-CsTenantDialPlan -Identity $NewPolicyName -ExternalAccessPrefix $AccessPrefix -OptimizeDeviceDialing `$true" -foreground "green"
				Set-CsTenantDialPlan -Identity $NewPolicyName -ExternalAccessPrefix $AccessPrefix -OptimizeDeviceDialing $true
			}
			$policyDropDownBox.SelectedIndex = $policyDropDownBox.Items.IndexOf("$PolicyName")
		}
		
	}
	GetNormalisationPolicy
	$StatusLabel.Text = ""
	
})
$mainForm.Controls.Add($EditPolicyButton)


#NewPolicy button
$RemovePolicyButton = New-Object System.Windows.Forms.Button
$RemovePolicyButton.Location = New-Object System.Drawing.Size(342,65)
$RemovePolicyButton.Size = New-Object System.Drawing.Size(100,20)
$RemovePolicyButton.Text = "Remove"
$RemovePolicyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$RemovePolicyButton.Add_Click(
{	
	$StatusLabel.Text = "STATUS: Removing Policy..."
	[System.Windows.Forms.Application]::DoEvents()
	$thePolicyName = $policyDropDownBox.SelectedItem.ToString()
	$a = new-object -comobject wscript.shell 
	$intAnswer = $a.popup("Are you sure you want to remove the entire $thePolicyName Dial Plan?",0,"Remove Dial Plan",4) 
	if ($intAnswer -eq 6) { 
					
		Write-Host "RUNNING: Remove-CsTenantDialPlan -Identity `"$thePolicyName`"" -foreground "green"
		Invoke-Expression "Remove-CsTenantDialPlan -Identity `"$thePolicyName`""
		
		$policyDropDownBox.Items.Clear()
		Get-CsTenantDialPlan | select-object identity | ForEach-Object {[void] $policyDropDownBox.Items.Add(($_.identity).Replace("Tag:",""))}
		
		$numberOfItems = $policyDropDownBox.count
		if($numberOfItems -gt 0)
		{
			$policyDropDownBox.SelectedIndex = 0
		}
	}
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($RemovePolicyButton)

# Create the Label
$AccessPrefixLabel = New-Object System.Windows.Forms.Label
$AccessPrefixLabel.Location = New-Object System.Drawing.Size(22,70) 
$AccessPrefixLabel.Size = New-Object System.Drawing.Size(80,20)
$AccessPrefixLabel.AutoSize = $true
$AccessPrefixLabel.TabStop = $false
$AccessPrefixLabel.Text = "Access Prefix:"
$mainForm.Controls.Add($AccessPrefixLabel)
	
$AccessPrefixTextBox = New-Object System.Windows.Forms.TextBox
$AccessPrefixTextBox.Location = New-Object System.Drawing.Size(105,67) 
$AccessPrefixTextBox.Size = New-Object System.Drawing.Size(50,20) 
$AccessPrefixTextBox.Text = ""
#$AccessPrefixTextBox.tabIndex = 3
$AccessPrefixTextBox.TabStop = $false
$AccessPrefixTextBox.Enabled = $false
$mainForm.Controls.Add($AccessPrefixTextBox)


# Create the Label
$OptimizeDeviceDialingLabel = New-Object System.Windows.Forms.Label
$OptimizeDeviceDialingLabel.Location = New-Object System.Drawing.Size(170,70) 
$OptimizeDeviceDialingLabel.Size = New-Object System.Drawing.Size(150,20)
$OptimizeDeviceDialingLabel.AutoSize = $true
$OptimizeDeviceDialingLabel.TabStop = $false
$OptimizeDeviceDialingLabel.Text = "OptimizeDeviceDialing: FALSE"
$mainForm.Controls.Add($OptimizeDeviceDialingLabel)


$lv = New-Object windows.forms.ListView
$lv.View = [System.Windows.Forms.View]"Details"
$lv.Size = New-Object System.Drawing.Size(422,250)
$lv.Location = New-Object System.Drawing.Size(20,97)
$lv.FullRowSelect = $true
$lv.GridLines = $true
$lv.HideSelection = $false
#$lv.MultiSelect = $false
#$lv.Sorting = [System.Windows.Forms.SortOrder]"Ascending"
$lv.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top
[void]$lv.Columns.Add("Name", 150)
#[void]$lv.Columns.Add("Priority", 50)
[void]$lv.Columns.Add("Description", 100)
[void]$lv.Columns.Add("Pattern", 75)
[void]$lv.Columns.Add("Translation", 75)
[void]$lv.Columns.Add("Extension", 0)
$mainForm.Controls.Add($lv)

$lv.add_MouseUp(
{
	$StatusLabel.Text = "STATUS: Updating list..."
	[System.Windows.Forms.Application]::DoEvents()
	UpdateListViewSettings
	$StatusLabel.Text = ""
})

# Groups Key Event ============================================================
$lv.add_KeyUp(
{
	if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
	{	
		$StatusLabel.Text = "STATUS: Updating list..."
		[System.Windows.Forms.Application]::DoEvents()
		UpdateListViewSettings
		$StatusLabel.Text = ""
	}
})

$ignoreInput = $false
$lv.add_ColumnWidthChanged(
{
	#Work around for being able to expand the hidden extension column
	if($_.ColumnIndex -eq 4 -and !$ignoreInput)
	{
		$ignoreInput = $true
		$lv.Columns[3].Width = $lv.Columns[3].Width + $lv.Columns[4].Width
		$lv.Columns[4].Width = 0
	}
	else
	{
		$ignoreInput = $false
	}
})


#Up button
$UpButton = New-Object System.Windows.Forms.Button
$UpButton.Location = New-Object System.Drawing.Size(450,160)
$UpButton.Size = New-Object System.Drawing.Size(50,20)
$UpButton.Text = "UP"
$UpButton.TabStop = $false
$UpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$UpButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Moving rule up..."
	[System.Windows.Forms.Application]::DoEvents()
	DisableAllButtons
	Move-Up
	EnableAllButtons
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($UpButton)

# Priority Label ============================================================
$PriorityLabel = New-Object System.Windows.Forms.Label
$PriorityLabel.Location = New-Object System.Drawing.Size(455,140) 
$PriorityLabel.Size = New-Object System.Drawing.Size(60,15) 
$PriorityLabel.Text = "Priority"
$PriorityLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$PriorityLabel.TabStop = $false
$mainForm.Controls.Add($PriorityLabel)

#Down button
$DownButton = New-Object System.Windows.Forms.Button
$DownButton.Location = New-Object System.Drawing.Size(450,190)
$DownButton.Size = New-Object System.Drawing.Size(50,20)
$DownButton.Text = "DOWN"
$DownButton.TabStop = $false
$DownButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$DownButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Moving rule down..."
	[System.Windows.Forms.Application]::DoEvents()
	DisableAllButtons
	Move-Down
	EnableAllButtons
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($DownButton)




#NameTextLabel Label ============================================================
$NameTextLabel = New-Object System.Windows.Forms.Label
$NameTextLabel.Location = New-Object System.Drawing.Size(10,10) 
$NameTextLabel.Size = New-Object System.Drawing.Size(60,15) 
$NameTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$NameTextLabel.Text = "Name:"
$NameTextLabel.TabStop = $false
$NameTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($NameTextLabel)


#Name Text box ============================================================
$NameTextBox = New-Object System.Windows.Forms.TextBox
$NameTextBox.location = new-object system.drawing.size(72,10)
$NameTextBox.size = new-object system.drawing.size(250,23)
$NameTextBox.tabIndex = 1
$NameTextBox.text = ""   
$NameTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.controls.add($NameTextBox)
$NameTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		#Do Nothing
	}
})


#Description Label ============================================================
$DescriptionTextLabel = New-Object System.Windows.Forms.Label
$DescriptionTextLabel.Location = New-Object System.Drawing.Size(5,35) 
$DescriptionTextLabel.Size = New-Object System.Drawing.Size(65,15) 
$DescriptionTextLabel.Text = "Description:"
$DescriptionTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$DescriptionTextLabel.TabStop = $false
$DescriptionTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($DescriptionTextLabel)

#$DescriptionTextBox Text box ============================================================
$DescriptionTextBox = New-Object System.Windows.Forms.TextBox
$DescriptionTextBox.location = new-object system.drawing.size(72,35)
$DescriptionTextBox.size = new-object system.drawing.size(250,23)
$DescriptionTextBox.tabIndex = 1
$DescriptionTextBox.text = ""   
$DescriptionTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.controls.add($DescriptionTextBox)
$DescriptionTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		#AddSetting
	}
})


#Pattern Label ============================================================
$PatternTextLabel = New-Object System.Windows.Forms.Label
$PatternTextLabel.Location = New-Object System.Drawing.Size(5,60) 
$PatternTextLabel.Size = New-Object System.Drawing.Size(65,15) 
$PatternTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$PatternTextLabel.Text = "Pattern:"
$PatternTextLabel.TabStop = $false
$PatternTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($PatternTextLabel)

#Pattern Text box ============================================================
$PatternTextBox = New-Object System.Windows.Forms.TextBox
$PatternTextBox.location = new-object system.drawing.size(72,60)
$PatternTextBox.size = new-object system.drawing.size(250,23)
$PatternTextBox.tabIndex = 1
$PatternTextBox.text = ""   
$PatternTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.controls.add($PatternTextBox)
$PatternTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		#AddSetting
	}
})


#Translation Label ============================================================
$TranslationTextLabel = New-Object System.Windows.Forms.Label
$TranslationTextLabel.Location = New-Object System.Drawing.Size(5,85) 
$TranslationTextLabel.Size = New-Object System.Drawing.Size(65,15) 
$TranslationTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$TranslationTextLabel.Text = "Translation:"
$TranslationTextLabel.TabStop = $false
$TranslationTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($TranslationTextLabel)

#Setting Text box ============================================================
$TranslationTextBox = New-Object System.Windows.Forms.TextBox
$TranslationTextBox.location = new-object system.drawing.size(72,85)
$TranslationTextBox.size = new-object system.drawing.size(250,23)
$TranslationTextBox.tabIndex = 1
$TranslationTextBox.text = ""   
$TranslationTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.controls.add($TranslationTextBox )
$TranslationTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{
		$StatusLabel.Text = "STATUS: Adding normalization rule..."
		[System.Windows.Forms.Application]::DoEvents()
		AddSetting
		$StatusLabel.Text = ""
	}
})

#ExtensionTextLabel Label ============================================================
$ExtensionTextLabel = New-Object System.Windows.Forms.Label
$ExtensionTextLabel.Location = New-Object System.Drawing.Size(5,110) 
$ExtensionTextLabel.Size = New-Object System.Drawing.Size(65,15)
$ExtensionTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight 
$ExtensionTextLabel.Text = "Extension:"
$ExtensionTextLabel.TabStop = $false
$ExtensionTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($ExtensionTextLabel)

$ExtensionCheckBox = New-Object System.Windows.Forms.Checkbox 
$ExtensionCheckBox.Location = New-Object System.Drawing.Size(72,110) 
$ExtensionCheckBox.Size = New-Object System.Drawing.Size(20,20)
$ExtensionCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$ExtensionCheckBox.tabIndex = 2
$ExtensionCheckBox.Add_Click(
{
	
})
#$mainForm.controls.add($ExtensionCheckBox)

#Add button
$AddButton = New-Object System.Windows.Forms.Button
$AddButton.Location = New-Object System.Drawing.Size(340,20)
$AddButton.Size = New-Object System.Drawing.Size(70,18)
$AddButton.Text = "Add / Edit"
$AddButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$AddButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Adding normalization rule..."
	[System.Windows.Forms.Application]::DoEvents()
	AddSetting
	$StatusLabel.Text = ""
})
#$mainForm.Controls.Add($AddButton)


#Delete button
$DeleteButton = New-Object System.Windows.Forms.Button
$DeleteButton.Location = New-Object System.Drawing.Size(340,45)
$DeleteButton.Size = New-Object System.Drawing.Size(70,18)
$DeleteButton.Text = "Delete"
$DeleteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$DeleteButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Deleting normalization rule..."
	[System.Windows.Forms.Application]::DoEvents()
	DeleteSetting
	$StatusLabel.Text = ""
})
#$mainForm.Controls.Add($DeleteButton)


#Add button
$DeleteAllButton = New-Object System.Windows.Forms.Button
$DeleteAllButton.Location = New-Object System.Drawing.Size(340,70)
$DeleteAllButton.Size = New-Object System.Drawing.Size(70,18)
$DeleteAllButton.Text = "Delete All"
$DeleteAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$DeleteAllButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Deleting all normalization rules..."
	[System.Windows.Forms.Application]::DoEvents()
	$a = new-object -comobject wscript.shell
	$intAnswer = $a.popup("Are you sure you want to remove all the rules from this Dial Plan?",0,"Remove All Rules",4) 
	if ($intAnswer -eq 6) { 
					
		DeleteAllSettings
	}
	$StatusLabel.Text = ""
	
})
#$mainForm.Controls.Add($DeleteAllButton)


$GroupBoxNormRule = New-Object System.Windows.Forms.Panel
$GroupBoxNormRule.Location = New-Object System.Drawing.Size(20,357) 
$GroupBoxNormRule.Size = New-Object System.Drawing.Size(420,135) 
$GroupBoxNormRule.MinimumSize = New-Object System.Drawing.Size(400,80) 
$GroupBoxNormRule.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$GroupBoxNormRule.TabStop = $False
$GroupBoxNormRule.Controls.Add($NameTextLabel)
$GroupBoxNormRule.Controls.Add($NameTextBox)
$GroupBoxNormRule.Controls.Add($DescriptionTextLabel)
$GroupBoxNormRule.Controls.Add($DescriptionTextBox)
$GroupBoxNormRule.Controls.Add($PatternTextLabel)
$GroupBoxNormRule.Controls.Add($PatternTextBox)
$GroupBoxNormRule.Controls.Add($TranslationTextLabel)
$GroupBoxNormRule.Controls.Add($TranslationTextBox)
$GroupBoxNormRule.Controls.Add($ExtensionTextLabel)
$GroupBoxNormRule.Controls.Add($ExtensionCheckBox)
$GroupBoxNormRule.Controls.Add($AddButton)
$GroupBoxNormRule.Controls.Add($DeleteButton)
$GroupBoxNormRule.Controls.Add($DeleteAllButton)
$GroupBoxNormRule.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$GroupBoxNormRule.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$mainForm.Controls.Add($GroupBoxNormRule)



#Test Label ============================================================
$TestPhoneTextLabel = New-Object System.Windows.Forms.Label
$TestPhoneTextLabel.Location = New-Object System.Drawing.Size(50,508) 
$TestPhoneTextLabel.Size = New-Object System.Drawing.Size(30,15) 
$TestPhoneTextLabel.Text = "Test:"
$TestPhoneTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$TestPhoneTextLabel.TabStop = $false
$mainForm.Controls.Add($TestPhoneTextLabel)

#Test Text box ============================================================
$TestPhoneTextBox = New-Object System.Windows.Forms.TextBox
$TestPhoneTextBox.location = new-object system.drawing.size(85,505)
$TestPhoneTextBox.size = new-object system.drawing.size(200,23)
$TestPhoneTextBox.tabIndex = 1
$TestPhoneTextBox.text = "0407532999"   
$TestPhoneTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$mainForm.controls.add($TestPhoneTextBox)
$TestPhoneTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{
		$StatusLabel.Text = "STATUS: Running test..."
		[System.Windows.Forms.Application]::DoEvents()
		DisableAllButtons
		TestPhoneNumberNew
		EnableAllButtons
		$StatusLabel.Text = ""
	}
})

#Add button
$TestPhoneButton = New-Object System.Windows.Forms.Button
$TestPhoneButton.Location = New-Object System.Drawing.Size(303,505)
$TestPhoneButton.Size = New-Object System.Drawing.Size(87,18)
$TestPhoneButton.Text = "Test Number"
$TestPhoneButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$TestPhoneButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Running Test..."
	[System.Windows.Forms.Application]::DoEvents()
	DisableAllButtons
	TestPhoneNumberNew
	EnableAllButtons
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($TestPhoneButton)


#Pattern Label ============================================================
$TestPhonePatternTextLabel = New-Object System.Windows.Forms.Label
$TestPhonePatternTextLabel.Location = New-Object System.Drawing.Size(20,10) 
$TestPhonePatternTextLabel.Size = New-Object System.Drawing.Size(400,15) 
$TestPhonePatternTextLabel.Text = "Matched Pattern:"
$TestPhonePatternTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$TestPhonePatternTextLabel.TabStop = $false
#$mainForm.Controls.Add($TestPhonePatternTextLabel)

#Translation Label ============================================================
$TestPhoneTranslationTextLabel = New-Object System.Windows.Forms.Label
$TestPhoneTranslationTextLabel.Location = New-Object System.Drawing.Size(20,30) 
$TestPhoneTranslationTextLabel.Size = New-Object System.Drawing.Size(400,15) 
$TestPhoneTranslationTextLabel.Text = "Matched Translation:"
$TestPhoneTranslationTextLabel.TabStop = $false
$TestPhoneTranslationTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($TestPhoneTranslationTextLabel)


#Result Label ============================================================
$TestPhoneResultTextLabel = New-Object System.Windows.Forms.Label
$TestPhoneResultTextLabel.Location = New-Object System.Drawing.Size(20,50) 
$TestPhoneResultTextLabel.Size = New-Object System.Drawing.Size(400,15) 
$TestPhoneResultTextLabel.Text = "Test Result:"
$TestPhoneResultTextLabel.TabStop = $false
$Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
$TestPhoneResultTextLabel.Font = $Font 
$TestPhoneResultTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($TestPhoneResultTextLabel)

$GroupBoxCurrent = New-Object System.Windows.Forms.Panel
$GroupBoxCurrent.Location = New-Object System.Drawing.Size(20,530) 
$GroupBoxCurrent.Size = New-Object System.Drawing.Size(420,80) 
$GroupBoxCurrent.MinimumSize = New-Object System.Drawing.Size(400,80) 
$GroupBoxCurrent.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$GroupBoxCurrent.TabStop = $False
$GroupBoxCurrent.Controls.Add($TestPhonePatternTextLabel)
$GroupBoxCurrent.Controls.Add($TestPhoneTranslationTextLabel)
$GroupBoxCurrent.Controls.Add($TestPhoneResultTextLabel)
$GroupBoxCurrent.BackColor = [System.Drawing.Color]::White
$GroupBoxCurrent.ForeColor = [System.Drawing.Color]::Black
$GroupBoxCurrent.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$GroupBoxCurrent.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$mainForm.Controls.Add($GroupBoxCurrent)


<#
#Import Label ============================================================
$ImportTextLabel = New-Object System.Windows.Forms.Label
$ImportTextLabel.Location = New-Object System.Drawing.Size(50,578) 
$ImportTextLabel.Size = New-Object System.Drawing.Size(80,15) 
$ImportTextLabel.Text = "Import/Export:"
$ImportTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ImportTextLabel.TabStop = $false
$mainForm.Controls.Add($ImportTextLabel)


#Import button
$ImportButton = New-Object System.Windows.Forms.Button
$ImportButton.Location = New-Object System.Drawing.Size(130,575)
$ImportButton.Size = New-Object System.Drawing.Size(120,20)
$ImportButton.Text = "Import Config"
$ImportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ImportButton.Add_Click(
{
	Import-Config
	UpdateListViewSettings
	
})
$mainForm.Controls.Add($ImportButton)


#Export button
$ExportButton = New-Object System.Windows.Forms.Button
$ExportButton.Location = New-Object System.Drawing.Size(260,575)
$ExportButton.Size = New-Object System.Drawing.Size(120,20)
$ExportButton.Text = "Export Config"
$ExportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ExportButton.Add_Click(
{
	Export-Config

})
$mainForm.Controls.Add($ExportButton)
#>


# Add the Status Label ============================================================
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Size(15,620) 
$StatusLabel.Size = New-Object System.Drawing.Size(420,15) 
$StatusLabel.Text = ""
$StatusLabel.forecolor = "DarkBlue"
$StatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$StatusLabel.TabStop = $false
$mainForm.Controls.Add($StatusLabel)


$ToolTip = New-Object System.Windows.Forms.ToolTip 
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
$ToolTip.IsBalloon = $true 
$ToolTip.InitialDelay = 500 
$ToolTip.ReshowDelay = 500 
$ToolTip.AutoPopDelay = 10000
#$ToolTip.ToolTipTitle = "Help:"
$ToolTip.SetToolTip($AddButton, "If the specified Name is the same as an existing rule`r`nthen than rule will be edited. If the Name is new`r`nthen a new rule will be created.") 
$ToolTip.SetToolTip($DeleteButton, "The Delete button will delete the selected rule(s).") 
$ToolTip.SetToolTip($DeleteAllButton, "The Delete All button will delete all the rules in this Dial Plan.") 
$ToolTip.SetToolTip($OptimizeDeviceDialingLabel, "Indicates whether Access Prefix`r`nis being applied by the system.")


function ConnectSkypeForBusinessOnline
{
	$ConnectOnlineButton.Text = "Connecting..."
	$StatusLabel.Text = "STATUS: Connecting to O365..."
	Write-Host "INFO: Connecting to O365..." -foreground "Yellow"
	[System.Windows.Forms.Application]::DoEvents()
	if($global:SFBOsession)
	{
		Remove-PSSession $global:SFBOsession
	}
	if (Get-Module -ListAvailable -Name SkypeOnlineConnector)
	{
		Import-module SkypeOnlineConnector
		
		if($Script:O365Creds -ne $null)
		{
			$cred = $Script:O365Creds
		}
		elseif($script:OnlineUsername -ne "" -and $script:OnlineUsername -ne $null -and $script:OnlinePassword -ne "" -and $script:OnlinePassword -ne $null)
		{
			$secpasswd = ConvertTo-SecureString $script:OnlinePassword -AsPlainText -Force
			$cred = New-Object System.Management.Automation.PSCredential ($script:OnlineUsername, $secpasswd)
		}
		elseif($script:OnlineUsername -ne "" -and $script:OnlineUsername -ne $null)
		{
			$cred = Get-Credential -Username $script:OnlineUsername -Message "Skype for Business Online"
		}
		else
		{
			$cred = Get-Credential -Message "Skype for Business Online"
		}
		
		if($cred)
		{
			try
			{
				$global:SFBOsession = New-CsOnlineSession -Credential $cred -ErrorAction Stop #-SessionOption $pso   #MFA FAILS HERE
				$result = Import-PSSession $global:SFBOsession -AllowClobber
				if($result -ne $null)
				{
					$Script:O365Creds = $cred
					$Script:O365ReconnectAttempts = 0
				}
				$Script:UserConnectedToSfBOnline = $true
				
				#Fill-Content
				CheckSkypeForBusinessOnlineInitial
				EnableAllButtons
				<#
				if(([array] (Get-CsOnlinePSTNGateway -ErrorAction SilentlyContinue)).count -eq 0)
				{
					$NoUsagesWarningLabel.Text = "No Gateways assigned. Add a gateway to get started."
				}
				else
				{
					$NoUsagesWarningLabel.Text = "This Voice Routing Policy has no Usages assigned."
				}
				#>
				$StatusLabel.Text = ""
				return $true
			}
			catch
			{
				if($_ -match "you must use multi-factor authentication to access") #MFA FALLBACK!
				{
					Import-Module SkypeOnlineConnector
					$sfbSession = New-CsOnlineSession -UserName $cred.Username
					$result = Import-PSSession $sfbSession
					if($result -ne $null)
					{
						$Script:O365Creds = $cred
						$Script:O365ReconnectAttempts = 0
					}
					$Script:UserConnectedToSfBOnline = $true
					
					#Fill-Content
					CheckSkypeForBusinessOnlineInitial
					EnableAllButtons
					
					<#
					if(([array] (Get-CsOnlinePSTNGateway -ErrorAction SilentlyContinue)).count -eq 0)
					{
						$NoUsagesWarningLabel.Text = "No Gateways assigned. Add a gateway to get started."
					}
					else
					{
						$NoUsagesWarningLabel.Text = "This Voice Routing Policy has no Usages assigned."
					}
					#>
					$StatusLabel.Text = ""
					return $true
				}
				else
				{
					Write-Host "Error: $_.Exception.Message" -foreground "red"
					$StatusLabel.Text = "ERROR: Connection failed."
					$Script:O365Creds = $null
					$StatusLabel.Text = ""
					return $false	
				}
			}
		}
		else
		{
			Write-Host "Error: No credentials supplied." -foreground "red"
			$StatusLabel.Text = "ERROR: No credentials supplied."
			$StatusLabel.Text = ""
			return $false
		}				
	} 
	else
	{
		Write-host "Please install the Skype for Business Online Windows PowerShell Module" -ForegroundColor "Red"
		Write-host "Located at: https://www.microsoft.com/en-us/download/details.aspx?id=39366" -ForegroundColor "Red"
		$StatusLabel.Text = ""
		return $false
	}
	
}

function CheckSkypeForBusinessOnlineInitial
{	
	#CHECK IF SESSIONS IS AVAILABLE
	$PSSessions = Get-PSSession
	$CurrentlyConnected = $false
	if($PSSessions.count -gt 0)
	{
		foreach($PSSession in $PSSessions)
		{
			if($PSSession.Availability -eq "Available" -and $PSSession.ComputerName -match "lync.com$" )
			{
				$CurrentlyConnected = $true
				$Script:UserConnectedToSfBOnline = $true
				$AvailableFound = $true
			}
			elseif($PSSession.Availability -eq "None" -and $PSSession.ComputerName -match "lync.com$")
			{
				#REMOVE THE MODULE AS IT CAUSES ISSUES
				$NoneFound = $true
			}
			else
			{
				#THIS SESSION IS NOT CONNECTED. IGNORE.
			}
		}
		
		if(!$AvailableFound -and $NoneFound) #No available skypeonline sessions available and old session still exists. Kill it.
		{
			$modules = Get-Module
				
			foreach($module in $modules)
			{
				if($module.name -match "tmp_")
				{
					Write-Host "INFO: Found stale module: " $module.name -foreground "green"
					Write-Host "RUNNING: Remove module " $module.name -foreground "green"
					Remove-Module -name $module.name
				}
			}
			#Force the dialog for the user to decide if they want to re-connect or not
			$CurrentlyConnected = $false
			$Script:UserConnectedToSfBOnline = $true
		}
	}
	
	#CHECK IF COMMANDS ARE AVAILABLE		
	$command = "Get-CsOnlineUser"
	if($CurrentlyConnected -and (Get-Command $command -errorAction SilentlyContinue) -and ($Script:UserConnectedToSfBOnline -eq $true))
	{
		#CHECK THAT SfB ONLINE COMMANDS WORK
		if(([array] (Get-CsOnlineUser -ResultSize 1 -ErrorAction SilentlyContinue)).count -gt 0)
		{
			#Write-Host "Connected to Skype for Business Online" -foreground "Green"
			$ConnectedOnlineLabel.Visible = $true
			$ConnectOnlineButton.Text = "Disconnect SfBO"
			
			Get-CsTenantDialPlan | select-object identity | ForEach-Object {[void] $policyDropDownBox.Items.Add(($_.identity).Replace("Tag:",""))}
	
			$numberOfItems = $policyDropDownBox.Items.count
			if($numberOfItems -gt 0)
			{
				$policyDropDownBox.SelectedIndex = 0
			}
			GetNormalisationPolicy
			
			if($currentIndex -ne $null)
			{
				if($currentIndex -lt $dgv.Rows.Count)
				{$dgv.Rows[$currentIndex].Selected = $True}
			}
		}
		else
		{
			Write-Host "INFO: Cannot access O365" -foreground "Yellow"
			$ConnectedOnlineLabel.Visible = $false
			$ConnectOnlineButton.Text = "Connect SfBO"
			#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
			
			DisableAllButtons
			
			[System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.MessageBox]::Show("The SfBOnline connection has been disconnected. Click OK to reconnect.", "SfB Online Connection", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
			if($result -eq [System.Windows.Forms.DialogResult]::OK)
			{
				Write-Host "INFO: Re-establishing connection" -foreground "yellow"
				$ConnectOnlineButton.Enabled = $false
	
				$ConnectResult = ConnectSkypeForBusinessOnline
				if($ConnectResult)
				{
					$ConnectedOnlineLabel.Visible = $true
					$ConnectOnlineButton.Text = "Disconnect SfBO"
				}
				else
				{
					$ConnectedOnlineLabel.Visible = $false
					$ConnectOnlineButton.Text = "Connect SfBO"
					#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
					$CurrentlyConnected = $false
					$Script:UserConnectedToSfBOnline = $false
					Write-Host "ERROR: Failed to connect to Skype for Business Online..." -foreground "red"
					$StatusLabel.Text = "ERROR: Connection failed."
					
					return $false
				}
				
				$ConnectOnlineButton.Enabled = $true
			}
			elseif($result -eq [System.Windows.Forms.DialogResult]::Cancel)
			{
				Write-Host "INFO: Disconnecting from O365" -foreground "yellow"
				
				DisconnectSkypeForBusinessOnline
				
				$ConnectOnlineButton.Enabled = $false
				$ConnectedOnlineLabel.Visible = $false
				$ConnectOnlineButton.Text = "Connect SfBO"
				#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
				$CurrentlyConnected = $false
				$Script:UserConnectedToSfBOnline = $false
				
				$ConnectOnlineButton.Enabled = $true
				$Script:O365Creds = $null #CANCELLING SO DELETE CREDS
							
			}
		
		}
	}
	elseif(($CurrentlyConnected -eq $false) -and ($Script:UserConnectedToSfBOnline -eq $true)) #User has connected to SfBOnline but SfBOnline is reporting being disconnected. Ask if they want to reconnect.
	{
		Write-Host "INFO: Not Connected to Skype for Business Online" -foreground "Yellow"
		$ConnectedOnlineLabel.Visible = $false
		$ConnectOnlineButton.Text = "Connect SfBO"
		#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
		
		DisableAllButtons
		
		[System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.MessageBox]::Show("The SfBOnline connection has been disconnected. Click OK to reconnect.", "SfB Online Connection", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
		if($result -eq [System.Windows.Forms.DialogResult]::OK)
		{
			Write-Host "INFO: Re-establishing connection" -foreground "yellow"
			$ConnectOnlineButton.Enabled = $false
	
			$ConnectResult = ConnectSkypeForBusinessOnline
			if($ConnectResult)
			{
				$ConnectedOnlineLabel.Visible = $true
				$ConnectOnlineButton.Text = "Disconnect SfBO"
			}
			else
			{
				$ConnectedOnlineLabel.Visible = $false
				$ConnectOnlineButton.Text = "Connect SfBO"
				#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
				$CurrentlyConnected = $false
				$Script:UserConnectedToSfBOnline = $false
				Write-Host "ERROR: Failed to connect to Skype for Business Online..." -foreground "red"
				$StatusLabel.Text = "ERROR: Connection failed."
				
				return $false
				
			}
			
			$ConnectOnlineButton.Enabled = $true
						
		}
		elseif($result -eq [System.Windows.Forms.DialogResult]::Cancel)
		{
			Write-Host "INFO: Disconnecting from O365" -foreground "yellow"
			
			DisconnectSkypeForBusinessOnline
			
			$ConnectOnlineButton.Enabled = $false
			$ConnectedOnlineLabel.Visible = $false
			$ConnectOnlineButton.Text = "Connect SfBO"
			#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
			$CurrentlyConnected = $false
			$Script:UserConnectedToSfBOnline = $false
			
			$ConnectOnlineButton.Enabled = $true
			$Script:O365Creds = $null #CANCELLING SO DELETE CREDS
						
		}
	}
	elseif(!$CurrentlyConnected) 
	{
		Write-Host "INFO: Cannot access Skype for Business Online" -foreground "Yellow"
		$ConnectedOnlineLabel.Visible = $false
		$ConnectOnlineButton.Text = "Connect SfBO"
		#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
		$ConnectOnlineButton.Enabled = $true
		
		DisableAllButtons
		
		return $false
	}
	
	return $true
}


function CheckSkypeForBusinessOnline
{	
	#CHECK IF SESSIONS IS AVAILABLE
	$PSSessions = Get-PSSession
	$CurrentlyConnected = $false
	if($PSSessions.count -gt 0)
	{
		foreach($PSSession in $PSSessions)
		{
			if($PSSession.Availability -eq "Available" -and $PSSession.ComputerName -match "lync.com$" )
			{
				$CurrentlyConnected = $true
				$Script:UserConnectedToSfBOnline = $true
				$AvailableFound = $true
			}
			elseif($PSSession.Availability -eq "None" -and $PSSession.ComputerName -match "lync.com$")
			{
				#REMOVE THE MODULE AS IT CAUSES ISSUES
				$NoneFound = $true
			}
			else
			{
				#THIS SESSION IS NOT CONNECTED. IGNORE.
			}
		}
		
		if(!$AvailableFound -and $NoneFound) #No available skypeonline sessions available and old session still exists. Kill it.
		{
			$modules = Get-Module
				
			foreach($module in $modules)
			{
				if($module.name -match "tmp_")
				{
					Write-Host "INFO: Found stale module: " $module.name -foreground "green"
					Write-Host "RUNNING: Remove module " $module.name -foreground "green"
					Remove-Module -name $module.name
				}
			}
			#Force the dialog for the user to decide if they want to re-connect or not
			$CurrentlyConnected = $false
			$Script:UserConnectedToSfBOnline = $true
		}
	}
	
	#CHECK IF COMMANDS ARE AVAILABLE		
	$command = "Get-CsOnlineUser"
	if($CurrentlyConnected -and (Get-Command $command -errorAction SilentlyContinue) -and ($Script:UserConnectedToSfBOnline -eq $true))
	{
		#CHECK THAT SfB ONLINE COMMANDS WORK
		if(([array] (Get-CsOnlineUser -ResultSize 1 -ErrorAction SilentlyContinue)).count -gt 0)
		{
			$ConnectedOnlineLabel.Visible = $true
			$ConnectOnlineButton.Text = "Disconnect SfBO"
		}
		else
		{
			Write-Host "INFO: Cannot access Skype for Business Online" -foreground "Yellow"
			$ConnectedOnlineLabel.Visible = $false
			$ConnectOnlineButton.Text = "Connect SfBO"
			#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
			
			DisableAllButtons
			
			[System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.MessageBox]::Show("The SfBOnline connection has been disconnected. Click OK to reconnect.", "SfB Online Connection", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
			if($result -eq [System.Windows.Forms.DialogResult]::OK)
			{
				Write-Host "INFO: Re-establishing connection" -foreground "yellow"
				$ConnectOnlineButton.Enabled = $false
	
				$ConnectResult = ConnectSkypeForBusinessOnline
				if($ConnectResult)
				{
					$ConnectedOnlineLabel.Visible = $true
					$ConnectOnlineButton.Text = "Disconnect SfBO"
				}
				else
				{
					$ConnectedOnlineLabel.Visible = $false
					$ConnectOnlineButton.Text = "Connect SfBO"
					#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
					$CurrentlyConnected = $false
					$Script:UserConnectedToSfBOnline = $false
					Write-Host "ERROR: Failed to connect to Skype for Business Online..." -foreground "red"
					$StatusLabel.Text = "ERROR: Connection failed."

					return $false
				}
				
				$ConnectOnlineButton.Enabled = $true
			}
			elseif($result -eq [System.Windows.Forms.DialogResult]::Cancel)
			{
				Write-Host "INFO: Disconnecting from O365" -foreground "yellow"
				
				DisconnectSkypeForBusinessOnline
				
				$ConnectOnlineButton.Enabled = $false
				$ConnectedOnlineLabel.Visible = $false
				$ConnectOnlineButton.Text = "Connect SfBO"
				#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
				$CurrentlyConnected = $false
				$Script:UserConnectedToSfBOnline = $false
				
				$ConnectOnlineButton.Enabled = $true
							
			}
			
		}
	}
	elseif(($CurrentlyConnected -eq $false) -and ($Script:UserConnectedToSfBOnline -eq $true)) #User has connected to SfBOnline but SfBOnline is reporting being disconnected. Ask if they want to reconnect.
	{
		
		Write-Host "INFO: Not Connected to Skype for Business Online" -foreground "Yellow"
		$ConnectedOnlineLabel.Visible = $false
		$ConnectOnlineButton.Text = "Connect SfBO"
		#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
		
		DisableAllButtons
		
		[System.Windows.Forms.DialogResult] $result = [System.Windows.Forms.MessageBox]::Show("The SfBOnline connection has been disconnected. Click OK to reconnect.", "SfB Online Connection", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
		if($result -eq [System.Windows.Forms.DialogResult]::OK)
		{
			#Write-Host "YES"
			Write-Host "INFO: Re-establishing connection" -foreground "yellow"
			$ConnectOnlineButton.Enabled = $false
	
			$ConnectResult = ConnectSkypeForBusinessOnline
			if($ConnectResult)
			{
				$ConnectedOnlineLabel.Visible = $true
				$ConnectOnlineButton.Text = "Disconnect SfBO"
			}
			else
			{
				$ConnectedOnlineLabel.Visible = $false
				$ConnectOnlineButton.Text = "Connect SfBO"
				#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
				$CurrentlyConnected = $false
				$Script:UserConnectedToSfBOnline = $false
				Write-Host "ERROR: Failed to connect to Skype for Business Online..." -foreground "red"
				$StatusLabel.Text = "ERROR: Connection failed."

				return $false
				
			}
			
			$ConnectOnlineButton.Enabled = $true
						
		}
		elseif($result -eq [System.Windows.Forms.DialogResult]::Cancel)
		{
			Write-Host "INFO: Disconnecting from O365" -foreground "yellow"
			
			DisconnectSkypeForBusinessOnline
			
			$ConnectOnlineButton.Enabled = $false
			$ConnectedOnlineLabel.Visible = $false
			$ConnectOnlineButton.Text = "Connect SfBO"
			#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
			$CurrentlyConnected = $false
			$Script:UserConnectedToSfBOnline = $false
			
			$ConnectOnlineButton.Enabled = $true
						
		}
	}
	elseif(!$CurrentlyConnected) 
	{
		$ConnectedOnlineLabel.Visible = $false
		$ConnectOnlineButton.Text = "Connect SfBO"
		#$NoUsagesWarningLabel.Text = "Press the `"Connect SfBO`" button to get started."
		$ConnectOnlineButton.Enabled = $true
		
		DisableAllButtons
		
		return $false
	}
	
	return $true
}

function DisconnectSkypeForBusinessOnline
{
	$PSSessions = Get-PSSession
	$CurrentlyConnected = $false
	if($PSSessions.count -gt 0)
	{
		foreach($PSSession in $PSSessions)
		{
			if($PSSession.ComputerName -match "lync.com$" )
			{
				Write-Host "RUNNING: Remove-PSSession" $PSSession.Name -foreground "Green"
				Remove-PSSession $PSSession
			}
		}
	}
	Write-Host "RUNNING: Remove-Module SkypeOnlineConnector" -foreground "Green"
	Remove-Module SkypeOnlineConnector
	
	Write-Host "RUNNING: Get-Module -ListAvailable -Name SkypeOnlineConnector" -foreground "Green"
	$result = Invoke-Expression "Get-Module -ListAvailable -Name SkypeOnlineConnector"
	if($result -ne $null)
	{
		Write-Host "SkypeOnlineConnector has been removed successfully" -foreground "Green"
	}
	else
	{
		Write-Host "ERROR: SkypeOnlineConnector was not removed." -foreground "red"
	}
	
	$modules = Get-Module
	foreach($module in $modules)
	{
		if($module.name -match "tmp_")
		{
			Write-Host "INFO: Removing module: " $module.name -foreground "yellow"
			Write-Host "RUNNING: Remove module " $module.name -foreground "green"
			Remove-Module -name $module.name
		}
	}

	#$Script:O365Creds = $null
	
	DisableAllButtons
}

function DisableAllButtons()
{
	$policyDropDownBox.Enabled = $false
	$NewPolicyButton.Enabled = $false
	$RemovePolicyButton.Enabled = $false
	$UpButton.Enabled = $false
	$DownButton.Enabled = $false
	$NameTextBox.Enabled = $false
	$DescriptionTextBox.Enabled = $false
	$PatternTextBox.Enabled = $false
	$TranslationTextBox.Enabled = $false
	$AddButton.Enabled = $false
	$DeleteButton.Enabled = $false
	$DeleteAllButton.Enabled = $false
	$TestPhoneTextBox.Enabled = $false
	$TestPhoneButton.Enabled = $false
	$EditPolicyButton.Enabled = $false
	$ExtensionCheckBox.Enabled = $false
}


function EnableAllButtons()
{
	$policyDropDownBox.Enabled = $true
	$NewPolicyButton.Enabled = $true
	$RemovePolicyButton.Enabled = $true
	$UpButton.Enabled = $true
	$DownButton.Enabled = $true
	$NameTextBox.Enabled = $true
	$DescriptionTextBox.Enabled = $true
	$PatternTextBox.Enabled = $true
	$TranslationTextBox.Enabled = $true
	$AddButton.Enabled = $true
	$DeleteButton.Enabled = $true
	$DeleteAllButton.Enabled = $true
	$TestPhoneTextBox.Enabled = $true
	$TestPhoneButton.Enabled = $true
	$EditPolicyButton.Enabled = $true
	$ExtensionCheckBox.Enabled = $true
}

function New-Policy([string]$Message, [string]$WindowTitle, [string]$DefaultText)
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
     
    # Create the Label
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10) 
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = $Message
     	
	$PolicyTextBox = New-Object System.Windows.Forms.TextBox
	$PolicyTextBox.Location = New-Object System.Drawing.Size(10,30) 
	$PolicyTextBox.Size = New-Object System.Drawing.Size(300,20) 
	$PolicyTextBox.Text = "<Enter Dial Plan Name>"
	$PolicyTextBox.tabIndex = 1

	$label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Size(10,60) 
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.AutoSize = $true
    $label2.Text = "Copy Normalization Rules from existing Dial Plan:"
	
	# PoliciesDropDownBox ============================================================
	$PoliciesDropDownBox = New-Object System.Windows.Forms.ComboBox 
	$PoliciesDropDownBox.Location = New-Object System.Drawing.Size(10,80) 
	$PoliciesDropDownBox.Size = New-Object System.Drawing.Size(280,20) 
	$PoliciesDropDownBox.DropDownHeight = 200 
	$PoliciesDropDownBox.tabIndex = 1
	$PoliciesDropDownBox.Sorted = $true
	$PoliciesDropDownBox.DropDownStyle = "DropDownList"
	$PoliciesDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	$checkResult = CheckSkypeForBusinessOnline
	if($checkResult)
	{
		Get-CsTenantDialPlan | select-object identity | ForEach-Object {[void] $PoliciesDropDownBox.Items.Add(($_.identity).Replace("Tag:",""))}
	}
	
	if($PoliciesDropDownBox.Items.Count -ge 0)
	{
		$PoliciesDropDownBox.SelectedIndex = 0
	}
	$PoliciesDropDownBox.Enabled = $false
	
	$CopyCheckBox = New-Object System.Windows.Forms.Checkbox 
	$CopyCheckBox.Location = New-Object System.Drawing.Size(295,80) 
	$CopyCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$CopyCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$CopyCheckBox.tabIndex = 2
	$CopyCheckBox.Add_Click(
	{
		if($CopyCheckBox.Checked -eq $false)
		{
			#$PolicyTextBox.Text = "<Enter Policy Name>"
			#$PolicyTextBox.Enabled = $true
			$PoliciesDropDownBox.Enabled = $false
		}
		else
		{
			#$PolicyTextBox.Text = ""
			#$PolicyTextBox.Enabled = $false
			$PoliciesDropDownBox.Enabled = $true
		}
	})
	
	# Create the Label
    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = New-Object System.Drawing.Size(10,113) 
    $label3.Size = New-Object System.Drawing.Size(90,20)
    $label3.AutoSize = $true
    $label3.Text = "Access Prefix:"
     	
	$AccessPrefixTextBox = New-Object System.Windows.Forms.TextBox
	$AccessPrefixTextBox.Location = New-Object System.Drawing.Size(95,110) 
	$AccessPrefixTextBox.Size = New-Object System.Drawing.Size(50,20) 
	$AccessPrefixTextBox.Text = ""
	$AccessPrefixTextBox.tabIndex = 3
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(150,150)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
	if($CopyCheckBox.Checked -eq $false)
	{
		$Result = New-Object PSObject -Property @{
		  NewPolicy = $PolicyTextBox.Text.ToString()
		  ExistingChecked = $false
		  ExistingPolicy = $PoliciesDropDownBox.Text.ToString()
		  AccessPrefix = $AccessPrefixTextBox.Text.ToString()
		}
		$form.Tag = $Result
	}
	else
	{
		$Result = New-Object PSObject -Property @{
		  NewPolicy = $PolicyTextBox.Text.ToString()
		  ExistingChecked = $true
		  ExistingPolicy = $PoliciesDropDownBox.Text.ToString()
		  AccessPrefix = $AccessPrefixTextBox.Text.ToString()
		}
		$form.Tag = $Result
	}
	$form.Close() 
	})
     
    # Create the Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(240,150)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })
     
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(350,220)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.ShowInTaskbar = $true
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
     
    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($PolicyTextBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
	$form.Controls.Add($label2)
	$form.Controls.Add($PoliciesDropDownBox)
	$form.Controls.Add($CopyCheckBox)
	$form.Controls.Add($label3)
	$form.Controls.Add($AccessPrefixTextBox)

     
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
    # Return the text that the user entered.
    return $form.Tag
}

function Edit-Policy([string]$Message, [string]$WindowTitle, [string]$PolicyName)
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
     
    # Create the Label
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10) 
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = $Message
     	
	$PolicyTextBox = New-Object System.Windows.Forms.TextBox
	$PolicyTextBox.Location = New-Object System.Drawing.Size(10,30) 
	$PolicyTextBox.Size = New-Object System.Drawing.Size(300,20) 
	$PolicyTextBox.Text = "$PolicyName"
	$PolicyTextBox.tabIndex = 1
	$PolicyTextBox.Enabled = $false
	
	
	$label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Size(10,60) 
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.AutoSize = $true
    $label2.Text = "Overwrite Normalization Rules from an existing Dial Plan:"
	
	# PoliciesDropDownBox ============================================================
	$PoliciesDropDownBox = New-Object System.Windows.Forms.ComboBox 
	$PoliciesDropDownBox.Location = New-Object System.Drawing.Size(10,80) 
	$PoliciesDropDownBox.Size = New-Object System.Drawing.Size(280,20) 
	$PoliciesDropDownBox.DropDownHeight = 200 
	$PoliciesDropDownBox.tabIndex = 1
	$PoliciesDropDownBox.Sorted = $true
	$PoliciesDropDownBox.DropDownStyle = "DropDownList"
	$PoliciesDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	$checkResult = CheckSkypeForBusinessOnline
	if($checkResult)
	{
		Get-CsTenantDialPlan | Select-Object identity | ForEach-Object {[void] $PoliciesDropDownBox.Items.Add(($_.identity).Replace("Tag:",""))}
	}
	
	if($PoliciesDropDownBox.Items.Count -ge 0)
	{
		$PoliciesDropDownBox.SelectedIndex = 0
	}
	$PoliciesDropDownBox.Enabled = $false
	
	$CopyCheckBox = New-Object System.Windows.Forms.Checkbox 
	$CopyCheckBox.Location = New-Object System.Drawing.Size(295,80) 
	$CopyCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$CopyCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$CopyCheckBox.tabIndex = 2
	$CopyCheckBox.Add_Click(
	{
		if($CopyCheckBox.Checked -eq $false)
		{
			#$PolicyTextBox.Text = "<Enter Policy Name>"
			#$PolicyTextBox.Enabled = $true
			$PoliciesDropDownBox.Enabled = $false
		}
		else
		{
			#$PolicyTextBox.Text = ""
			#$PolicyTextBox.Enabled = $false
			$PoliciesDropDownBox.Enabled = $true
		}
	})
	
	# Create the Label
    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = New-Object System.Drawing.Size(10,113) 
    $label3.Size = New-Object System.Drawing.Size(90,20)
    $label3.AutoSize = $true
    $label3.Text = "Access Prefix:"
     	
	$AccessPrefixTextBox = New-Object System.Windows.Forms.TextBox
	$AccessPrefixTextBox.Location = New-Object System.Drawing.Size(95,110) 
	$AccessPrefixTextBox.Size = New-Object System.Drawing.Size(50,20) 
	$AccessPrefixTextBox.Text = ""
	$AccessPrefixTextBox.tabIndex = 3
	
	$checkResult = CheckSkypeForBusinessOnline
	if($checkResult)
	{
		Get-CsTenantDialPlan -identity $PolicyName | ForEach-Object {$AccessPrefixTextBox.Text = $_.ExternalAccessPrefix; $Optimized = $_.OptimizeDeviceDialing; $OptimizeDeviceDialingLabel.Text = "OptimizeDeviceDialing: $Optimized"}
	}
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(150,150)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
	if($CopyCheckBox.Checked -eq $false)
	{
		$Result = New-Object PSObject -Property @{
		  NewPolicy = $PolicyTextBox.Text.ToString()
		  ExistingChecked = $false
		  ExistingPolicy = $PoliciesDropDownBox.Text.ToString()
		  AccessPrefix = $AccessPrefixTextBox.Text.ToString()
		}
		$form.Tag = $Result
	}
	else
	{
		$Result = New-Object PSObject -Property @{
		  NewPolicy = $PolicyTextBox.Text.ToString()
		  ExistingChecked = $true
		  ExistingPolicy = $PoliciesDropDownBox.Text.ToString()
		  AccessPrefix = $AccessPrefixTextBox.Text.ToString()
		}
		$form.Tag = $Result
	}
	$form.Close() 
	})
     
    # Create the Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(240,150)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })
     
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(350,220)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.ShowInTaskbar = $true
	[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
     
    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($PolicyTextBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
	$form.Controls.Add($label2)
	$form.Controls.Add($PoliciesDropDownBox)
	$form.Controls.Add($CopyCheckBox)
	$form.Controls.Add($label3)
	$form.Controls.Add($AccessPrefixTextBox)

     
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
    # Return the text that the user entered.
    return $form.Tag
}


function Move-Up
{
	$checkResult = CheckSkypeForBusinessOnline
	if($checkResult)
	{
		foreach ($lvi in $lv.SelectedItems)
		{
			#GET SETTINGS OF SELECTED ITEM
			$item = $lv.Items[$lvi.Index]
			$itemValue = $item.SubItems

			[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
			[string]$Name = $item.Text
			#[string]$Priority = $itemValue[1].Text
			[string]$Description = $itemValue[1].Text
			[string]$Pattern = $itemValue[2].Text
			[string]$Translation = $itemValue[3].Text
			[bool]$ExtensionValue = $itemValue[4].Text
							
			$orgIndex = $lvi.Index
			if($orgIndex -gt 0)
			{
				$index = $orgIndex - 1
					
				Write-Host "RUNNING: `$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index -IsInternalExtension $ExtensionValue -InMemory" -foreground "green"
				$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index -IsInternalExtension $ExtensionValue -InMemory			
				
				$NormArray = (Get-CsTenantDialPlan -identity $Scope | select NormalizationRules).NormalizationRules
				Write-Verbose "INITIAL ARRAY"
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host
				#Remove Item
				$NormArray.RemoveAt($orgIndex)
				Write-Verbose "AFTER DELETE"
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host 
				#Insert Item
				$NormArray.Insert($index,$nr)
				Write-Verbose "AFTER INSERT"
				foreach($item in $NormArray){Write-Verbose $item}
				
				Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Replace=$NormArray}	
				
				GetNormalisationPolicy
				
				$lv.Items[$index].Selected = $true
				$lv.Items[$index].EnsureVisible()
			}
			else
			{
				Write-Host "INFO: Cannot move item any higher..." -foreground "Yellow"
			}
		}
	}

}


function Move-Down
{
	$checkResult = CheckSkypeForBusinessOnline
	if($checkResult)
	{
		foreach ($lvi in $lv.SelectedItems)
		{
			#GET SETTINGS OF SELECTED ITEM
			$item = $lv.Items[$lvi.Index]
			$itemValue = $item.SubItems

			[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
			[string]$Name = $item.Text
			#[string]$Priority = $itemValue[1].Text
			[string]$Description = $itemValue[1].Text
			[string]$Pattern = $itemValue[2].Text
			[string]$Translation = $itemValue[3].Text
			[bool]$ExtensionValue = $itemValue[4].Text
			#$ExtensionValue = $ExtensionValue.ToLower()
												
			
			$orgIndex = $lvi.Index
			if($orgIndex -lt ($lv.Items.Count - 1))
			{
				$index = $orgIndex + 1
				
				Write-Host "RUNNING: `$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index -IsInternalExtension $ExtensionValue -InMemory" -foreground "green"
				$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index -IsInternalExtension $ExtensionValue  -InMemory			
				
				$NormArray = (Get-CsTenantDialPlan -identity $Scope | select NormalizationRules).NormalizationRules
				Write-Verbose "INITIAL ARRAY"
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host
				#Remove Item
				$NormArray.RemoveAt($orgIndex)
				Write-Verbose "AFTER DELETE"
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host 
				#Insert Item
				$NormArray.Insert($index,$nr)
				Write-Verbose "AFTER INSERT"
				foreach($item in $NormArray){Write-Verbose $item}
				
				Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Replace=$NormArray}
				
				GetNormalisationPolicy
				
				$lv.Items[$index].Selected = $true
				$lv.Items[$index].EnsureVisible()
			}
			else
			{
				Write-Host "INFO: Cannot move item any lower..." -foreground "Yellow"
			}
		}
	}
}

function GetNormalisationPolicy
{
	$lv.Items.Clear()
	
	$theIdentity = $policyDropDownBox.SelectedItem.ToString()
	Write-Host "INFO: Getting rules for $theIdentity" -foreground "yellow"
	$checkResult = CheckSkypeForBusinessOnline
	if($checkResult)
	{
		$TenantDialPlan = Get-CsTenantDialPlan -identity $theIdentity
		$NormRules = ($TenantDialPlan).NormalizationRules
		$AccessPrefixTextBox.Text = $TenantDialPlan.ExternalAccessPrefix
		$OptimizeDeviceDialingLabel.Text = "OptimizeDeviceDialing: " + $TenantDialPlan.OptimizeDeviceDialing
	}
	
	
	foreach($NormRule in $NormRules)
	{
		$Name = $NormRule.Name
		$Priority = "" #$NormRule.Priority
		$Description = $NormRule.Description
		if($Description -eq $null)
		{
			$Description = "<Not Set>"
		}
		$Pattern = $NormRule.Pattern
		$Tranlation = $NormRule.Translation
		$Extension = $NormRule.IsInternalExtension.ToString()
		
		$lvItem = new-object System.Windows.Forms.ListViewItem($Name)
		$lvItem.ForeColor = "Black"
		
		#[void]$lvItem.SubItems.Add($Priority)
		[void]$lvItem.SubItems.Add($Description)
		[void]$lvItem.SubItems.Add($Pattern)
		[void]$lvItem.SubItems.Add($Tranlation)
		[void]$lvItem.SubItems.Add($Extension)
		
		[void]$lv.Items.Add($lvItem)
	}
}

function UpdateListViewSettings
{
	if($lv.SelectedItems.count -eq 0)
	{
		$NameTextBox.Text = ""
		$DescriptionTextBox.Text = ""
		$PatternTextBox.Text = ""
		$TranslationTextBox.Text = ""
		$ExtensionCheckBox.Checked = $false
	}
	else
	{
		foreach ($item in $lv.SelectedItems)
		{
			[string]$itemName = $item.Text
			$itemValue = $item.SubItems
			
			$NameTextBox.Text = $itemName
			
			[string]$settingValue1 = $itemValue[1].Text
			$DescriptionTextBox.Text = $settingValue1
			[string]$settingValue3 = $itemValue[2].Text
			$PatternTextBox.Text = $settingValue3
			[string]$settingValue4 = $itemValue[3].Text
			$TranslationTextBox.Text = $settingValue4
			[string]$settingValue4 = $itemValue[4].Text
			$settingValue4 = $settingValue4.ToLower()
			if($settingValue4 -eq "true")
			{$ExtensionCheckBox.Checked = $true}
			else
			{$ExtensionCheckBox.Checked = $false}
			
		}
	}
}


#Add / Edit an item
function AddSetting
{
	[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
	[string]$Name = $NameTextBox.Text
	[string]$Description = $DescriptionTextBox.Text
	[string]$Pattern = $PatternTextBox.Text
	[string]$Translation = $TranslationTextBox.Text
	$ExtensionBool = $ExtensionCheckBox.Checked
	
	if($Scope -ne "" -and $Scope -ne $null -and $Name -ne "" -and $Name -ne $null -and $Pattern -ne "" -and $Pattern -ne $null -and $Translation -ne "" -and $Translation -ne $null)
	{
		$checkResult = CheckSkypeForBusinessOnline
		if($checkResult)
		{
			[string]$Name = $NameTextBox.Text
			$EditSetting = $false
			$LoopNo = 0
			foreach($item in $lv.Items)
			{
				[string]$listName = $item.Text
				if($listName.ToLower() -eq $Name.ToLower())
				{
					$EditSetting = $true
					$Priority = $LoopNo
					break
				}
				$LoopNo++
			}
			if($EditSetting)
			{
				Write-Host "INFO: Name is already in the list. Editing setting" -foreground "yellow"
				
				Write-Verbose "SELECTED INDEX: $($lv.SelectedIndices[0])" 
				$orgIndex = $lv.SelectedIndices[0] #$lvi.Index
				$NewIndex = $orgIndex - 1
							
				$NormArray = (Get-CsTenantDialPlan -identity $Scope | select NormalizationRules).NormalizationRules
				foreach($item in $NormArray){Write-Verbose $item}
				#Remove Item
				$NormArray.RemoveAt($orgIndex)
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host 
				#Insert Item
				Write-Host "RUNNING: `$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -IsInternalExtension $ExtensionBool -InMemory" -foreground "green"
				$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -IsInternalExtension $ExtensionBool -InMemory
				$NormArray.Insert($orgIndex,$nr)
				foreach($item in $NormArray){Write-Verbose $item}
				
				Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Replace=$NormArray}
							
				GetNormalisationPolicy
				
				$lv.Items[$Priority].Selected = $true
				$lv.Items[$Priority].EnsureVisible()
			}
			else   # ADD
			{
				
				Write-Host "RUNNING: `$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -IsInternalExtension $ExtensionBool -InMemory" -foreground "green"
				$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -IsInternalExtension $ExtensionBool -InMemory
				Write-Host "RUNNING: New-CsTenantDialPlan -Identity ${Scope} -NormalizationRules @{Add=`$nr}" -foreground "green"
				Set-CsTenantDialPlan -Identity ${Scope} -NormalizationRules @{Add=$nr}
				
				GetNormalisationPolicy
				
				$count = $lv.Items.Count - 1
				$lv.Items[$count].Selected = $true
				$lv.Items[$count].EnsureVisible()
				
			}
		}
	}
	else
	{
		Write-Host "ERROR: Please enter values for Name, Pattern and Translation." -foreground "red"
	}
}


function DeleteSetting
{
	$checkResult = CheckSkypeForBusinessOnline
	if($checkResult)
	{
		if($lv.SelectedItems.Count -le 0)
		{Write-Host "ERROR: No items selected to delete. Please select items before selecting delete." -foreground "red"}
		foreach ($lvi in $lv.SelectedItems)
		{
			#GET SETTINGS OF SELECTED ITEM
			$item = $lv.Items[$lvi.Index]
			$itemValue = $item.SubItems

			[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
			[string]$Scope = $Scope.Replace("site:","")
			[string]$Name = $item.Text
			#[string]$Priority = $itemValue[1].Text
			[string]$Description = $itemValue[1].Text
			[string]$Pattern = $itemValue[2].Text
			[string]$Translation = $itemValue[3].Text
			
			$orgIndex = $lvi.Index
			
			Write-Host "INFO: Removing - ${Scope}/${Name}" -foreground "Yellow"
			
			Write-Host "RUNNING: (Get-CsTenantDialPlan -identity ${Scope}).NormalizationRules | Where-Object {$_.Name -eq ${Name}}" -foreground "green"
			$policyItem = (Get-CsTenantDialPlan -identity ${Scope}).NormalizationRules | Where-Object {$_.Name -eq "${Name}"}
			Write-Host "RUNNING: Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Remove=$policyItem}" -foreground "green"
			Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Remove=$policyItem}
			
			GetNormalisationPolicy
			
			if($orgIndex -ge 1)
			{
				$index = $orgIndex - 1
				
				$lv.Items[$index].Selected = $true
				$lv.Items[$index].EnsureVisible()
			}
			else
			{
				$lv.Items[0].Selected = $true
				$lv.Items[0].EnsureVisible()
			}
			
			UpdateListViewSettings
		}
	}
}

function DeleteAllSettings
{
	$checkResult = CheckSkypeForBusinessOnline
	if($checkResult)
	{
		[string]$Scope = $policyDropDownBox.SelectedItem.ToString()

		Write-Host "RUNNING: (Get-CsTenantDialPlan -identity ${Scope}).NormalizationRules" -foreground "green"
		$policyItem = (Get-CsTenantDialPlan -identity ${Scope}).NormalizationRules #| Where-Object {$_.Name -eq "${Name}"}
		
		Write-Host "RUNNING: Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Remove=$policyItem}" -foreground "green"
		Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Remove=$policyItem}
		
		GetNormalisationPolicy
	}
}


function TestPhoneNumberNew()
{

	$TestPhoneResultTextLabel.Text = "Test Result: No Match"
	$TestPhonePatternTextLabel.Text = "Matched Pattern: No Match"
	$TestPhoneTranslationTextLabel.Text = "Matched Translation: No Match"
	
	foreach($tempitem in $lv.Items)
	{
		$tempitem.ForeColor = "Black"
	}
	$PhoneNumber = $TestPhoneTextBox.Text
	#$Rules = Get-CsAddressBookNormalizationRule
	
	Write-Host ""
	Write-Host "-------------------------------------------------------------" -foreground "Green"
	Write-Host "TESTING: $PhoneNumber" -foreground "Green"
	Write-Host ""

	$TopLoopNo = 0
	$firstFound = $true
	foreach($item in $lv.Items)
	{
		$itemValue = $item.SubItems

		[string]$Pattern = $itemValue[2].Text
		[string]$Translation = $itemValue[3].Text
		
		#Clean up the Phone Number
		$PhoneNumberStripped = $PhoneNumber.Replace(" ","").Replace("(","").Replace(")","").Replace("[","").Replace("]","").Replace("{","").Replace("}","").Replace(".","").Replace("-","").Replace(":","")
		
		Write-Verbose "TESTING PATTERN: $Pattern" #DEBUG
		
		$PatternStartEnd = "^$Pattern$"
		Try
		{
			$StartPatternResult = $PhoneNumberStripped -cmatch $PatternStartEnd
		}
		Catch
		{
			#This error was already reported. So don't bother reporting it again.
		}
		
		if($StartPatternResult)
		{
			if($firstFound)
			{
				Write-Host "First Matched Pattern: $Pattern" -foreground "Green"
				Write-Host "First Matched Translation: $Translation" -foreground "Green"
				$TestPhonePatternTextLabel.Text = "Matched Pattern: $Pattern"
				$TestPhoneTranslationTextLabel.Text = "Matched Translation: $Translation"
				
				$Group1 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[1].Value
				$Group2 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[2].Value
				$Group3 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[3].Value
				$Group4 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[4].Value
				$Group5 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[5].Value
				$Group6 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[6].Value
				$Group7 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[7].Value
				$Group8 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[8].Value
				$Group9 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[9].Value
				
				Write-Host
				if($Group1 -ne ""){Write-Host "Group 1: " $Group1 -foreground "Yellow"}
				if($Group2 -ne ""){Write-Host "Group 2: " $Group2 -foreground "Yellow"}
				if($Group3 -ne ""){Write-Host "Group 3: " $Group3 -foreground "Yellow"}
				if($Group4 -ne ""){Write-Host "Group 4: " $Group4 -foreground "Yellow"}
				if($Group5 -ne ""){Write-Host "Group 5: " $Group5 -foreground "Yellow"}
				if($Group6 -ne ""){Write-Host "Group 6: " $Group6 -foreground "Yellow"}
				if($Group7 -ne ""){Write-Host "Group 7: " $Group7 -foreground "Yellow"}
				if($Group8 -ne ""){Write-Host "Group 8: " $Group8 -foreground "Yellow"}
				if($Group9 -ne ""){Write-Host "Group 9: " $Group9 -foreground "Yellow"}
				
				Write-Host				
				$Result = $Translation.Replace('$1',"$Group1")
				$Result = $Result.Replace('$2',"$Group2")
				$Result = $Result.Replace('$3',"$Group3")
				$Result = $Result.Replace('$4',"$Group4")
				$Result = $Result.Replace('$5',"$Group5")
				$Result = $Result.Replace('$6',"$Group6")
				$Result = $Result.Replace('$7',"$Group7")
				$Result = $Result.Replace('$8',"$Group8")
				$Result = $Result.Replace('$9',"$Group9")
				Write-Host "Result: " $Result -foreground "Green"
				$TestPhoneResultTextLabel.Text = "Test Result: ${Result}"
				
				$firstFound = $false
				$item.ForeColor = "Green"
			}
			else
			{
				$item.ForeColor = "Blue"
			}
		}
	}
	$lv.SelectedItems.Clear()
	$lv.Focus()	
	Write-Host "-------------------------------------------------------------" -foreground "Green"
}


function Import-Config
{
	$Filter = "All Files (*.*)|*.*"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$fileForm = New-Object System.Windows.Forms.OpenFileDialog
	$fileForm.InitialDirectory = $pathbox.text
	$fileForm.Filter = $Filter
	$fileForm.Title = "Open File"
	$Show = $fileForm.ShowDialog()
	if ($Show -eq "OK")
	{
		#IMPORT CODE
	}
	else
	{
		Write-Host "INFO: Operation cancelled by user." -foreground "yellow"
	}
	
}

function Export-Config
{
	#File Dialog
	[string] $pathVar = "C:\"
	$Filter="All Files (*.*)|*.*"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objDialog = New-Object System.Windows.Forms.SaveFileDialog
	#$objDialog.InitialDirectory = 
	$objDialog.FileName = "Company_Phone_Number_Normalization_Rules.txt"
	$objDialog.Filter = $Filter
	$objDialog.Title = "Export File Name"
	$objDialog.CheckFileExists = $false
	$Show = $objDialog.ShowDialog()
	if ($Show -eq "OK")
	{
		[string]$outputFile = $objDialog.FileName
		$outputFile = "${outputFile}"
		$output = ""
		foreach($item in $lv.Items)
		{
			$itemValue = $item.SubItems

			[string]$Pattern = $itemValue[2].Text
			[string]$Translation = $itemValue[3].Text
			[string]$Name = $item.Text
			[string]$Description = $itemValue[1].Text
			
			$output += "# $Name $Description`r`n$Pattern`r`n$Translation`r`n`r`n"
		}
		
		$output | out-file -Encoding UTF8 -FilePath $outputFile -Force					
		Write-Host "Written File to $outputFile...." -foreground "yellow"
	}
	else
	{
		return
	}
	
}


$result = CheckSkypeForBusinessOnlineInitial


# Activate the form ============================================================
$mainForm.Add_Shown({$mainForm.Activate()})
[void] $mainForm.ShowDialog()	


# SIG # Begin signature block
# MIIcZgYJKoZIhvcNAQcCoIIcVzCCHFMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUSqq9tV78jQx25u75MI0BF8ZJ
# oiyggheVMIIFHjCCBAagAwIBAgIQDGWW2SJRLPvqOO0rxctZHTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE5MDIwNjAwMDAwMFoXDTIwMDIw
# NjEyMDAwMFowWzELMAkGA1UEBhMCQVUxDDAKBgNVBAgTA1ZJQzEQMA4GA1UEBxMH
# TWl0Y2hhbTEVMBMGA1UEChMMSmFtZXMgQ3Vzc2VuMRUwEwYDVQQDEwxKYW1lcyBD
# dXNzZW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDHPwqNOkuXxh8T
# 7y2cCWgLtpW30x/3rEUFnrlCv2DFgULLfZHFTd+HWhCiTUMHVESj+X8s+cmgKVWN
# bmEWPri590V6kfUmjtC+4/iKdVpvjgwrwAm6O6lHZ91y4Sn90A7eUV/EvUmGREVx
# uFk2s7jD/cYjTzm0fACQBuPz5sVjTzgFzbZMndPcptB8uEjtIF/k6BGCy7XyAMn6
# 0IncNguxGZBsS/CQQlsXlVhTnBn0QQxa7nRcpJQs/84OXjDypgjW6gVOf3hOzfXY
# rXNR54nqIh/VKFKz+PiEIW11yLW0608cI0xEE03yBOg14NGIapNBwOwSpeLMlQbH
# c9twu9BhAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg+S32
# ZXUOWDAdBgNVHQ4EFgQU2P05tP7466o6clrA//AUqWO4b2swDgYDVR0PAQH/BAQD
# AgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWgM6Ax
# hi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNy
# bDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYBBQUHAQEEeDB2
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYBBQUH
# MAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1
# cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEB
# CwUAA4IBAQCdaeq4xJ8ISuvZmb+ojTtfPN8PDWxDIsWos6e0KJ4sX7jYR/xXiG1k
# LgI5bVrb95YQNDIfB9ZeaVDrtrhEBu8Z3z3ZQFcwAudIvDyRw8HQCe7F3vKelMem
# TccwqWw/UuWWicqYzlK4Gz8abnSYSlCT52F8RpBO+T7j0ZSMycFDvFbfgBQk51uF
# mOFZk3RZE/ixSYEXlC1mS9/h3U9o30KuvVs3IfyITok4fSC7Wl9+24qmYDYYKh8H
# 2/jRG2oneR7yNCwUAMxnZBFjFI8/fNWALqXyMkyWZOIgzewSiELGXrQwauiOUXf4
# W7AIAXkINv7dFj2bS/QR/bROZ0zA5bJVMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1
# U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQD
# ExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcN
# MjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2Vy
# dCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid
# 2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sj
# lOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjf
# DPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzL
# fnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR
# 93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckw
# EgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgw
# OqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIE
# MCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYI
# YIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQY
# MBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1a
# JLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUP
# UbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1
# UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjF
# Emifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM
# 1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhs
# RDKyZqHnGKSaZFHvMIIGajCCBVKgAwIBAgIQAwGaAjr/WLFr1tXq5hfwZjANBgkq
# hkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBB
# c3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAwWhcNMjQxMDIyMDAwMDAwWjBH
# MQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxJTAjBgNVBAMTHERpZ2lD
# ZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBTqZ8fZFnmfGt/a4ydVfiS457V
# WmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWRn8YUOawk6qhLLJGJzF4o9GS2
# ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRVfRiGBYxVh3lIRvfKDo2n3k5f
# 4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3vJ+P3mvBMMWSN4+v6GYeofs/s
# jAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA8bLOcEaD6dpAoVk62RUJV5lW
# MJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGjggM1MIIDMTAOBgNVHQ8BAf8E
# BAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCCAb8G
# A1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIBkjAoBggrBgEFBQcCARYcaHR0
# cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQGCCsGAQUFBwICMIIBVh6CAVIA
# QQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMA
# YQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4A
# YwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAA
# UwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAA
# QQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkA
# YQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIA
# YQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUA
# LjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQASKxOYspkH7R7for5XDStnAs0w
# HQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9MH0GA1UdHwR2MHQwOKA2oDSG
# Mmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEu
# Y3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcnQwDQYJKoZIhvcN
# AQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI//+x1GosMe06FxlxF82pG7xa
# FjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7easGAm6mlXIV00Lx9xsIOUGQVr
# NZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8OxwYtNiS7Dgc6aSwNOOMdgv420X
# Ewbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQNJsQOfxu19aDxxncGKBXp2JPl
# VRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNtomHpigtt7BIYvfdVVEADkitr
# wlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbNMIIFtaADAgECAhAG/fkDlgOt
# 6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0wNjExMTAwMDAwMDBa
# Fw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/JM/xNRZFcgZ/tLJz4FlnfnrUk
# FcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPsi3o2CAOrDDT+GEmC/sfHMUiA
# fB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ8DIhFonGcIj5BZd9o8dD3QLo
# Oz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNugnM/JksUkK5ZZgrEjb7Szgau
# rYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJrGGWxwXOt1/HYzx4KdFxCuGh+
# t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3owggN2MA4GA1UdDwEB/wQEAwIB
# hjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUHAwIGCCsGAQUFBwMDBggrBgEF
# BQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIBxTCCAbQGCmCGSAGG/WwAAQQw
# ggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9zc2wtY3Bz
# LXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUA
# cwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMA
# bwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYA
# IAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQA
# IAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUA
# bQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkA
# dAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAA
# aABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMAsGCWCGSAGG
# /WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsNC5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMB0GA1UdDgQW
# BBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYun
# pyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+ybcoJKc4HbZbKa9Sz1LpMUer
# Vlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6hnKtOHisdV0XFzRyR4WUVtHr
# uzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5PsQXSDj0aqRRbpoYxYqioM+Sb
# OafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke/MV5vEwSV/5f4R68Al2o/vsH
# OE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qquAHzunEIOz5HXJ7cW7g/DvXwK
# oO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQnHcUwZ1PL1qVCCkQJjGCBDsw
# ggQ3AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNI
# QTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0ECEAxlltkiUSz76jjtK8XLWR0w
# CQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcN
# AQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
# IwYJKoZIhvcNAQkEMRYEFFPdhuNE4YPBkgY5jllH9P0+vpkvMA0GCSqGSIb3DQEB
# AQUABIIBAJ35NG4tIDSjat+dr/bfNHup8Ftvh3FNI76vSIemMUeEXmd01pQKrMXG
# s/cv4Q5YpQcTLWUHJkuBOAWFl7SkTbsUMkMoCPVMhNkWheTwox59uGUe2HM3rlML
# XomJxG59P41hhuRpXpL4YcknRydOYXnTBdHNrMvNdCojX0pqpMzzZjoHsABgNoF8
# yRGadce530w/Q/EZFLrIgJpdz1e8hXlnXTTZwOzVaTIo78P+93zowJdH7wYjv0Hv
# uJR/v4CuapVEnWRZW4+xqTo7IPhxvbdafO/Jbsukr36DC2K4BENdkaMHNC8N1E0s
# AhENOfCBa6D0cnUzJT7OR1GP6QNB5sKhggIPMIICCwYJKoZIhvcNAQkGMYIB/DCC
# AfgCAQEwdjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkw
# FwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1
# cmVkIElEIENBLTECEAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE5MDkwNTExMzA1
# NlowIwYJKoZIhvcNAQkEMRYEFPihWMPSe7n5COQeqJqbtr4ulHSsMA0GCSqGSIb3
# DQEBAQUABIIBAD6XXC/4euv8STKaJMlg4KYayrM2qfmCVvuHAhRALQ3FV6sWXlPi
# ouTRBqnr/EkU0ZohEVBxf7d1IDCdxoDz4K+BVWSganwOd6d9lNdBvCxEu3FAeGp9
# Gm6ts0fBW9jYibjQp/rA8jPLtQJAU8guEZ6gNpYOnv9AIBgyUvOlgO9DMEr7PHYX
# WEophcQKZwuKhbFc664SPEY4PtHQ+dTD8874gX+4ohp8Mf/O54LS2s/pRHrh6mpk
# 5K7OFGItZq60D5jZ+Czyi9NoeuSqdtq3+N4aJIZZlZ3BWgo80vOcaNmLdxsOzzsU
# z5M97shXFRzCwjAq4z7e7H4T4K9FYX83bCg=
# SIG # End signature block
