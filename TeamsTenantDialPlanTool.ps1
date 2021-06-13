########################################################################
# Name: Teams Tenant Dial Plan Tool 
# Version: v1.01 (13/6/2021)
# Date: 1/9/2019
# Created By: James Cussen
# Web Site: http://www.myteamslab.com
# 
# Notes: This is a PowerShell tool. To run the tool, open it from the PowerShell command line on a PC that has the MicrosoftTeams PowerShell module installed. Get it by opening a PowerShell window using Run as Administrator and running "Install-Module MicrosoftTeams -AllowClobber"
#		 For more information on the requirements for setting up and using this tool please visit http://www.myteamslab.com.
#
# Copyright: Copyright (c) 2021, James Cussen (www.myteamslab.com) All rights reserved.
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
#
# 1.01 Teams Module Update
#	- The Skype for Business PowerShell module is being deprecated and the Teams Module is finally good enough to use with this tool. As a result, this tool has now been updated for use with the Teams PowerShell Module version 2.3.1 or above.
#
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
elseif(([int]$MajorVersion) -ge  6)
{
	Write-Host "This machine has version $MajorVersion PowerShell installed. This version uses .NET Core which doesn't support Windows Forms. Please use PowerShell 5 instead." -foreground "red"
	exit
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

Function Get-MyModule 
{ 
Param([string]$name) 
	
	if(-not(Get-Module -name $name)) 
	{ 
		if(Get-Module -ListAvailable | Where-Object { $_.name -eq $name }) 
		{ 
			Import-Module -Name $name 
			return $true 
		} #end if module available then import 
		else 
		{ 
			return $false 
		} #module not available 
	} # end if not module 
	else 
	{ 
		return $true 
	} #module already loaded 
} #end function get-MyModule

$Script:TeamsModuleAvailable = $false

Write-Host "--------------------------------------------------------------" -foreground "green"
Write-Host "Checking for PowerShell Modules..." -foreground "green"
#Import MicrosoftTeams Module
if(Get-MyModule "MicrosoftTeams")
{
	#Invoke-Expression "Import-Module Lync"
	Write-Host "INFO: Teams module should be at least 2.3.1" -foreground "yellow"
	$version = (Get-Module -name "MicrosoftTeams").Version
	Write-Host "INFO: Your current version of Teams Module: $version" -foreground "yellow"
	if([System.Version]$version -ge [System.Version]"2.3.1")
	{
		Write-Host "Congratulations, your version is acceptable!" -foreground "green"
	}
	else
	{
		Write-Host "ERROR: You need to update your Teams Version to higher than 2.3.1. Use the command Update-Module MicrosoftTeams" -foreground "red"
		exit
	}
	Write-Host "Found MicrosoftTeams Module..." -foreground "green"
	$Script:TeamsModuleAvailable = $true
}
else
{
	Write-Host "ERROR: You do not have the Microsoft Teams Module installed. Get it by opening a PowerShell window using `"Run as Administrator`" and running `"Install-Module MicrosoftTeams -AllowClobber`"" -foreground "red"
	#Can't find module so exit
	exit
}

Write-Host "--------------------------------------------------------------" -foreground "green"



# Set up the form  ============================================================

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$mainForm = New-Object System.Windows.Forms.Form 
$mainForm.Text = "Teams Tenant Dial Plan Tool 1.01"
$mainForm.Size = New-Object System.Drawing.Size(525,680) 
$mainForm.MinimumSize = New-Object System.Drawing.Size(520,450) 
$mainForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$mainForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$mainForm.KeyPreview = $True
$mainForm.TabStop = $false


$global:SFBOsession = $null
#ConnectButton
$ConnectOnlineButton = New-Object System.Windows.Forms.Button
$ConnectOnlineButton.Location = New-Object System.Drawing.Size(20,7)
$ConnectOnlineButton.Size = New-Object System.Drawing.Size(110,20)
$ConnectOnlineButton.Text = "Connect Teams"
$ConnectTooltip = New-Object System.Windows.Forms.ToolTip
$ConnectToolTip.SetToolTip($ConnectOnlineButton, "Connect/Disconnect from Teams")
#$ConnectButton.tabIndex = 1
$ConnectOnlineButton.Enabled = $true
$ConnectOnlineButton.Add_Click({	

	$ConnectOnlineButton.Enabled = $false
	
	$StatusLabel.Text = "STATUS: Connecting to Teams..."
	
	if($ConnectOnlineButton.Text -eq "Connect Teams")
	{
		ConnectTeamsModule
		[System.Windows.Forms.Application]::DoEvents()
		CheckTeamsOnline
	}
	elseif($ConnectOnlineButton.Text -eq "Disconnect Teams")
	{	
		$ConnectOnlineButton.Text = "Disconnecting..."
		$StatusLabel.Text = "STATUS: Disconnecting from Teams..."
		DisconnectTeams
		CheckTeamsOnline
	}
	
	$ConnectOnlineButton.Enabled = $true
	
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($ConnectOnlineButton)


$ConnectedOnlineLabel = New-Object System.Windows.Forms.Label
$ConnectedOnlineLabel.Location = New-Object System.Drawing.Size(135,10) 
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
$MyLinkLabel.Text = "www.myteamslab.com"
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("http://www.myteamslab.com")
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



function ConnectTeamsModule
{
	$ConnectOnlineButton.Text = "Connecting..."
	$StatusLabel.Text = "Connecting to Microsoft Teams..."
	Write-Host "INFO: Connecting to Microsoft Teams..." -foreground "Yellow"
	[System.Windows.Forms.Application]::DoEvents()
	
	if (Get-Module -ListAvailable -Name MicrosoftTeams)
	{
		Import-Module MicrosoftTeams
		$cred = Get-Credential
		if($cred)
		{
			try
			{
				(Connect-MicrosoftTeams -Credential $cred) 2> $null
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
				
				EnableAllButtons
				
				$ConnectOnlineButton.Text = "Disconnect Teams"
				
				return $true
			}
			catch
			{
				if($_.Exception -match "you must use multi-factor authentication to access" -or $_.Exception -match "The security token could not be authenticated or authorized") #MFA FALLBACK!
				{
					try
					{
						(Connect-MicrosoftTeams) 2> $null
						
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
						
						EnableAllButtons
						
						$ConnectOnlineButton.Text = "Disconnect Teams"
					
						return $true
					}
					catch
					{
						if($_.Exception -match "User canceled authentication")
						{
							Write-Host "INFO: Canceled authentication." -foreground "yellow"
							DisableAllButtons
							return $false
						}
						else
						{
							Write-Host "ERROR: " $_.Exception -foreground "red"
							DisableAllButtons
							return $false
						}
					}
				}
				elseif($_.Exception -match "Error validating credentials due to invalid username or password.")
				{
					Write-Host "ERROR: Error validating credentials due to invalid username or password." -foreground "red"
					DisableAllButtons
					return $false
				}
				else
				{
					Write-Host "ERROR: " $_.Exception -foreground "red"
					DisableAllButtons
					return $false	
				}
				
				Write-Host "ERROR: " $_.Exception -foreground "red"
				DisableAllButtons
				return $false	
			}
		}
	}
}

function DisconnectTeams
{
	Write-Host "RUNNING: Disconnect-MicrosoftTeams" -foreground "Green"
	$disconnectResult = Disconnect-MicrosoftTeams
	Write-Host "RUNNING: Remove-Module MicrosoftTeams" -foreground "Green"
	Remove-Module MicrosoftTeams
	
	Write-Host "RUNNING: Get-Module -ListAvailable -Name MicrosoftTeams" -foreground "Green"
	$result = Invoke-Expression "Get-Module -ListAvailable -Name MicrosoftTeams"
	if($result -ne $null)
	{
		Write-Host "MicrosoftTeams has been removed successfully" -foreground "Green"
	}
	else
	{
		Write-Host "ERROR: MicrosoftTeams was not removed." -foreground "red"
	}
	
	$ConnectOnlineButton.Text = "Connect Teams"
	DisableAllButtons
}



function CheckTeamsOnlineInitial
{	
	#CHECK IF COMMANDS ARE AVAILABLE		
	$command = "Get-CsOnlineUser"
	#if($CurrentlyConnected -and (Get-Command $command -errorAction SilentlyContinue) -and ($Script:UserConnectedToTeamsOnline -eq $true))
	if((Get-Command $command -errorAction SilentlyContinue))
	{
		$isConnected = $false
		try{
			(Get-CsOnlineUser -ResultSize 1 -ErrorAction SilentlyContinue) 2> $null
			$isConnected = $true
		}
		catch
		{
			#Write-Host "ERROR: " $_ -foreground "red"
			$isConnected = $false
		}
		#CHECK THAT SfB ONLINE COMMANDS WORK
		if($isConnected)
		{
			#Write-Host "Connected to Teams" -foreground "Green"
			$ConnectedOnlineLabel.Visible = $true
			$ConnectOnlineButton.Text = "Disconnect Teams"
			$StatusLabel.Text = ""

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
			
			EnableAllButtons
			
			return $true
		}
		else
		{
			Write-Host "INFO: Cannot access Teams. Please use the Connect Teams button." -foreground "Yellow"
			$ConnectedOnlineLabel.Visible = $false
			$ConnectOnlineButton.Text = "Connect Teams"
			$StatusLabel.Text = "Press the `"Connect Teams`" button to get started."
			
			DisableAllButtons
		}
	}
}


function CheckTeamsOnline
{	
	
	#CHECK IF COMMANDS ARE AVAILABLE		
	$isConnected = $false
	try{
		(Get-CsOnlineUser -ResultSize 1 -ErrorAction SilentlyContinue) 2> $null
		$isConnected = $true
	}
	catch
	{
		#Write-Host "ERROR: " $_ -foreground "red"
		$isConnected = $false
	}
	#CHECK THAT SfB ONLINE COMMANDS WORK
	if($isConnected)
	{
		$ConnectedOnlineLabel.Visible = $true
		$ConnectOnlineButton.Text = "Disconnect Teams"
		#$StatusLabel.Text = ""
		return $true
		
	}
	else
	{
		Write-Host "INFO: Cannot access Teams. Please use the Connect Teams button." -foreground "Yellow"
		$ConnectedOnlineLabel.Visible = $false
		$ConnectOnlineButton.Text = "Connect Teams"
		$StatusLabel.Text = "Press the `"Connect Teams`" button to get started."
		
		DisableAllButtons
	}
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
	
	$checkResult = CheckTeamsOnline
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
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
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
	
	$checkResult = CheckTeamsOnline
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
	
	$checkResult = CheckTeamsOnline
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
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
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
	$checkResult = CheckTeamsOnline
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
	$checkResult = CheckTeamsOnline
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
	$checkResult = CheckTeamsOnline
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
		$checkResult = CheckTeamsOnline
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
	$checkResult = CheckTeamsOnline
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
	$checkResult = CheckTeamsOnline
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


$result = CheckTeamsOnlineInitial


# Activate the form ============================================================
$mainForm.Add_Shown({$mainForm.Activate()})
[void] $mainForm.ShowDialog()	


# SIG # Begin signature block
# MIIZlgYJKoZIhvcNAQcCoIIZhzCCGYMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUijbMFj+rLz9Y1vg/JDzPiHZU
# rD+gghSkMIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTAwggQY
# oAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsx
# SRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawO
# eSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJ
# RdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEc
# z+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whk
# PlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8l
# k9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQD
# AgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARI
# MEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdp
# Y2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg
# +S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG
# 9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/E
# r4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3
# nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpo
# aK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW
# 6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ
# 92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTEwggQZoAMCAQICEAqhJdbW
# Mht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIG
# A1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTE2MDEwNzEyMDAw
# MFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnFOVQoV7YjSsQOB0Uz
# URB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQAOPcuHjvuzKb2Mln+
# X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhisEeTwmQNtO4V8CdPu
# XciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQjMF287DxgaqwvB8z9
# 8OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+fMRTWrdXyZMt7HgXQ
# hBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW/5MCAwEAAaOCAc4w
# ggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEFBQcBAQRtMGswJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBQBgNV
# HSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cu
# ZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggEB
# AHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafDDiBCLK938ysfDCFa
# KrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6HHssIeLWWywUNUME
# aLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4H9YLFKWA1xJHcLN1
# 1ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHKeZR+WfyMD+NvtQEm
# tmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIoxhhWz0E0tmZdtnR7
# 9VYzIi8iNrJLokqV2PWmjlIwggU1MIIEHaADAgECAhAKNIchv70WQdoZqmZoB0Fg
# MA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwMTA2MDAw
# MDAwWhcNMjMwMTEwMTIwMDAwWjByMQswCQYDVQQGEwJBVTEMMAoGA1UECBMDVklD
# MRAwDgYDVQQHEwdNaXRjaGFtMRUwEwYDVQQKEwxKYW1lcyBDdXNzZW4xFTATBgNV
# BAsTDEphbWVzIEN1c3NlbjEVMBMGA1UEAxMMSmFtZXMgQ3Vzc2VuMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzerRQ8lBU+cD9jWzQV7i2saGNzXrXzaN
# kEUkUdZbem54qullpGQwp6Bb0hzEFsPIaPSd796kIvvQCdb2W6VM9zp5ZxZj8dIh
# 539for2NW7Av8kjj+qpq+geD7BGWhLXKdMRdfdVZgf9hgWi+FOv+bJHp5MCKi9pN
# WEi8mgvaRZd2FuGJ7+RlYpYhGamYNw9KaV32/T9t2Mm7b9As1jlss+/Zja+Jsb5R
# pDFfhSX5eKG1Fy8T0QnaEvJm0Ljr2KD2E9AAmB96ZalNuwhqPociEUflTUyrmSlY
# w9HxFZ6cWXvHidcXnFW9exHpasXC2agwxYzYs+FqobL6cDw258kidQIDAQABo4IB
# xTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYE
# FP58C0FrWPjmUX3IhXKbtjOP8UHQMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAK
# BggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQu
# ZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3
# BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQu
# Y29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNpZ25p
# bmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAfBTJIJek
# j0gumY4Ej7cbSHLnbMh0hrDXrAaUxFadLJgKMCl8N0YOuR/5Vw4voCvgWuFv1dVO
# ns7lZzu/Y9T/kPqNpxVzxLO6jZDN3zEPmpt2E1nqelL3BdBF0eEK8i8mEkrFdi8Q
# gT1VhqjeROCLKUm8N928wM3iBEjH9pFyQlBDNHFgiFt9H/NXhFJ5IfC8yDzbt7a/
# 9hVwtcWMWygxvSKjL6pCTAXBXPWajiU+ddcV6VRs3QuRYsex0DGrABM1AcDXnRKZ
# OlLu2bhh7abbeWBWXCAaBHYmCFbPpspUj6eb5R8AI52+leeMEggPIw1SX21HHh6j
# rHLF9RJUBJDeSzGCBFwwggRYAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNV
# BAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0ECEAo0
# hyG/vRZB2hmqZmgHQWAwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKA
# AKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFLQJZe+DKdlzh9nQWV3QIpyf
# hV1lMA0GCSqGSIb3DQEBAQUABIIBAECVnapCN4xY2AWioMdxGINLs7KgymnHQsmM
# KpvM5CChQC5Ub1D+5EPSWcXVdlrnWXaQLxIMDYlIE/UObZixSjq7IzuqQqMCq1iN
# QdRWVi6i4KTu9ddGKc836yR5DmjqZpnwPhmfzbC6QuN49uumsSlvQ3h0qCq9CuNy
# IuxMHw3Zd0wrt5SGcehDcvEY5Snny4n6UWSGvo8RPOrF1AWTHnRg1GAvwA240wv7
# yV0Fr3TfqWH3IyRgmqWgk3vZOTOybgB1cx4A/Wzdv7S5kW3QvnFh+4dXoQ5PFPZa
# 6S/cixFX7LP4fPQOF1WTop0qQ5kv7YlyFgHLJJO38S3iau6Ii3OhggIwMIICLAYJ
# KoZIhvcNAQkGMYICHTCCAhkCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UE
# AxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQQIQDUJK
# 4L46iP9gQCHOFADw3TANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIxMDYxMzEzMTQzMFowLwYJKoZIhvcN
# AQkEMSIEIB94TaklQmMgOJjJfiNcdS9382D++4iAAvqTmyeXtkiEMA0GCSqGSIb3
# DQEBAQUABIIBAAauZI/vm2ke2aa+01G9lBoUikjCjSKIzEkg8I4eIZQ1jaLvYouy
# dGCKNVIZkvEQSH0oXN8Lv+AQkJoXbnJrHCzXDgFf9NVLvSrY6mdE2jWKCuQhVNWH
# EBCmLyNbeiTxiukjQfnfqvqTaGKO2P6+Ke9ALuywx8PwpJbeKbnxwWMJnxe2EAID
# f/UFVeF5a6mN5eJLVemO36VroewLs0VFh8ImiDSZxGpEbEaRVPvIXPrL/eGXm8mJ
# 9AVu06eYD0pcM+NiQ+R41uMB4mO+2d24UpOX6wA3opqDfVlwOYro2RYrlXyAo4AV
# nsgKxY7wRbTON2c95o3vlJB+74LoGi7bW/Q=
# SIG # End signature block
