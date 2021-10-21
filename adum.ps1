#####################################################################################################
### Helper
#####################################################################################################
### getting all parameters for align: System.Drawing.ContentAlignment
### $InputSearchLL | Get-Member -MemberType Property | Where-Object -FilterScript {$_.name -like "*ali*"}
#####################################################################################################
### Convert ico to base64
### [Convert]::ToBase64String((Get-Content "\\example.domain\files\scripts\ADUM\ico\adum16x16.ico" -Encoding Byte))
#####################################################################################################

#----------[Clear All variables]------------
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

#----------[Variables]------------
$span_width = 10
$frame_size = 20
$width = 1000
$height = 820
$window_width = $width - $frame_size

#------------[Logic/Script/Functions]------------
function enable_search_button {
	if ($this.Text -and $InputSearchTB.Text) {
		$SearchButton.Enabled = $true
	} else {
		clearall
		$SearchButton.Enabled = $False
		$returnStatus.Text = "Ready."
		$SaveButton.Enabled = $False
	}
}

function message($Level,$Text) {
    $date = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    if ($Level -like "info"){
        $OutputSearchRTB.Text = "[ INFO ] " + $date + " [" + $InputSearchTB.Text + "]: " + $Text + "`r`n" + $OutputSearchRTB.Text
    } elseif ($Level -like "error"){
        $OutputSearchRTB.Text = "[ ERROR ] " + $date + " [" + $InputSearchTB.Text + "]: " + $Text + "`r`n" + $OutputSearchRTB.Text
    } elseif ($Level -eq "ok"){
        $OutputSearchRTB.Text = "[ OK ] " + $date + " [" + $InputSearchTB.Text + "]: " + $Text + "`r`n" + $OutputSearchRTB.Text
    } elseif ($Level -eq "line"){
        $OutputSearchRTB.Text = "__________________________________________________`r`n" + $OutputSearchRTB.Text
    } elseif ($Level -eq "system"){
        $OutputSearchRTB.Text = "[ SYSTEM ] " + $date + "$Text" + "`r`n" + $OutputSearchRTB.Text
    } else {
        $OutputSearchRTB.Text = "[ DEBUG ] " + $date + "[" + $InputSearchTB.Text + "]: Something went wrong. Contact with Administrator." + "`r`n" + $OutputSearchRTB.Text
    }
}
function user_data {
    $returnStatus.Text = ("Loading... Please Wait.")
    $user = $null
    $user = Get-ADUser -Filter {sAMAccountName -like $InputSearchTB.Text } -Properties * | Select-Object `
    sAMAccountName,
    displayName,
    mail,
    StreetAddress,
    postOfficeBox,
    physicalDeliveryOfficeName,
    l,
    st,
    PostalCode,
    c,
    pager,
    custom-attr,
    mobile,
    title,
    department,
    company, @{Name='Manager';Expression={(Get-ADUser $_.Manager).sAMAccountName}},sAMAccountName -ErrorAction SilentlyContinue

    $global:sAMAccountName              = $user.sAMAccountName
    $global:displayName                 = $user.displayName
    $global:mail                        = $user.mail
    $global:StreetAddress               = $user.StreetAddress
    $global:postofficebox               = $user.postOfficeBox
    $global:physicalDeliveryOfficeName  = $user.physicalDeliveryOfficeName
    $global:l                           = $user.l
    $global:st                          = $user.st
    $global:PostalCode                  = $user.PostalCode
    $global:c                           = $user.c
    $global:pager                       = $user.pager
    $global:custom_attribute            = $user.'custom-attr'
    $global:mobile                      = $user.mobile
    $global:title                       = $user.title
    $global:department                  = $user.department 
    $global:company                     = $user.company
    $global:Manager                     = $user.Manager

    If($user){
        $InputSearchTB.Text     = $sAMAccountName
        $FullNameLB.Text        = $displayName
        $mailLB.Text            = $mail
        $StreetAddressTB.Text   = $streetAddress
        $postOfficeBoxTB.Text   = $postOfficeBox
        $OfficeTB.Text          = $physicalDeliveryOfficeName
        $CityTB.Text            = $l
        $StateTB.Text           = $st
        $PostalCodeTB.Text      = $PostalCode
        $CountryTB.Text         = $c
        $PagerTB.Text           = $pager
        $CustomAttrTB.Text      = $custom_attribute
        $MobileTB.Text          = $mobile
        $JobTitleTB.Text        = $title
        $DepartmentTB.Text      = $department
        $CompanyTB.Text         = $company
        $ManagerTB.Text         = $Manager

        $SaveButton.Enabled = $true
        $returnStatus.Text = ("Successfuly loaded user: " + $FullNameLB.Text + " (login: " + $sAMAccountName + ")")
    } else {
        message -Level "info" -Text ("User does not exist in Active Directory Domain. Try again...")
        clearall
        $returnStatus.Text = ("Ready.")
    }
    return $user
}

function get_user_data_textbox {
	if ($_.KeyCode -eq "Enter") {
		if (![string]::IsNullOrEmpty($InputSearchTB.Text)){
			user_data
		} else {
			$SaveButton.Enabled = $False
        }
    }
}

function get_user_data_button {
    if (![string]::IsNullOrEmpty($InputSearchTB.Text)){
        user_data
    } else {
        $SaveButton.Enabled = $False
    }
}

function SetAttrUser($UserLogin,$Type,$AttrName,$AttrValue){
	if ($AttrValue) {
		switch ($Type) {
			Set {
				if ($AttrName -like "manager") {
					Set-ADUser -Identity $UserLogin -Manager $attrValue
				}
			}
			Default	{
				Set-ADUser -Identity $UserLogin -Replace @{ $attrName = ($attrValue) }
			}
		}
	} else {
		Set-ADUser -Identity $UserLogin -Clear $attrName
	}
}

function save() {
    $InputSearchTB_Val = $InputSearchTB.Text
    If($InputSearchTB_Val) {
		$SaveButton.Enabled = $false
        $returnStatus.Text = "Saving... Please wait."

		#### General ###
        #--- Office ---#
        if (!($physicalDeliveryOfficeName -like $OfficeTB.Text)) {
			SetAttrUser -UserLogin $InputSearchTB_Val -AttrName "physicalDeliveryOfficeName" -AttrValue $OfficeTB.Text
			message -Level "ok" -Text (($OfficeLL.Text) + " '$physicalDeliveryOfficeName' => '" + $OfficeTB.Text + "'")
        }
        
        #### Address ###
		#--- Street Address ---#
		if (!($StreetAddress -like $StreetAddressTB.Text)) {
			SetAttrUser -UserLogin $InputSearchTB_Val -AttrName "StreetAddress" -attrValue $StreetAddressTB.Text
			message -Level "ok" -Text (($StreetAddressLL.Text) + " '$StreetAddress' => '" + $StreetAddressTB.Text + "'")
        }

        #### Telephones ###
		#--- Pager/Custom-Attribute ---#
		if (!($pager -like $PagerTB.Text)) {
			SetAttrUser -UserLogin $InputSearchTB_Val -AttrName "pager" -attrValue $PagerTB.Text
			SetAttrUser -UserLogin $InputSearchTB_Val -AttrName "custom-attr" -attrValue $PagerTB.Text
			message -Level "ok" -Text (($PagerLL.Text) + " '$pager' => '" + $PagerTB.Text + "'")
        }
        
        #### Organization ###
		#--- Job title ---#
		if (!($title -like $JobTitleTB.Text)) {
			SetAttrUser -UserLogin $InputSearchTB_Val -AttrName "title" -attrValue $JobTitleTB.Text
			message -Level "ok" -Text (($JobTitleLL.Text) + " '$title' => '" + $JobTitleTB.Text + "'")
		}
		#--- Department ---#
		if (!($department -like $DepartmentTB.Text)) {
			SetAttrUser -UserLogin $InputSearchTB_Val -AttrName "department" -attrValue $DepartmentTB.Text
			message -Level "ok" -Text (($DepartmentLL.Text) + " '$department' => '" + $DepartmentTB.Text + "'")
		}
		#--- Company ---#
		if (!($company -like $CompanyTB.Text)) {
			SetAttrUser -UserLogin $InputSearchTB_Val -AttrName "company" -attrValue $CompanyTB.Text
			message -Level "ok" -Text (($CompanyLL.Text) + " '$company' => '" + $CompanyTB.Text + "'")
		}
		#--- Manager ---#
		if (!($Manager -like $ManagerTB.Text)) {
            if ($ManagerTB.Text) {
                try {
                    Get-ADUser -Identity $ManagerTB.Text
                } catch {
                    $NotExist = 1
                }
                if ($NotExist -ne 1) {
                    SetAttrUser -UserLogin $InputSearchTB_Val -Type Set -AttrName "manager" -attrValue $ManagerTB.Text
                    message -Level "ok" -Text (($ManagerLL.Text) + " '$Manager' => '" + $ManagerTB.Text + "'")
                } else {
                    message -Level "error" -Text (($ManagerLL.Text) + " '" + ($ManagerTB.Text) + "' does not exist.")
                }
            } else {
                SetAttrUser -UserLogin $InputSearchTB_Val -AttrName "manager" -attrValue $ManagerTB.Text
                message -Level "ok" -Text (($ManagerLL.Text) + " '$Manager' => deleted '" + $ManagerTB.Text + "'")
            }
        }
        $SaveButton.Enabled = $true
        get_user_data_button
    } else {
        $returnStatus.Text = "No user found."
    }
}

function clearall{
    $InputSearchTB.Text     = $global:string
    $FullNameLB.Text        = $global:string
    $MailLB.Text            = $global:string
    $StreetAddressTB.Text   = $global:string
    $postOfficeBoxTB.Text   = $global:string
    $OfficeTB.Text          = $global:string
    $CityTB.Text            = $global:string
    $StateTB.Text           = $global:string
    $PostalCodeTB.Text      = $global:string
    $CountryTB.Text         = $global:string
    $PagerTB.Text           = $global:string
    $CustomAttrTB.Text      = $global:string
    $MobileTB.Text          = $global:string
    $JobTitleTB.Text        = $global:string
    $DepartmentTB.Text      = $global:string
    $CompanyTB.Text         = $global:string
    $ManagerTB.Text         = $global:string
}

#------------[PowerShell Form Initialisations]------------
Add-Type -AssemblyName PresentationFramework
[void]([System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms"))
[void]([System.Reflection.Assembly]::LoadWithPartialName("System.Drawing"))
[System.Windows.Forms.Application]::EnableVisualStyles()

#### Form start here ###############################################################
$form                                       = New-Object system.Windows.Forms.form
$form.ClientSize                            = New-Object System.Drawing.Point($width,$height)
$form.Font                                  = New-Object System.Drawing.Font('Segoe UI',10)
$form.StartPosition                         = "CenterScreen"
$form.text                                  = "[ver. 1.0.3] ADUM - Active Directory User Management by Bogumił Kraszewski and Marcin Siwicki"
$form.TopMost                               = $false
$form.KeyPreview                            = $true
$form.AutoSize                              = $false
$iconBase64                                 = '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'
$iconBytes                                  = [Convert]::FromBase64String($iconBase64)
$stream                                     = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage                                  = [System.Drawing.Image]::FromStream($stream, $true)
$form.Icon                                  = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
$form.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$form.Close()}})
$form.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$SearchButton}})

### Input area #####################################################################
$InputSearchLL                              = New-Object system.Windows.Forms.Label
$InputSearchLL.location                     = New-Object System.Drawing.Point($span_width,10)
$InputSearchLL.Size                         = New-Object System.Drawing.Size(105,25)
$InputSearchLL.Font                         = New-Object System.Drawing.Font('Segoe UI',10)
$InputSearchLL.text                         = "Enter email:"
$InputSearchLL.AutoSize                     = $false
$InputSearchLL.TextAlign                    = 'MiddleRight'
$InputSearchLL.BorderStyle                  = 'None'#'FixedSingle'
$form.Controls.Add($InputSearchLL)

$InputSearchLR                              = New-Object system.Windows.Forms.Label
$InputSearchLR.location                     = New-Object System.Drawing.Point(($span_width + 305), 10)
$InputSearchLR.Size                         = New-Object System.Drawing.Size(120,25)
$InputSearchLR.Font                         = New-Object System.Drawing.Font('Segoe UI', 10)
$InputSearchLR.text                         = "@example.domain"
$InputSearchLR.AutoSize                     = $false
$InputSearchLR.TextAlign                    = 'MiddleLeft'
$InputSearchLR.BorderStyle                  = 'None'
$form.Controls.Add($InputSearchLR)

$InputSearchTB                              = New-Object system.Windows.Forms.TextBox
$InputSearchTB.location                     = New-Object System.Drawing.Point(115,10)
$InputSearchTB.Size                         = New-Object System.Drawing.Size(200,20)
$InputSearchTB.Font                         = New-Object System.Drawing.Font('Segoe UI',10)
$InputSearchTB.multiline                    = $false
$InputSearchTB.Name                         = 'InputSearchTB'
$InputSearchTB.AutoCompleteSource           = 'CustomSource'
$InputSearchTB.AutoCompleteMode             = 'SuggestAppend'
$InputSearchTB.AutoCompleteCustomSource.AddRange((Get-ADUser -Filter * | Where-Object { $_.enabled -like $true }).samaccountname)
$InputSearchTB.Add_TextChanged({enable_search_button})
$InputSearchTB.Add_KeyDown({get_user_data_textbox})
$form.Controls.Add($InputSearchTB)

$FullNameLB                                 = New-Object system.Windows.Forms.Label
$FullNameLB.location                        = New-Object System.Drawing.Point(700,5)
$FullNameLB.Size                            = New-Object System.Drawing.Size(290,20)
$FullNameLB.Font                            = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$FullNameLB.TextAlign                       = 'MiddleCenter'
$FullNameLB.ForeColor                       = 'Green'
$FullNameLB.BorderStyle                     = 'None'
$form.Controls.Add($FullNameLB)

$MailLB                                     = New-Object system.Windows.Forms.Label
$MailLB.location                            = New-Object System.Drawing.Point(700,25)
$MailLB.Size                                = New-Object System.Drawing.Size(290,20)
$MailLB.Font                                = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$MailLB.TextAlign                           = 'MiddleCenter'
$MailLB.ForeColor                           = 'Green'
$MailLB.BorderStyle                         = 'None'
$form.Controls.Add($MailLB)

#### Group TexBoxes ###################################################################
$groupbox_global                            = New-Object System.Windows.Forms.GroupBox
$groupbox_global.Location                   = New-Object System.Drawing.Point($span_width,240)
$groupbox_global.size                       = New-Object System.Drawing.Size($window_width,60)
$groupbox_global.Font                       = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$groupbox_global.Anchor                     = 'top,right,left'
$groupbox_global.text                       = "General"

$groupbox_address                           = New-Object System.Windows.Forms.GroupBox
$groupbox_address.Location                  = New-Object System.Drawing.Point($span_width,300)
$groupbox_address.size                      = New-Object System.Drawing.Size($window_width,180)
$groupbox_address.Font                      = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$groupbox_address.Anchor                    = 'top,right,left'
$groupbox_address.text                      = "Address"

$groupbox_telephones                        = New-Object System.Windows.Forms.GroupBox
$groupbox_telephones.Location               = New-Object System.Drawing.Point($span_width,480)
$groupbox_telephones.size                   = New-Object System.Drawing.Size($window_width,60)
$groupbox_telephones.Font                   = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$groupbox_telephones.Anchor                 = 'top,right,left'
$groupbox_telephones.text                   = "Telephones"

$groupbox_organization                      = New-Object System.Windows.Forms.GroupBox
$groupbox_organization.Location             = New-Object System.Drawing.Point($span_width,540)
$groupbox_organization.size                 = New-Object System.Drawing.Size($window_width,200)
$groupbox_organization.Font                 = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
$groupbox_organization.Anchor               = 'top,right,left'
$groupbox_organization.text                 = "Organization"

#### Output/Input textboxes ###########################################################
$OutputSearchRTB                            = New-Object System.Windows.Forms.RichTextBox
$OutputSearchRTB.Location                   = New-Object System.Drawing.Point($span_width,45)
$OutputSearchRTB.Size                       = New-Object System.Drawing.Size($window_width,190)
$OutputSearchRTB.Font                       = New-Object System.Drawing.Font("Courier", "8")
$OutputSearchRTB.Anchor                     = 'top,right,left'
$OutputSearchRTB.MultiLine                  = $true
$OutputSearchRTB.Wordwrap                   = $true
$OutputSearchRTB.ReadOnly                   = $true
$OutputSearchRTB.ScrollBars                 = "Vertical"
$OutputSearchRTB.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#fff")
$form.Controls.Add($OutputSearchRTB)

$OfficeLL                                   = New-Object system.Windows.Forms.Label
$OfficeLL.location                          = New-Object System.Drawing.Point(11,265)
$OfficeLL.Size                              = New-Object System.Drawing.Size(100,25)
$OfficeLL.Font                              = New-Object System.Drawing.Font('Segoe UI',10)
$OfficeLL.text                              = "Office:"
$OfficeLL.TextAlign                         = 'MiddleRight'
$OfficeLL.BorderStyle                       = 'None'
$OfficeLL.AutoSize                          = $false
$form.Controls.Add($OfficeLL)
$OfficeTB                                   = New-Object system.Windows.Forms.TextBox
$OfficeTB.location                          = New-Object System.Drawing.Point(115,265)
$OfficeTB.Size                              = New-Object System.Drawing.Size(200,25)
$OfficeTB.Font                              = New-Object System.Drawing.Font('Segoe UI',10)
$OfficeTB.ReadOnly                          = $false
$form.Controls.Add($OfficeTB)

$StreetAddressLL                            = New-Object system.Windows.Forms.Label
$StreetAddressLL.location                   = New-Object System.Drawing.Point(11,320)
$StreetAddressLL.Size                       = New-Object System.Drawing.Size(100,150)
$StreetAddressLL.Font                       = New-Object System.Drawing.Font('Segoe UI',10)
$StreetAddressLL.text                       = "Street Address:"
$StreetAddressLL.TextAlign                  = 'MiddleRight'
$StreetAddressLL.BorderStyle                = 'None'
$StreetAddressLL.AutoSize                   = $false
$form.Controls.Add($StreetAddressLL)
$StreetAddressTB                            = New-Object system.Windows.Forms.TextBox
$StreetAddressTB.location                   = New-Object System.Drawing.Point(115,320)
$StreetAddressTB.Size                       = New-Object System.Drawing.Size(200,150)
$StreetAddressTB.Font                       = New-Object System.Drawing.Font('Segoe UI',10)
$StreetAddressTB.ScrollBars                 = "Vertical"
$StreetAddressTB.multiline                  = $true
$form.Controls.Add($StreetAddressTB)

$PostOfficeBoxLL                            = New-Object system.Windows.Forms.Label
$PostOfficeBoxLL.location                   = New-Object System.Drawing.Point(330,320)
$PostOfficeBoxLL.Size                       = New-Object System.Drawing.Size(100,25)
$PostOfficeBoxLL.Font                       = New-Object System.Drawing.Font('Segoe UI',10)
$PostOfficeBoxLL.text                       = "P.O. Box:"
$PostOfficeBoxLL.TextAlign                  = 'MiddleRight'
$PostOfficeBoxLL.BorderStyle                = 'None'
$PostOfficeBoxLL.AutoSize                   = $false
$form.Controls.Add($PostOfficeBoxLL)
$postOfficeBoxTB                            = New-Object system.Windows.Forms.TextBox
$postOfficeBoxTB.location                   = New-Object System.Drawing.Point(435,320)
$postOfficeBoxTB.Size                       = New-Object System.Drawing.Size(200,20)
$postOfficeBoxTB.Font                       = New-Object System.Drawing.Font('Segoe UI',10)
$postOfficeBoxTB.ReadOnly                   = $true
$form.Controls.Add($postOfficeBoxTB)

$CityLL                                     = New-Object system.Windows.Forms.Label
$CityLL.location                            = New-Object System.Drawing.Point(330,350)
$CityLL.Size                                = New-Object System.Drawing.Size(100,25)
$CityLL.Font                                = New-Object System.Drawing.Font('Segoe UI',10)
$CityLL.text                                = "City:"
$CityLL.TextAlign                           = 'MiddleRight'
$CityLL.BorderStyle                         = 'None'
$CityLL.AutoSize                            = $false
$form.Controls.Add($CityLL)
$CityTB                                     = New-Object system.Windows.Forms.TextBox
$CityTB.location                            = New-Object System.Drawing.Point(435,350)
$CityTB.Size                                = New-Object System.Drawing.Size(200,20)
$CityTB.Font                                = New-Object System.Drawing.Font('Segoe UI',10)
$CityTB.ReadOnly                            = $true
$form.Controls.Add($CityTB)

$StateLL                                    = New-Object system.Windows.Forms.Label
$StateLL.location                           = New-Object System.Drawing.Point(330,380)
$StateLL.Size                               = New-Object System.Drawing.Size(100,25)
$StateLL.Font                               = New-Object System.Drawing.Font('Segoe UI',10)
$StateLL.text                               = "State:"
$StateLL.TextAlign                          = 'MiddleRight'
$StateLL.BorderStyle                        = 'None'
$StateLL.AutoSize                           = $false
$form.Controls.Add($StateLL)
$StateTB                                    = New-Object system.Windows.Forms.TextBox
$StateTB.location                           = New-Object System.Drawing.Point(435,380)
$StateTB.Size                               = New-Object System.Drawing.Size(200,20)
$StateTB.Font                               = New-Object System.Drawing.Font('Segoe UI',10)
$StateTB.ReadOnly                           = $true
$form.Controls.Add($StateTB)

$PostalCodeLL                               = New-Object system.Windows.Forms.Label
$PostalCodeLL.location                      = New-Object System.Drawing.Point(330,410)
$PostalCodeLL.Size                          = New-Object System.Drawing.Size(100,25)
$PostalCodeLL.Font                          = New-Object System.Drawing.Font('Segoe UI',10)
$PostalCodeLL.text                          = "Postal Code:"
$PostalCodeLL.TextAlign                     = 'MiddleRight'
$PostalCodeLL.BorderStyle                   = 'None'
$PostalCodeLL.AutoSize                      = $false
$form.Controls.Add($PostalCodeLL)
$PostalCodeTB                               = New-Object system.Windows.Forms.TextBox
$PostalCodeTB.location                      = New-Object System.Drawing.Point(435,410)
$PostalCodeTB.Size                          = New-Object System.Drawing.Size(100,20)
$PostalCodeTB.Font                          = New-Object System.Drawing.Font('Segoe UI',10)
$PostalCodeTB.ReadOnly                      = $true
$form.Controls.Add($PostalCodeTB)

$CountryLL                                  = New-Object system.Windows.Forms.Label
$CountryLL.location                         = New-Object System.Drawing.Point(330,440)
$CountryLL.Size                             = New-Object System.Drawing.Size(100,25)
$CountryLL.Font                             = New-Object System.Drawing.Font('Segoe UI',10)
$CountryLL.text                             = "Country:"
$CountryLL.TextAlign                        = 'MiddleRight'
$CountryLL.BorderStyle                      = 'None'
$CountryLL.AutoSize                         = $false
$form.Controls.Add($CountryLL)
$CountryTB                                  = New-Object system.Windows.Forms.TextBox
$CountryTB.location                         = New-Object System.Drawing.Point(435,440)
$CountryTB.Size                             = New-Object System.Drawing.Size(200,20)
$CountryTB.Font                             = New-Object System.Drawing.Font('Segoe UI',10)
$CountryTB.ReadOnly                         = $true
$form.Controls.Add($CountryTB)

$PagerLL                                    = New-Object system.Windows.Forms.Label
$PagerLL.location                           = New-Object System.Drawing.Point(11,500)
$PagerLL.Size                               = New-Object System.Drawing.Size(100,25)
$PagerLL.Font                               = New-Object System.Drawing.Font('Segoe UI',10)
$PagerLL.text                               = "Pager (Short):"
$PagerLL.TextAlign                          = 'MiddleRight'
$PagerLL.BorderStyle                        = 'None'
$PagerLL.AutoSize                           = $false
$form.Controls.Add($PagerLL)
$PagerTB                                    = New-Object system.Windows.Forms.TextBox
$PagerTB.location                           = New-Object System.Drawing.Point(115,500)
$PagerTB.Size                               = New-Object System.Drawing.Size(200,20)
$PagerTB.Font                               = New-Object System.Drawing.Font('Segoe UI',10)
$PagerTB.multiline                          = $false
$form.Controls.Add($PagerTB)

$MobileLL                                   = New-Object system.Windows.Forms.Label
$MobileLL.location                          = New-Object System.Drawing.Point(330,500)
$MobileLL.Size                              = New-Object System.Drawing.Size(100,25)
$MobileLL.Font                              = New-Object System.Drawing.Font('Segoe UI',10)
$MobileLL.text                              = "Mobile:"
$MobileLL.TextAlign                         = 'MiddleRight'
$MobileLL.BorderStyle                       = 'None'
$MobileLL.AutoSize                          = $false
$form.Controls.Add($MobileLL)
$MobileTB                                   = New-Object system.Windows.Forms.TextBox
$MobileTB.location                          = New-Object System.Drawing.Point(435,500)
$MobileTB.Size                              = New-Object System.Drawing.Size(200,20)
$MobileTB.Font                              = New-Object System.Drawing.Font('Segoe UI',10)
$MobileTB.multiline                         = $false
$MobileTB.ReadOnly                          = $true
$form.Controls.Add($MobileTB)

$JobTitleLL                                 = New-Object system.Windows.Forms.Label
$JobTitleLL.location                        = New-Object System.Drawing.Point(11,570)
$JobTitleLL.Size                            = New-Object System.Drawing.Size(100,25)
$JobTitleLL.Font                            = New-Object System.Drawing.Font('Segoe UI',10)
$JobTitleLL.text                            = "Job Title:"
$JobTitleLL.TextAlign                       = 'MiddleRight'
$JobTitleLL.BorderStyle                     = 'None'
$JobTitleLL.AutoSize                        = $false
$form.Controls.Add($JobTitleLL)
$JobTitleTB                                 = New-Object system.Windows.Forms.TextBox
$JobTitleTB.location                        = New-Object System.Drawing.Point(115,570)
$JobTitleTB.Size                            = New-Object System.Drawing.Size(350,20)
$JobTitleTB.Font                            = New-Object System.Drawing.Font('Segoe UI',10)
$JobTitleTB.multiline                       = $false
$form.Controls.Add($JobTitleTB)

$DepartmentLL                               = New-Object system.Windows.Forms.Label
$DepartmentLL.location                      = New-Object System.Drawing.Point(11,600)
$DepartmentLL.Size                          = New-Object System.Drawing.Size(100,25)
$DepartmentLL.Font                          = New-Object System.Drawing.Font('Segoe UI',10)
$DepartmentLL.text                          = "Department:"
$DepartmentLL.TextAlign                     = 'MiddleRight'
$DepartmentLL.BorderStyle                   = 'None'
$DepartmentLL.AutoSize                      = $false
$form.Controls.Add($DepartmentLL)
$DepartmentTB                               = New-Object system.Windows.Forms.TextBox
$DepartmentTB.location                      = New-Object System.Drawing.Point(115,600)
$DepartmentTB.Size                          = New-Object System.Drawing.Size(350,20)
$DepartmentTB.Font                          = New-Object System.Drawing.Font('Segoe UI',10)
$DepartmentTB.multiline                     = $false
$form.Controls.Add($DepartmentTB)

$CompanyLL                                  = New-Object system.Windows.Forms.Label
$CompanyLL.location                         = New-Object System.Drawing.Point(11,630)
$CompanyLL.Size                             = New-Object System.Drawing.Size(100,25)
$CompanyLL.Font                             = New-Object System.Drawing.Font('Segoe UI',10)
$CompanyLL.text                             = "Company:"
$CompanyLL.TextAlign                        = 'MiddleRight'
$CompanyLL.BorderStyle                      = 'None'
$CompanyLL.AutoSize                         = $false
$form.Controls.Add($CompanyLL)
$CompanyTB                                  = New-Object system.Windows.Forms.TextBox
$CompanyTB.location                         = New-Object System.Drawing.Point(115,630)
$CompanyTB.Size                             = New-Object System.Drawing.Size(350,20)
$CompanyTB.Font                             = New-Object System.Drawing.Font('Segoe UI',10)
$CompanyTB.multiline                        = $false
$form.Controls.Add($CompanyTB)

$ManagerLL                                  = New-Object system.Windows.Forms.Label
$ManagerLL.location                         = New-Object System.Drawing.Point(11,660)
$ManagerLL.Size                             = New-Object System.Drawing.Size(100,25)
$ManagerLL.Font                             = New-Object System.Drawing.Font('Segoe UI',10)
$ManagerLL.text                             = "Manager:"
$ManagerLL.TextAlign                        = 'MiddleRight'
$ManagerLL.BorderStyle                      = 'None'
$ManagerLL.AutoSize                         = $false
$form.Controls.Add($ManagerLL)
$ManagerTB                                  = New-Object system.Windows.Forms.TextBox
$ManagerTB.location                         = New-Object System.Drawing.Point(115,660)
$ManagerTB.Size                             = New-Object System.Drawing.Size(200,20)
$ManagerTB.Font                             = New-Object System.Drawing.Font('Segoe UI',10)
$ManagerTB.multiline                        = $false
$ManagerTB.AutoCompleteSource               = 'CustomSource'
$ManagerTB.AutoCompleteMode                 = 'SuggestAppend'
$ManagerTB.AutoCompleteCustomSource.AddRange((Get-ADUser -Filter * | Where-Object { $_.enabled -like $true }).samaccountname)
$form.Controls.Add($ManagerTB)

$CustomAttrLL                               = New-Object system.Windows.Forms.Label
$CustomAttrLL.location                      = New-Object System.Drawing.Point(11,690)
$CustomAttrLL.Size                          = New-Object System.Drawing.Size(100,25)
$CustomAttrLL.Font                          = New-Object System.Drawing.Font('Segoe UI',10)
$CustomAttrLL.text                          = "Short:"
$CustomAttrLL.Visible                       = $false
$CustomAttrLL.TextAlign                     = 'MiddleRight'
$CustomAttrLL.BorderStyle                   = 'None'
$CustomAttrLL.AutoSize                      = $false
$form.Controls.Add($CustomAttrLL)
$CustomAttrTB                               = New-Object system.Windows.Forms.TextBox
$CustomAttrTB.location                      = New-Object System.Drawing.Point(115,690)
$CustomAttrTB.Size                          = New-Object System.Drawing.Size(200,20)
$CustomAttrTB.Font                          = New-Object System.Drawing.Font('Segoe UI',10)
$CustomAttrTB.Visible                       = $false
$CustomAttrTB.ReadOnly                      = $true
$form.Controls.Add($CustomAttrTB)

# ----------[Button Clicks - Find the user when the search button is pressed]------------
$SearchButton                               = New-Object system.Windows.Forms.Button
$SearchButton.location                      = New-Object System.Drawing.Point(435,10)
$SearchButton.Size                          = New-Object System.Drawing.Size(80,25)
$SearchButton.Font                          = New-Object System.Drawing.Font('Segoe UI',10)
$SearchButton.text                          = "Search"
$SearchButton.Enabled                       = $False
$SearchButton.FlatStyle                     = 'System'
$SearchButton.Add_Click({get_user_data_button})
$form.Controls.Add($SearchButton)

#----------[Button Clicks - Set the user attributes when the set button is pressed]------------
$SaveButton                                 = New-Object system.Windows.Forms.Button
$SaveButton.location                        = New-Object System.Drawing.Point(10,760)
$SaveButton.Size                            = New-Object System.Drawing.Size(80,25)
$SaveButton.Font                            = New-Object System.Drawing.Font('Segoe UI',10)
$SaveButton.text                            = "Save"
$SaveButton.Enabled                         = $False
$SaveButton.Anchor                          = 'Left,Top'
$SaveButton.Add_Click({save})
$form.Controls.Add($SaveButton)

#----------[Button Clicks - Close application when the Quit button is pressed]------------
$QuitButton                                 = New-Object system.Windows.Forms.Button
$QuitButton.location                        = New-Object System.Drawing.Point(115,760)
$QuitButton.Size                            = New-Object System.Drawing.Size(80,25)
$QuitButton.Font                            = New-Object System.Drawing.Font('Segoe UI',10)
$QuitButton.text                            = "Quit"
$QuitButton.Anchor                          = 'Left,Top'
$QuitButton.Add_Click({$form.Close()})
$form.Controls.Add($QuitButton)

#----------[Button Clicks - Clear all values/data when the Clearr All button is pressed]------------
$ClearAllButton                             = New-Object system.Windows.Forms.Button
$ClearAllButton.location                    = New-Object System.Drawing.Point(910,760)
$ClearAllButton.Size                        = New-Object System.Drawing.Size(80,25)
$ClearAllButton.Font                        = New-Object System.Drawing.Font('Segoe UI',10)
$ClearAllButton.text                        = "Clear All"
$ClearAllButton.Anchor                      = 'Right,Top'
$ClearAllButton.Add_Click({clearall})
$form.Controls.Add($ClearAllButton)

$tooltipinfo = New-Object 'System.Windows.Forms.ToolTip'

#### Return Status Bar ################################################################
$returnStatus                               = New-Object System.Windows.Forms.StatusBar
$returnStatus.Text                          = "Please enter user login. Example for user Jan Kowalski: jkowalski"
$form.Controls.Add($returnStatus)

# Activates the form and sets the focus on it
$form.Add_Shown({$InputSearchTB.Select()})
$form.Controls.Add($groupbox_global)
$form.Controls.Add($groupbox_address)
$form.Controls.Add($groupbox_telephones)
$form.Controls.Add($groupbox_organization)

# Check if Module is installed
$ADModuleInstall = Import-Module -Name ActiveDirectory
$ADModuleVerify = Get-Module -Name ActiveDirectory
if (!$ADModuleVerify){
    message -Level "system" -Text ("ActiveDirectory module is not imported! Please contact your System Administrator.")
}

$form.ShowDialog()  | Out-Null
