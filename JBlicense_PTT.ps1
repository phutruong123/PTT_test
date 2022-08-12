<# Source file information
Author: DKL - ELCA VN
Version: 1.0
Last update: 11.07.2019
##################### Release Notes ##########################
+ 11.07.2019: First released
#>
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Generated Windows Form objects
$Main_Form = New-Object system.Windows.Forms.Form
$lblProduct = New-Object system.Windows.Forms.Label
$lblVisa = New-Object system.Windows.Forms.Label
$lblTicket = New-Object system.Windows.Forms.Label
$cmbProduct = New-Object System.Windows.Forms.ComboBox
$txtVisa = New-Object System.Windows.Forms.TextBox
$txtTicket = New-Object System.Windows.Forms.TextBox
$btnOK = New-Object system.Windows.Forms.Button
$btnReset = New-Object system.Windows.Forms.Button
$btnCancel = New-Object system.Windows.Forms.Button
$ErrorProvider = New-Object System.Windows.Forms.ErrorProvider
$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
#endregion Generated Windows Form objects

$AccessFile = "\\rslics-srv.elcanet.local\c$\LicenseService\access-config.json"
$LicenseFile = "\\elcanet.local\vn\infras\sw_licences\_audit SW\License_details_report.XLSX"

#region Setup Log function
function Setup-Log
{	
    $logdir = 'C:\ITSVN\Logs\Assigned_JBLicense'
	if (!(Test-Path $logdir)) {
		New-Item -ItemType Directory -Force -Path $logdir
	}
    $now = Get-Date
	$day = $now.Day
	$month = $now.Month
	$year = $now.Year
	
	$hour = $now.Hour
	$minute = $now.Minute
	$seconde = $now.Second
	
	$file_name = "Assigned-JBLicense-$day-$month-$year-$hour-$minute-$seconde.log"
    $startTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss K"
    		
	Set-Variable -Name logfile -Scope Script -Value (New-Item  (Join-path $logdir $file_name) -Type file)
    Add-Content -Path $logfile -Value ("================= [Automation Script - Assign Jetbrains license]")
    Add-Content -Path $logfile -Value ("================= [Task run by: $env:UserName - On machine: $env:ComputerName]")
    Add-Content -Path $logfile -Value ("================= [Starting time: $startTime]")
}
function AddLog ($Etat,$Data)
{
	$LogTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss K"
	Add-Content -Path $logfile -Value ('[' +$LogTime +'] [' + $Etat + ': ' + $Data + ']')
}
#endregion Setup Log function

#region Send information mail to user
function Send-Mail {
    param (
        $visa,
        $productId
    )
    if (!([String]::IsNullOrEmpty($visa)) -and !([String]::IsNullOrEmpty($productId))) {
        $product = ""
        switch ($productId) {
            "ReSharper 2017.1" { $product = "Resharper 2017.1"; break }
            "ReSharper 2017.3" { $product = "Resharper 2017.3"; break }
            "ReSharper 2019.2" { $product = "Resharper 2019.2"; break }
            "ReSharper 2019.3" { $product = "Resharper 2019.3"; break }
            "ReSharper 2020.2" { $product = "Resharper 2020.2"; break }
            "ReSharper 2020.3" { $product = "Resharper 2020.3"; break }
            "ReSharper Ultimate 2017.1" { $product = "Resharper Ultimate 2017.1"; break }
            "ReSharper 9.2" { $product = "Resharper 9.2"; break }
            "RS0" { $product = "Resharper Toolbox For VS2019"; break }
            "IntelliJ IDEA Ultimate 2017.2" { $product = "IntelliJ IDEA Ultimate 2017.2"; break }
            "IntelliJ IDEA Ultimate 2018.1" { $product = "IntelliJ IDEA Ultimate 2018.1"; break }
            "IntelliJ IDEA Ultimate 2018.2" { $product = "IntelliJ IDEA Ultimate 2018.2"; break }
            "IntelliJ IDEA Ultimate 2019.3" { $product = "IntelliJ IDEA Ultimate 2019.3"; break }
            "IntelliJ IDEA Ultimate 2020.2" { $product = "IntelliJ IDEA Ultimate 2020.2"; break }
            "IntelliJ IDEA Ultimate 2020.3" { $product = "IntelliJ IDEA Ultimate 2020.3"; break }
            "IntelliJ IDEA Ultimate Toolbox" { $product = "IntelliJ IDEA Ultimate Toolbox"; break }
            "WS 2017.2" { $product = "Webstorm 2017.2"; break }
            "WS 2018.1" { $product = "Webstorm 2018.1"; break }
            "PS" { $product = "PhpStorm"; break }
        }
        #$smtp = $visa + "@elca.vn"
        $smtp = Get-ADUser $visa -Properties * | Select mail
        $SMTPServer = "192.168.200.50"
        $Mailer = new-object Net.Mail.SMTPclient($SMTPServer)
        Add-PSSnapin Microsoft.Exchange.Management.Powershell.Admin -erroraction silentlyContinue
        $msg = new-object Net.Mail.MailMessage
        #$smtpclient = new-object Net.Mail.SmtpClient($smtpServer)
        $From = "noreply-licssrv@elca.vn" 
        $To = $smtp.mail
        #$Cc = "it@elca.vn"
        $Subject = "License activation for $product / " + $visa.ToUpper()
        $addCont = ""
        $ticket = $txtTicket.Text.Trim()
        if (!([String]::IsNullOrEmpty($ticket))){
            $addCont = "<br><br>
            Visa: <b>" + $visa.ToUpper() +"</b><br>					
			Product: <b>"+ $product.ToUpper() +"</b><br>
            Ticket Number: <a href='https://projectportal.elca.ch/jira/servicedesk/customer/portal/7/ITSDESK-$ticket'>ITSDESK-$ticket</a><br><br>"
        }
        $Body = "
        <html>
            <head>
                <title></title>
            </head>
            <body style='font-family:arial;font-size: 13px;line-height: 1.2;color: #000000; '>
                <div>
                    Dear User,<br><br>
                    As per your request via our Service Desk system, you have been granted access permission to our <b>JetBrains License Server</b> to activate your <b>$product</b>.
                    $addCont
                    <h4 style='text-decoration: underline; color: #F56217;'>How to activate</h4>
                    Please open license information of your <b>$product</b> then choose <b>License Server</b> activation option and input server address: <b>http://rslics-srv.elcanet.local:8080</b>
                    <div>
                        <h4 style='text-decoration: underline; color: #F56217;'>Support</h4>        
                        If you have any question or need a further assistance, please send your request to our Jira Service Desk: <a href='https://sdo.svc.elca.ch/'> https://sdo.svc.elca.ch/</a>
                    </div>
                </div>
                <br>
                <div>
                    Thanks and best regards,<br>
                    Your ITS Team
                </div>                
            </body>
        </html>"
                        
        $Msg = new-object Net.Mail.MailMessage($From,$To,$Subject,$Body)
        #$Msg.CC.Add($cc)
        $Msg.IsBodyHTML = $True
        $Mailer.send($Msg)
    }    
}
#endregion

#region Update license report file function
function UpdateLicenseFile {
    param (        
        $visa,
        $product
    )    
    $ticket = $txtTicket.Text.Trim()
    $objExcel = New-Object -ComObject Excel.Application
    $objExcel.Visible = $false
    $objExcel.DisplayAlerts = $false
    $WorkBook = $objExcel.Workbooks.Open($LicenseFile)
    $SheetName = "Users"
    $Worksheet = $WorkBook.sheets.item($SheetName)
    $Range = $Worksheet.Range("B1").EntireColumn
    $Target = $Range.Find($visa.ToUpper())
    if ($Target) {
        $col = 0
        switch ($product) {
        "ReSharper 9.2" { $col = 36; break }
        "ReSharper 2017.1" { $col = 37; break }
        "ReSharper 2017.3" { $col = 38; break }
        "ReSharper 2019.2" { $col = 39; break }
        "ReSharper 2019.3" { $col = 40; break }
        "ReSharper 2020.2" { $col = 41; break }
        "ReSharper 2020.3" { $col = 42; break }
        "ReSharper Ultimate 2017.1" { $col = 43; break }
        "RS0" { $col = 44; break }
        "IntelliJ IDEA Ultimate 2017.2" { $col = 45; break }
        "IntelliJ IDEA Ultimate 2018.1" { $col = 46; break }
        "IntelliJ IDEA Ultimate 2018.2" { $col = 47; break }
        "IntelliJ IDEA Ultimate 2019.3" { $col = 48; break }
        "IntelliJ IDEA Ultimate 2020.2" { $col = 49; break }
        "IntelliJ IDEA Ultimate 2020.3" { $col = 50; break }
        "IntelliJ IDEA Ultimate Toolbox" { $col = 51; break }
        "WS 2017.2" { $col = 52; break }
        "WS 2018.1" { $col = 53; break }
        "PS" { $col = 54; break }
        }
        $row = $Target.Row
        $Worksheet.Cells.Item($row,$col) = "x"
        $tickets = $Worksheet.Cells.Item($row,56).Text
        if (!([String]::IsNullOrEmpty($ticket))) {            
            if ($tickets.Length -gt 0) {
                $Worksheet.Cells.Item($row,56) = $tickets + ",#" + $ticket
            }
            else {
                $Worksheet.Cells.Item($row,56) = "#"+ $ticket
            }
        }
        
        # if (!([String]::IsNullOrEmpty($ticket))){
        #     $link = "https://projectportal.elca.ch/jira/servicedesk/customer/portal/7/ITSDESK-$ticket"
        #     $Worksheet.Hyperlinks.Add($Worksheet.Cells.Item($row,43),$link,"","Link to SDO ticket","ITSDESK-$ticket")
        # }
        $WorkBook.Save()
    } 
    $objExcel.Quit()
}
#endregion

#region Add user to access and update license file
function UpdateLicenseAccessFile {
    param (
        [ValidateNotNull()]
		[Parameter(Mandatory = $true)]
        $visa,
        [ValidateNotNull()]
		[Parameter(Mandatory = $true)]
        $productId
    )
    if ((Test-Path $AccessFile)) {
        $JBLicenseObject = Get-Content -Path $AccessFile -Raw | ConvertFrom-Json
        if (!([String]::IsNullOrEmpty($visa)) -and !([String]::IsNullOrEmpty($productId))) {            
            $JBLicenseObject.whitelist | % {if($_.product -eq $productId){                          
                $result = $_.userName | Select-String -Pattern $visa
                    if ($result) {
                        #[System.Windows.Forms.MessageBox]::Show("This user has been assigned to allowed list of this product. Please check!!!","License assigned already","OK","Error")
                        AddLog ERROR "This user has been assigned to allowed list of this product"
                        throw ("This user has been assigned to allowed list of this product. Please check !!!")                
                    }
                    else {
                        #Add user to access-config.json file on License server
                        $replaceString = "|" + $visa.ToUpper() + ")"
                        $_.userName = $_.userName.Replace(")",$replaceString)
                        $JBLicenseObject | ConvertTo-Json -Depth 32 | Set-Content $AccessFile
                        Write-Host "Add user to Access file successful!!!" -BackgroundColor Red -foregroundcolor Green
                        AddLog SUCCEED "Add user $visa to Access file on License Server successful!!!"

                        #Update user to license report file
                        UpdateLicenseFile -visa $visa -product $productId
                        Write-Host "Update user to License file successful!!!" -BackgroundColor Red -foregroundcolor Green
                        AddLog SUCCEED "Update user $visa to License_details_report.XLSX file successful!!!"

                        #Send mail information to user
                        Send-Mail -visa $visa -productId $productId
                        Write-Host "Update Access and License file successful!!!" -BackgroundColor Red -foregroundcolor Green                        
                        AddLog SUCCEED "Update Access and License file for user $visa successful!!!"
                    } 
                }
            }                        
        }
    }              
    else {
        [System.Windows.Forms.MessageBox]::Show("File doesn't esxited or you dont have permission to access","File access error","OK","Error")
        AddLog ERROR "File doesn't esxited or you dont have permission to access"
    }
}
#endregion

#region Remove all event handlers from the controls
$Form_StateCorrection_Load =
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$Main_Form.WindowState = $InitialFormWindowState
	}

$Form_Cleanup_FormClosed =
{		
	try
	{
		$btnOK.remove_Click($btnOK_Click)
		$btnReset.remove_Click($btnReset_Click)
		$btnCancel.remove_Click($btnCancel_Click)
		$Main_Form.remove_Load($MainForm_Load)
		$Main_Form.remove_Load($Form_StateCorrection_Load)
        $Main_Form.remove_FormClosed($Form_Cleanup_FormClosed)
	}
	catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
}
#endregion Remove all event handlers from the controls

#region Load product datatable
function Load-Droplist {
    param (
        [ValidateNotNull()]
		[Parameter(Mandatory = $true)]
		[System.Windows.Forms.ComboBox]$ComboBox
    )
    $dataTable = New-Object System.Data.DataTable
    $dataTable.Columns.Add("ProductID")
    $dataTable.Columns.Add("ProductName")

    $dataTable.Rows.Add("ReSharper 9.2","ReSharper 9.2")
    $dataTable.Rows.Add("ReSharper 2017.1","ReSharper 2017.1")
    $dataTable.Rows.Add("ReSharper 2017.3","ReSharper 2017.3")
    $dataTable.Rows.Add("ReSharper 2019.2","ReSharper 2019.2")
    $dataTable.Rows.Add("ReSharper 2019.3","ReSharper 2019.3")
    $dataTable.Rows.Add("ReSharper 2020.2","ReSharper 2020.2")
    $dataTable.Rows.Add("ReSharper 2020.3","ReSharper 2020.3")
    $dataTable.Rows.Add("ReSharper Ultimate 2017.1","ReSharper Ultimate 2017.1")
    $dataTable.Rows.Add("RS0","Resharper Toolbox For VS2019")  
    $dataTable.Rows.Add("IntelliJ IDEA Ultimate 2017.2","IntelliJ IDEA Ultimate 2017.2")
    $dataTable.Rows.Add("IntelliJ IDEA Ultimate 2018.1","IntelliJ IDEA Ultimate 2018.1")
    $dataTable.Rows.Add("IntelliJ IDEA Ultimate 2018.2","IntelliJ IDEA Ultimate 2018.2")
    $dataTable.Rows.Add("IntelliJ IDEA Ultimate 2019.3","IntelliJ IDEA Ultimate 2019.3")
    $dataTable.Rows.Add("IntelliJ IDEA Ultimate 2020.2","IntelliJ IDEA Ultimate 2020.2")
    $dataTable.Rows.Add("IntelliJ IDEA Ultimate 2020.3","IntelliJ IDEA Ultimate 2020.3")   
    $dataTable.Rows.Add("IntelliJ IDEA Ultimate Toolbox","IntelliJ IDEA Ultimate Toolbox")
    $dataTable.Rows.Add("WS 2017.2","WebStorm 2017.2")
    $dataTable.Rows.Add("WS 2018.1","WebStorm 2018.1")
    $dataTable.Rows.Add("PS","PhpStorm")

    #Clear combobox before we bind it
    $Combobox.Items.Clear()

    #Bind combobox to datatable
    $ComboBox.DataSource = $dataTable
    $ComboBox.ValueMember = "ProductID"
    $ComboBox.DisplayMember = "ProductName"    
}
#endregion

#region Load form
$MainForm_Load = {		
	# We check if the correct PowerShell version 
	if ($PSVersionTable.PSVersion.Major -lt 3)
	{
		[System.Windows.Forms.MessageBox]::Show('PowerShell 3.0 or above is required!', 'Wrong PowerShell Version', 'Ok', 'Error')
		$Main_Form.Close()
    }        

    #Preload Jetbrains products    
    Load-Droplist -ComboBox $cmbProduct
    $cmbProduct.SelectedIndex = -1
}

function Load-Form {
    #Clean up the control events
    $Main_Form.add_Shown({$Main_Form.Activate()})
    $Main_Form.add_Load($MainForm_Load)
    $Main_Form.add_Load($Form_StateCorrection_Load)
	$Main_Form.add_FormClosed($Form_Cleanup_FormClosed)
    $Main_Form.ResumeLayout()
    [void] $Main_Form.ShowDialog()
}
#endregion Load form

#region Generated events
$btnOK_Click = {
    $ErrorProvider.Clear()        
    if($cmbProduct.Text.Trim().Length -eq 0){
        $ErrorProvider.SetError($cmbProduct, "Please choose Product!!!")
    }
    elseif($txtVisa.Text.Trim().Length -lt 3){
        $ErrorProvider.SetError($txtVisa, "Visa of user is 3 characters!!!")
    }
    else{
        try {
            $visa = $txtVisa.Text.Trim()
            $g_visas = ($visa -split ",").Trim()
            foreach ($u_visa in $g_visas) {
                #Check if user is existed or not
                if(@(Get-ADUser -Filter {SamAccountName -eq $u_visa}).Count -eq 0){
                    AddLog ERROR "The visa $u_visa does not exist. Please check if account already created"
                    [System.Windows.Forms.MessageBox]::Show("The visa $u_visa does not exist. Please check if account already created.","User doesn't exist","OK","Error")
                    return
                }
                else {
                    Write-Host $cmbProduct.SelectedItem["ProductID"]
                    UpdateLicenseAccessFile -visa $u_visa -productId $cmbProduct.SelectedItem["ProductID"]                    
                }  
            }
            AddLog Finished "Complete update Access and License file for $visa"
            $result = [System.Windows.Forms.MessageBox]::Show("Complete update Access and License file for $visa. Do you want to close this application?","Update completed",[System.Windows.Forms.MessageBoxButtons]::OKCancel)                            
            switch ($result){
                "OK" {
                    $Main_Form.Close()
                } 
                "Cancel" {
                    #return
                } 
            }                
        }
        catch {
            $Text = "Update license FAIL !
            Exception Message : $_.Exception.Message"
            AddLog Fail $Text
            [System.Windows.Forms.MessageBox]::Show($Text,"Update access and license failed","OK","Error")            
            return
        }        
    }
}

$btnReset_Click = {
    $cmbProduct.SelectedIndex = -1    
    $txtVisa.Text = ""
    $txtTicket.Text = ""
    $ErrorProvider.Clear() 
}

$btnCancel_Click = {
	# Close the form and quit
    $Main_Form.Close()    
}	
#endregion Generated events
function Load-ComboBox
{
	<#
		.SYNOPSIS
			This functions helps you load items into a ComboBox.
	
		.DESCRIPTION
			Use this function to dynamically load items into the ComboBox control.
	
		.PARAMETER  ComboBox
			The ComboBox control you want to add items to.
	
		.PARAMETER  Items
			The object or objects you wish to load into the ComboBox's Items collection.
	
		.PARAMETER  DisplayMember
			Indicates the property to display for the items in this control.
		
		.PARAMETER  Append
			Adds the item(s) to the ComboBox without clearing the Items collection.
		
		.EXAMPLE
			Load-ComboBox $combobox1 "Red", "White", "Blue"
		
		.EXAMPLE
			Load-ComboBox $combobox1 "Red" -Append
			Load-ComboBox $combobox1 "White" -Append
			Load-ComboBox $combobox1 "Blue" -Append
		
		.EXAMPLE
			Load-ComboBox $combobox1 (Get-Process) "ProcessName"
	#>
		Param (
			[ValidateNotNull()]
			[Parameter(Mandatory = $true)]
			[System.Windows.Forms.ComboBox]$ComboBox,
			[ValidateNotNull()]
			[Parameter(Mandatory = $true)]
			$Items,
			[Parameter(Mandatory = $false)]
			[string]$DisplayMember,
			[switch]$Append
		)
		
		if (-not $Append){
			$ComboBox.Items.Clear()
		}
		
		if ($Items -is [Object[]]){
			$ComboBox.Items.AddRange($Items)
		}
		elseif ($Items -is [Array])
		{
			$ComboBox.BeginUpdate()
			foreach ($obj in $Items)
			{
				$ComboBox.Items.Add($obj)
			}
			$ComboBox.EndUpdate()
		}
		else{
			$ComboBox.Items.Add($Items)
		}		
		$ComboBox.DisplayMember = $DisplayMember
}

#region Generated form and its component
# Main Form
$Main_Form.Size = '480,340'
$Main_Form.text = "Jetbrains License Tool v1.0"
#$Main_Form.AutoScaleDimensions = '8, 20'
$Main_Form.AutoScaleMode = 'Font'
$Main_Form.AutoSize = $true
$Main_Form.AutoSizeMode = 'GrowOnly'
$Main_Form.FormBorderStyle = 'FixedSingle'
#$Main_Form.Margin = '15, 15, 15, 15'
$Main_Form.MaximizeBox = $False
$Main_Form.StartPosition = 'CenterScreen'
#$Main_Form.TopMost = $true

#Product components
$lblProduct.text = "Please choose Product:"
$lblProduct.AutoSize = $true
$lblProduct.Size = '25,10'
$lblProduct.location = '15,25'

$cmbProduct.Size = '380,20'
$cmbProduct.location = '45,50'
$cmbProduct.Font = "Microsoft Sans Serif,15"
$cmbProduct.Sorted = $true
$cmbProduct.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;

#Visa of user components
$lblVisa.text = "Please enter visa of user
(You can enter a list of users by using comma <,> to seperate each visa)"
$lblVisa.AutoSize = $true
$lblVisa.Size = '25,20'
$lblVisa.location = '15,100'

$txtVisa.Size = '380,20'
$txtVisa.location = '45,140'
$txtVisa.Font = "Microsoft Sans Serif,15"

#Ticket number components
$lblTicket.text = "Please enter ticket number (If any):"
$lblTicket.AutoSize = $true
$lblTicket.Size = '25,20'
$lblTicket.location = '15,175'

$txtTicket.Size = '380,20'
$txtTicket.location = '45,200'
$txtTicket.Font = "Microsoft Sans Serif,15"

#Button components
$btnOK.text = "Update"
$btnOK.Size = '90,30'
$btnOK.location = '40,250'
$btnOK.Margin = '4, 4, 4, 4'
$btnOK.add_Click($btnOK_Click)

$btnReset.text = "Reset"
$btnReset.Size = '90,30'
$btnReset.location = '160,250'
$btnReset.Margin = '4, 4, 4, 4'
$btnReset.add_Click($btnReset_Click)

$btnCancel.text = "Cancel"
$btnCancel.Size = '90,30'
$btnCancel.location = '280,250'
$btnCancel.Margin = '4, 4, 4, 4'
$btnCancel.add_Click($btnCancel_Click)

$Main_Form.controls.AddRange(@($lblProduct,$cmbProduct,$lblVisa,$txtVisa,$lblTicket,$txtTicket,$btnOK,$btnReset,$btnCancel))
#endregion Generated Form and its component

Setup-Log
Load-Form