<# Version:    2.2
License:    GNU GENERAL PUBLIC LICENSE Version 3
Autor:      Béla Richartz
Created:    2019/09
Purpose:    Adding and removing of network printer.
            Printer are importet from Printserver over Get-Printer function. (write PrintServer to config.xml)
            Connected printers (local and network Printer) can be set as default.
Github:     https://github.com/sevi-kun/PrinterTool
Edited from:    Béla Richartz
Last edited:    2020/01
#>
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()


#           Config File
################################################
[xml]$Global:ConfigFile = Get-Content ".\config.xml"
################################################

# Read config.xml
$Global:PrinterTool_Version = $Global:ConfigFile.Config.General.PrinterTool_Version
$Global:Printservers = $Global:ConfigFile.Config.General.Printserver.Server
$Global:Blacklist = $Global:ConfigFile.Config.Blacklist.Printer


############# Functions
################################################################


###
function func_GUI_Variables {# creating GUI
    ############# Form
    $Global:Form = New-Object system.Windows.Forms.Form
    $Global:Form.ClientSize = '600,300'
    $Global:Form.text = "PrinterTool" + ' ' + $Global:PrinterTool_Version
    $Global:Form.TopMost = $false
    $Global:Form.FormBorderStyle = 'FixedDialog'
    $Global:Form.Icon = $icon
        
    ############# Group Select/Search Printer
    #######################################################################################
    $Global:Group_RemotePrinter = New-Object system.Windows.Forms.Groupbox
    $Global:Group_RemotePrinter.height = 275
    $Global:Group_RemotePrinter.width = 200
    $Global:Group_RemotePrinter.text = "Remote printers:"
    $Global:Group_RemotePrinter.location = New-Object System.Drawing.Point(10, 15)
        
    $Global:GrRP_Lbl_Search = New-Object system.Windows.Forms.Label
    $Global:GrRP_Lbl_Search.text = "Search:"
    $Global:GrRP_Lbl_Search.AutoSize = $true
    $Global:GrRP_Lbl_Search.height = 10
    $Global:GrRP_Lbl_Search.width = 25
    $Global:GrRP_Lbl_Search.location = New-Object System.Drawing.Point(13, 23)
    $Global:GrRP_Lbl_Search.Font = 'Microsoft Sans Serif1,8'
        
    $Global:GrRP_TxtBox_Search = New-Object system.Windows.Forms.TextBox
    $Global:GrRP_TxtBox_Search.multiline = $false
    $Global:GrRP_TxtBox_Search.height = 20
    $Global:GrRP_TxtBox_Search.width = 115
    $Global:GrRP_TxtBox_Search.location = New-Object System.Drawing.Point(75, 20)
    $Global:GrRP_TxtBox_Search.Font = 'Microsoft Sans Serif,10'
        
    $Global:GrRP_List_Prt = New-Object system.Windows.Forms.ListBox
    $Global:GrRP_List_Prt.text = "GrRPList_Prt"
    $Global:GrRP_List_Prt.height = 220
    $Global:GrRP_List_Prt.width = 180
    $Global:GrRP_List_Prt.location = New-Object System.Drawing.Point(10, 50)
        
    ############# Group Current/Connected Printer
    $Global:Group_CurrentPrinter = New-Object system.Windows.Forms.Groupbox
    $Global:Group_CurrentPrinter.Height = 280
    $Global:Group_CurrentPrinter.Width = 200
    $Global:Group_CurrentPrinter.Text = "Connected printers:"
    $Global:Group_CurrentPrinter.Location = New-Object System.Drawing.Point(390, 15)
        
    $Global:GrCP_List_Prt = New-Object system.Windows.Forms.ListBox
    $Global:GrCP_List_Prt.text = "GrCP_List_Prt"
    $Global:GrCP_List_Prt.height = 230
    $Global:GrCP_List_Prt.width = 180
    $Global:GrCP_List_Prt.Location = New-Object System.Drawing.Point(10, 20)
        
    $Global:GrCP_Btn_SetStd = New-Object system.Windows.Forms.Button
    $Global:GrCP_Btn_SetStd.text = 'Set as default'
    $Global:GrCP_Btn_SetStd.Height = 20
    $Global:GrCP_Btn_SetStd.Width = 180
    $Global:GrCP_Btn_SetStd.Location = New-Object System.Drawing.Point(10, 255)
        
    ############# Buttons to add and remove Printers
    #######################################################################################
    $Global:Btn_Prt_Add = New-Object System.Windows.Forms.Button
    $Global:Btn_Prt_Add.Text = 'Add >>'
    $Global:Btn_Prt_Add.Height = 20
    $Global:Btn_Prt_Add.Width = 160
    $Global:Btn_Prt_Add.Location = New-Object System.Drawing.Point(220, 130)
        
    $Global:Btn_Prt_rm = New-Object System.Windows.Forms.Button
    $Global:Btn_Prt_rm.Text = '<< Remove'
    $Global:Btn_Prt_rm.Height = 20
    $Global:Btn_Prt_rm.Width = 160
    $Global:Btn_Prt_rm.Location = New-Object System.Drawing.Point(220, 170)
        
    $Global:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $Global:ProgressBar.Name = 'ProgressBar'
    $Global:ProgressBar.Style = "Continuous"
    $Global:ProgressBar.Maximum = 100
    $Global:ProgressBar.Minimum = 0
    $Global:ProgressBar.Step = 2
    $Global:ProgressBar.Height = 15
    $Global:ProgressBar.Width = 160
    $Global:ProgressBar.Location = New-Object System.Drawing.Size (220, 220)
        
    $Global:ProgressBar_Label = New-Object System.Windows.Forms.Label
    $Global:ProgressBar_Label.Name = 'Label ProgressBar'
    $Global:ProgressBar_Label.Font = 'Microsoft Sans Serif1,8'
    $Global:ProgressBar_Label.ForeColor = "#ff001e"
    $Global:ProgressBar_Label.Height = 20
    $Global:ProgressBar_Label.Width = 160
    $Global:ProgressBar_Label.Location = New-Object System.Drawing.Size (220, 240)
}
function func_GUI_Build {# build GUI and add connections
    ############# Build Form
    $Global:Form.controls.AddRange(@($Global:Group_RemotePrinter))
    $Global:Form.controls.AddRange(@($Global:Group_CurrentPrinter))
    $Global:Group_RemotePrinter.controls.AddRange(@($Global:GrRP_List_Prt, $Global:GrRP_TxtBox_Search, $Global:GrRP_Lbl_Search))
    $Global:Group_CurrentPrinter.Controls.AddRange(@($Global:GrCP_List_Prt, $Global:GrCP_Btn_SetStd))
    $Global:Form.Controls.AddRange(@($Global:Btn_Prt_rm, $Global:Btn_Prt_Add, $Global:ProgressBar, $Global:ProgressBar_Label))
        
    ############# Connections GUI <> Funktionen
    $Global:GrRP_TxtBox_Search.Add_TextChanged( { func_search })
    $Global:GrCP_Btn_SetStd.Add_Click( { func_SetStd })
    $Global:GrCP_List_Prt.Add_SelectedValueChanged( { func_IsStd })
    $Global:GrRP_List_Prt.Add_SelectedValueChanged( { func_GrRPSelected })
    $Global:Btn_Prt_Add.Add_Click( { func_add_Prt })
    $Global:Btn_Prt_rm.Add_Click( { func_rm_Prt })
}
function func_Startfunction {# Loading variables and some start functions

    ############# Get variables
    $Global:Obj_Connected_Prt = Get-Printer | Where-Object Name -NotLike *OneNote* | Where-Object Name -NotLike *PDF*
    [System.Collections.ArrayList]$Global:Connected_Prt_Name = $Global:Obj_Connected_Prt.Name
    $Global:Obj_StdPrinter = Get-WmiObject -Query " SELECT * FROM Win32_Printer WHERE Default=$true"
    foreach ($srv in $Global:Printservers) {
        $Global:Obj_PrtSrv_Prt += Get-Printer -ComputerName $srv | Where-Object -Property DeviceType -EQ "Print"
    }
    [System.Collections.ArrayList]$Global:PrtSrv_Prt_Name = $Global:Obj_PrtSrv_Prt.Name

    ###### Start of: Loading content
    ################################################################

    # clean lists
    $Global:ProgressBar.Hide()
    $Global:GrRP_List_Prt.Items.Clear()
    $Global:GrCP_List_Prt.Items.Clear()

    try {
        # adjust printer lists (removing Connected printers from PrtSrv-List / mark connected printers)
        foreach ($obj in $Global:Connected_Prt_Name) {
            # removing Connected printers from PrtSrv-List
            $Local:printer = $obj
            foreach ($srv in $Global:Printservers) {
                $Local:printer = $Local:printer.Replace("\\" + $srv + "\", "")
                $Local:printer = $Local:printer.Replace("\\" + $srv.ToLower() + "\", "")
                $Local:printer = $Local:printer.Replace("\\" + $srv.ToUpper() + "\", "")
            }
            $Local:printer = $Local:printer.ToString()
            $Global:PrtSrv_Prt_Name.Remove($Local:printer)
        }
        foreach ($obj in $Global:Blacklist) {
            $Global:PrtSrv_Prt_Name.Remove($obj)
        }


    }
    catch {
        if (!$Global:Connected_Prt) {
            [System.Windows.Forms.MessageBox]::Show("func_Startfunction: Printer list could not be loaded or adjusted.`n`n" + $_, '$Connected_Prt')
        }
        elseif (!$Global:Obj_StdPrinter) {
            [System.Windows.Forms.MessageBox]::Show("func_Startfunction: Printer list could not be loaded or adjusted.`n`n" + $_, '$StdPrinter')
        }
        elseif (!$Global:Obj_PrtSrv_Prt) {
            [System.Windows.Forms.MessageBox]::Show("func_Startfunction: Printer list could not be loaded or adjusted.`n`n" + $_, '$PrtSrv_Prt')
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("func_Startfunction: Printer list could not be loaded or adjusted.`n`n" + $_, 'Information')
        }
    }

    try {
        # GrSP/GrCP lists fill
        foreach ($obj in $Global:Connected_Prt_Name) {
            if ($obj -like "*\\*") {
                $Local:printrstr = $Global:Obj_StdPrinter.ShareName
                $Local:printer = $obj
                foreach ($srv in $Global:Printservers) {
                    $Local:printer = $Local:printer.Replace("\\" + $srv + "\", "")
                    $Local:printer = $Local:printer.Replace("\\" + $srv.ToLower() + "\", "")
                    $Local:printer = $Local:printer.Replace("\\" + $srv.ToUpper() + "\", "")
                }
                if ($Local:printer -eq $Local:printrstr) { $Local:printer = $Local:printer + ' (Default)' }
                $Global:GrCP_List_Prt.Items.Add($Local:printer)
            }
            else {
                $Local:printrstr = $Global:Obj_StdPrinter.Name
                $Local:printer = $obj + ' (Local)'
                if ($Local:printer -eq ($Local:printrstr + ' (Local)')) { $Local:printer = $Local:printer + ' (Default)' }
                $Global:GrCP_List_Prt.Items.Add($Local:printer)
            }
        }

        foreach ($obj in $Global:PrtSrv_Prt_Name) {
            $Local:printer = $obj.Replace("PRT-", "") 
            $Global:GrRP_List_Prt.Items.Add($Local:printer)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("func_Startfunction: Printer could not be loaded.`n`n" + $_, "Information")
    }



    ###### End of: Loading content
    ################################################################
}
###


###
function func_add_Prt {# Connecting selected printer
    $Global:Btn_Prt_Add.Enabled = $false
    $Global:Btn_Prt_rm.Enabled = $false
    $Global:GrCP_Btn_SetStd.Enabled = $false
    $Global:ProgressBar_Label.Text = "Adding printer.."
    func_ProgressBar_T1
    $Local:Printer = $Global:GrRP_List_Prt.SelectedItem     # Get selected Item
    $Local:Printer_ComputerName = $Global:Obj_PrtSrv_Prt | Where-Object Name -EQ $Local:Printer | Select-Object -Property ComputerName
    $Local:Printer_Name = '\\' + $Local:Printer_ComputerName + '\' + $Local:Printer.Replace("PRT-", "")
    $Local:Printer_Name = $Local:Printer_Name.Replace('@{ComputerName=', '')
    $Local:Printer_Name = $Local:Printer_Name.Replace('}', '')
    try {
        (New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($Local:Printer_Name)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("func_add_Prt: There appeard an error while adding the printer. `n`n" + $_, "Information")
    }
    func_ProgressBar_T2
    func_UpdateLists
    func_ProgressBar_T3
    $Global:ProgressBar_Label.Text = ""
}
function func_rm_Prt {# Removing selected printer
    $Global:ProgressBar_Label.Text = "Removing printer.."
    $Global:Btn_Prt_Add.Enabled = $false
    $Global:Btn_Prt_rm.Enabled = $false
    $Global:GrCP_Btn_SetStd.Enabled = $false
    func_ProgressBar_T1
    $Local:Printer = $Global:GrCP_List_Prt.SelectedItem     # Get selected Item
    $Local:Printer = $Local:Printer.Replace(" (Default)", "")
    $Local:Printer = $Local:Printer.Replace(" (Local)", "")
    $Local:Printer_ComputerName = $Global:Obj_PrtSrv_Prt | Where-Object Name -EQ $Local:Printer | Select-Object -Property ComputerName
    $Local:Printer_Name = '\\' + $Local:Printer_ComputerName + '\' + $Local:Printer.Replace("PRT-", "")
    $Local:Printer_Name = $Local:Printer_Name.Replace('@{ComputerName=', '')
    $Local:Printer_Name = $Local:Printer_Name.Replace('}', '')
    $Local:Printer_Name = $Local:Printer_Name.Replace(" (Default)", "")
    $Local:Printer_Name = $Local:Printer_Name.Replace(" (Local)", "")
    try {
        (New-Object -ComObject WScript.Network).RemovePrinterConnection($Local:Printer_Name)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("func_rm_Prt: There appeard an error while removing the printer. `n`n" + $_, "Information")
    }
    func_ProgressBar_T2
    func_UpdateLists    # Updating Lists
    func_ProgressBar_T3
    $Global:ProgressBar_Label.Text = ""
}
function func_SetStd {# Setting selected printer as default
    $Global:ProgressBar_Label.Text = "Setting the default printer.."
    func_ProgressBar_T1
    $Global:GrCP_Btn_SetStd.Enabled = $false
    $Global:Btn_Prt_Add.Enabled = $false
    $Global:Btn_Prt_rm.Enabled = $false
    $Local:Printer = $Global:GrCP_List_Prt.SelectedItem
    if ($Local:Printer -like "*(Local)*") {
        $Local:Printer_Name = $Local:Printer.Replace(" (Local)", "")
    }
    else {
        $Local:Printer_ComputerName = $Global:Obj_PrtSrv_Prt | Where-Object Name -EQ $Local:Printer | Select-Object -Property ComputerName
        $Local:Printer_Name = '\\' + $Local:Printer_ComputerName + '\' + $Local:Printer.Replace("PRT-", "")
        $Local:Printer_Name = $Local:Printer_Name.Replace('@{ComputerName=', '')
        $Local:Printer_Name = $Local:Printer_Name.Replace('}', '')
    }

    try {
        (New-Object -ComObject WScript.Network).SetDefaultPrinter($Local:Printer_Name)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("func_SetStd: There appeard an error while setting the selected printer as default. `n`n" + $_, "Information")
    }

    func_ProgressBar_T2
        
    # Updating Lists Button etc..
    func_UpdateLists
    $Global:Obj_StdPrinter = Get-WmiObject -Query " SELECT * FROM Win32_Printer WHERE Default=$true"
    func_IsStd
    func_ProgressBar_T3
    $Global:ProgressBar_Label.Text = ""
}
###

###
function func_search {# searchfunction
    try {
        # Clear Listbox
        $Global:GrRP_List_Prt.Items.Clear()

        # Get Text from TxtBox_Search
        $Local:SearchValue = $Global:GrRP_TxtBox_Search.Text

        foreach ($obj in $Global:PrtSrv_Prt_Name) {
            $Local:printer = $obj.Replace("PRT-", "") 

            If ($Local:printer -like "*$Local:SearchValue*") {
                $Global:GrRP_List_Prt.Items.Add($Local:printer)
            }
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("func_search: There appeard an error while searching the printer.`n`n" + $_, "Information")
    }
}

function func_UpdateLists {# Updating Variables and ListBoxes
    # Clean Variables
    $Global:Obj_Connected_Prt = $null
    [System.Collections.ArrayList]$Global:Connected_Prt_Name = $null
    $Global:Obj_StdPrinter = $null
    $Global:Obj_PrtSrv_Prt = $null
    [System.Collections.ArrayList]$Global:PrtSrv_Prt_Name = $null

    # Get Variables
    $Global:Obj_Connected_Prt = Get-Printer | Where-Object Name -NotLike *OneNote* | Where-Object Name -NotLike *PDF*
    [System.Collections.ArrayList]$Global:Connected_Prt_Name = $Global:Obj_Connected_Prt.Name
    $Global:Obj_StdPrinter = Get-WmiObject -Query " SELECT * FROM Win32_Printer WHERE Default=$true"
    foreach ($srv in $Global:Printservers) {
        $Global:Obj_PrtSrv_Prt += Get-Printer -ComputerName $srv | Where-Object -Property DeviceType -EQ "Print"
    }
    [System.Collections.ArrayList]$Global:PrtSrv_Prt_Name = $Global:Obj_PrtSrv_Prt.Name

    # Clean ListBoxes
    $Global:GrRP_List_Prt.Items.Clear()
    $Global:GrCP_List_Prt.Items.Clear()

    try {
        foreach ($obj in $Global:Connected_Prt_Name) {
            $Local:printer = $obj

            $Local:printer = $obj
            foreach ($srv in $Global:Printservers) {
                $Local:printer = $Local:printer.Replace("\\" + $srv + "\", "")
                $Local:printer = $Local:printer.Replace("\\" + $srv.ToLower() + "\", "")
                $Local:printer = $Local:printer.Replace("\\" + $srv.ToUpper() + "\", "")
            }
            $Local:printer = $Local:printer.ToString()
            $Global:PrtSrv_Prt_Name.Remove($Local:printer)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("func_UpdateLists: There appeard an error while changing the printserver list. `n`n" + $_, "Information")
    }


    try {
        foreach ($obj in $Global:Connected_Prt_Name) {
            if ($obj -like "*\\*") {
                $Local:printrstr = $Global:Obj_StdPrinter.ShareName     # Get Default printer name

                $Local:printer = $obj
                foreach ($srv in $Global:Printservers) {
                    $Local:printer = $Local:printer.Replace("\\" + $srv + "\", "")
                    $Local:printer = $Local:printer.Replace("\\" + $srv.ToLower() + "\", "")
                    $Local:printer = $Local:printer.Replace("\\" + $srv.ToUpper() + "\", "")
                }
                if ($Local:printer -eq $Local:printrstr) { $Local:printer = $Local:printer + ' (Default)' }
                $Global:GrCP_List_Prt.Items.Add($Local:printer)
            }
            else {
                $Local:printrstr = $Global:Obj_StdPrinter.Name
                $Local:printer = $obj + ' (Local)'
                if ($Local:printer -eq ($Local:printrstr + ' (Local)')) { $Local:printer = $Local:printer + ' (Default)' }
                $Global:GrCP_List_Prt.Items.Add($Local:printer)
            }
        }

        foreach ($obj in $Global:PrtSrv_Prt_Name) {
            $Local:printer = $obj.Replace("PRT-", "") 
            $Global:GrRP_List_Prt.Items.Add($Local:printer)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("func_UpdateLists: There appeard an error while loading the printer. `n`n" + $_, "Information")
    }
    func_search
}
###

function func_IsStd {# Disabeling buttons if Printer is Default (Btn_Prt_rm, Btn_Prt_Add, GrCP_Btn_SetStd)
    $Global:Btn_Prt_rm.Enabled = $true
    $Global:Btn_Prt_Add.Enabled = $true
    $Global:GrCP_Btn_SetStd.Enabled = $true
    $Local:StdPrinter = $Global:Obj_StdPrinter.Name
    foreach ($srv in $Global:Printservers) {
        $Local:StdPrinter = $Local:StdPrinter.Replace("\\" + $srv + "\", "")
        $Local:StdPrinter = $Local:StdPrinter.Replace("\\" + $srv.ToLower() + "\", "")
        $Local:StdPrinter = $Local:StdPrinter.Replace("\\" + $srv.ToUpper() + "\", "")
    }
    $Local:StdPrinter = $Local:StdPrinter.ToString()
    

    $Local:SelectedPrinter = $Global:GrCP_List_Prt.SelectedItem
    if ($Local:SelectedPrinter -like "*(Local)*") { $Local:SelectedPrinter = $Local:SelectedPrinter.Replace(" (Local)", "") }
    if ($Local:SelectedPrinter -like "*(Default)*") { $Local:SelectedPrinter = $Local:SelectedPrinter.Replace(' (Default)', '') }
    try {
        if ($Local:StdPrinter -eq $Local:SelectedPrinter) {
            $Global:GrCP_Btn_SetStd.Enabled = $false
            $Global:Btn_Prt_rm.Enabled = $false
            $Global:Btn_Prt_Add.Enabled = $false
        }
        else {
            $Global:GrCP_Btn_SetStd.Enabled = $true
            $Global:Btn_Prt_rm.Enabled = $true
            $Global:Btn_Prt_Add.Enabled = $false
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("func_IsStd: There appeard an error while updating the UI. `n`n" + $_, "Information")
    }

}

function func_GrRPSelected {# Disabeling buttons (Btn_Prt_rm, GrCP_Btn_SetStd), when selecting GrRP_List_Prt
    $Global:Btn_Prt_rm.Enabled = $true
    $Global:Btn_Prt_Add.Enabled = $true
    $Global:GrCP_Btn_SetStd.Enabled = $true
    $Global:Btn_Prt_rm.Enabled = $false
    $Global:GrCP_Btn_SetStd.Enabled = $false
}


###
function func_CheckConfig {# Check if there is an error with the config.xml
if (!$Global:ConfigFile -or !$Global:Printservers) { [System.Windows.Forms.MessageBox]::Show("Configuration file was not found or could not be read.", "Information") }
}

function func_ProgressBar_T1 {# Progress bar 30%
    $Global:ProgressBar.Show()
    while ($Global:ProgressBar.Value -lt ($Global:ProgressBar.Maximum - 70)) {
        $Global:ProgressBar.PerformStep()
        Start-Sleep -Milliseconds 10
    }
}
function func_ProgressBar_T2 {# Progress bar 60%
    while ($Global:ProgressBar.Value -lt ($Global:ProgressBar.Maximum - 40)) {
        $Global:ProgressBar.PerformStep()
        Start-Sleep -Milliseconds 20
    }
}
function func_ProgressBar_T3 {# Progress bar 100%
    while ($Global:ProgressBar.Value -lt $Global:ProgressBar.Maximum) {
        $Global:ProgressBar.PerformStep()
        Start-Sleep -Milliseconds 20
    }
    $Global:ProgressBar.Hide()
    $Global:ProgressBar.Value = 0
}
###



# Starting application
func_CheckConfig
func_GUI_Variables
func_GUI_Build
func_Startfunction

# show GUI
[void]$Global:Form.ShowDialog()


################################################################################################################################
# EOF
################################################################################################################################
