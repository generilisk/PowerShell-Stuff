<#
.Synopsis
    Displays a graphical user interface for managing printers. 
    Allows users to add a printer, remove a printer, or set a default printer.
.DESCRIPTION
    This script provides a Windows Forms-based GUI for managing printers in a networked environment. 
    Users can:
    - Add a printer by specifying its name. 
    - Remove one or more printers that are hosted on a specific server, excluding certain printers.
    - Set a selected printer as the default printer.
    The GUI includes buttons for each action and provides feedback to the user on success or failure.
.EXAMPLE
    # Launch the main GUI with options to add, remove, or set default printer.
    .\PrinterManagementGUI.ps1
    This command opens the main GUI dialog where users can choose to add a printer, remove a printer, or set a printer as default.
.COPYRIGHT
    Copyright (c) 2024 Will Hughes. All rights reserved.
    This script is licensed for use by Shasta Community Health Center employees only.
.LICENSE
    This script may be used and modified for the benefit of Shasta Community Health Center employees only.
    Unauthorized use or distribution outside of Shasta Community Health Center is prohibited.

#>

# Script Name: PrinterManagementGUI.ps1
# Created by Will Hughes
# Date: 2024-08-12
#Patch Notes:

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to create the main dialog box
function Show-MainDialog {
    $form = New-Object Windows.Forms.Form
    $form.Text = 'Printer Management'
    $form.Size = New-Object Drawing.Size(300,250)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog' # Make the form non-resizable
    $form.MaximizeBox = $false # Disable maximize button
    $form.MinimizeBox = $false # Disable minimize button

    $buttonWidth = 200
    $buttonHeight = 30
    $horizontalCenter = ($form.ClientSize.Width - $buttonWidth) / 2 # Centering horizontally

    $addPrinterButton = New-Object Windows.Forms.Button
    $addPrinterButton.Text = 'Add Printer'
    $addPrinterButton.Location = New-Object Drawing.Point($horizontalCenter,30)
    $addPrinterButton.Size = New-Object Drawing.Size($buttonWidth, $buttonHeight)
    $addPrinterButton.Add_Click({ Show-AddPrinterDialog })

    $removePrinterButton = New-Object Windows.Forms.Button
    $removePrinterButton.Text = 'Remove Printer'
    $removePrinterButton.Location = New-Object Drawing.Point($horizontalCenter,70)
    $removePrinterButton.Size = New-Object Drawing.Size($buttonWidth, $buttonHeight)
    $removePrinterButton.Add_Click({ Show-RemovePrinterDialog })

    $setDefaultPrinterButton = New-Object Windows.Forms.Button
    $setDefaultPrinterButton.Text = 'Set Default Printer'
    $setDefaultPrinterButton.Location = New-Object Drawing.Point($horizontalCenter,110)
    $setDefaultPrinterButton.Size = New-Object Drawing.Size($buttonWidth, $buttonHeight)
    $setDefaultPrinterButton.Add_Click({ Show-SetDefaultPrinterDialog })

    $doneButton = New-Object Windows.Forms.Button
    $doneButton.Text = 'Done'
    $doneButton.Location = New-Object Drawing.Point($horizontalCenter,150)
    $doneButton.Size = New-Object Drawing.Size($buttonWidth, $buttonHeight)
    $doneButton.Add_Click({ $form.Close() })

    $form.Controls.AddRange(@($addPrinterButton, $removePrinterButton, $setDefaultPrinterButton, $doneButton))

    $form.ShowDialog()
}

# Function for Add Printer dialog
function Show-AddPrinterDialog {
    $addForm = New-Object Windows.Forms.Form
    $addForm.Text = 'Add Network Printer'
    $addForm.Size = New-Object Drawing.Size(350,180)
    $addForm.StartPosition = 'CenterScreen'
    $addForm.FormBorderStyle = 'FixedDialog'
    $addForm.MaximizeBox = $false
    $addForm.MinimizeBox = $false

    # Printer Name Label
    $nameLabel = New-Object Windows.Forms.Label
    $nameLabel.Text = 'Printer Name:'
    $nameLabel.Location = New-Object Drawing.Point(10,20)
    $nameLabel.Size = New-Object Drawing.Size(100,20)

    # Printer Name TextBox
    $nameTextBox = New-Object Windows.Forms.TextBox
    $nameTextBox.Location = New-Object Drawing.Point(120,20)
    $nameTextBox.Size = New-Object Drawing.Size(200,20)

    # Label for Patience Message
    $patienceLabel = New-Object Windows.Forms.Label
    $patienceLabel.Text = 'Adding the printer may take some time. Please be patient.'
    $patienceLabel.Location = New-Object Drawing.Point(10,40)
    $patienceLabel.Size = New-Object Drawing.Size(360,40)
    $patienceLabel.ForeColor = [System.Drawing.Color]::Red
    $patienceLabel.Font = New-Object Drawing.Font('Arial', 10, [System.Drawing.FontStyle]::Italic)

    # Add Button
    $addButton = New-Object Windows.Forms.Button
    $addButton.Text = 'Add'
    $addButton.Location = New-Object Drawing.Point(120,100)
    $addButton.Size = New-Object Drawing.Size(100,30)
    $addButton.Add_Click({
        $printerName = $nameTextBox.Text
        $printerPath = "\\prtsvr2\$printerName"

        if ([string]::IsNullOrEmpty($printerName)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a printer name.")
        } else {
            try {
                # Try to add the network printer
                Add-Printer -ConnectionName $printerPath -ErrorAction Stop
                [System.Windows.Forms.MessageBox]::Show("Printer '$printerPath' added successfully.")

                # Ask if the printer should be set as default
                $result = [System.Windows.Forms.MessageBox]::Show("Do you want to set '$printerPath' as the default printer?", "Set Default Printer", [System.Windows.Forms.MessageBoxButtons]::YesNo)

                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                    try {
                        # Try to set the printer as default
                        Set-DefaultPrinter -PrinterName $printerName
                        [System.Windows.Forms.MessageBox]::Show("Printer '$printerPath' set as default.")
                    }
                    catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to set '$printerPath' as default printer.")
                    }
                }

                $addForm.Close()
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Printer '$printerPath' not found. Please check the printer name.")
            }
        }
    })

    $addForm.Controls.AddRange(@($nameLabel, $nameTextBox, $addButton))
    $addForm.ShowDialog()
}

function Set-DefaultPrinter {
    param (
        [string]$printerName
    )
    Write-Host "Setting default to $printerName"
    # Set default printer using WScript.Network
    $printer = "\\prtsvr2\"+$printerName
    (New-Object -ComObject WScript.Network).SetDefaultPrinter($printer)
}

# Function for Remove Printer dialog
function Show-RemovePrinterDialog {
    $removeForm = New-Object Windows.Forms.Form
    $removeForm.Text = 'Remove Network Printers'
    $removeForm.Size = New-Object Drawing.Size(400,300)
    $removeForm.StartPosition = 'CenterScreen'
    $removeForm.FormBorderStyle = 'FixedDialog'
    $removeForm.MaximizeBox = $false
    $removeForm.MinimizeBox = $false

    # ListBox to display printers
    $printerListBox = New-Object Windows.Forms.ListBox
    $printerListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiSimple
    $printerListBox.Location = New-Object Drawing.Point(10,10)
    $printerListBox.Size = New-Object Drawing.Size(360,200)

    # Load printers hosted on prtsvr2 and exclude "Secure Print"
    $printers = Get-CimInstance -ClassName Win32_Printer | Where-Object {
        $_.Name -like "\\prtsvr2\*" -and $_.Name -notlike "*Secure Print*"
    }

    foreach ($printer in $printers) {
        $printerListBox.Items.Add($printer.Name)
    }

    # Remove Button
    $removeButton = New-Object Windows.Forms.Button
    $removeButton.Text = 'Remove'
    $removeButton.Location = New-Object Drawing.Point(270,220)
    $removeButton.Size = New-Object Drawing.Size(100,30)
    $removeButton.Add_Click({
        $selectedPrinters = $printerListBox.SelectedItems
        if ($selectedPrinters.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one printer to remove.")
        } else {
            foreach ($printer in $selectedPrinters) {
                try {
                    Remove-Printer -Name $printer -ErrorAction Stop
                    [System.Windows.Forms.MessageBox]::Show("Printer '$printer' removed successfully.")
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Failed to remove printer '$printer'.")
                }
            }
            $removeForm.Close()
        }
    })

    $removeForm.Controls.AddRange(@($printerListBox, $removeButton))
    $removeForm.ShowDialog()
}

# Function for Set Default Printer dialog
function Show-SetDefaultPrinterDialog {
    $defaultForm = New-Object Windows.Forms.Form
    $defaultForm.Text = 'Set Default Printer'
    $defaultForm.Size = New-Object Drawing.Size(400,300)
    $defaultForm.StartPosition = 'CenterScreen'
    $defaultForm.FormBorderStyle = 'FixedDialog'
    $defaultForm.MaximizeBox = $false
    $defaultForm.MinimizeBox = $false

    # ListBox to display printers
    $printerListBox = New-Object Windows.Forms.ListBox
    $printerListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::One
    $printerListBox.Location = New-Object Drawing.Point(10,10)
    $printerListBox.Size = New-Object Drawing.Size(360,200)

    # Load printers hosted on prtsvr2
    $printers = Get-CimInstance -ClassName Win32_Printer | Where-Object {
        $_.Name -like "\\prtsvr2\*"
    }

    foreach ($printer in $printers) {
        $printerListBox.Items.Add($printer.Name)
    }

    # Set Default Button
    $setDefaultButton = New-Object Windows.Forms.Button
    $setDefaultButton.Text = 'Set Default'
    $setDefaultButton.Location = New-Object Drawing.Point(270,220)
    $setDefaultButton.Size = New-Object Drawing.Size(100,30)
    $setDefaultButton.Add_Click({
        $selectedPrinter = $printerListBox.SelectedItem
        if (-not $selectedPrinter) {
            [System.Windows.Forms.MessageBox]::Show("Please select a printer to set as default.")
        } else {
            try {
                # Strip off "\\prtsvr2\" prefix correctly
                $printerName = $selectedPrinter -replace '\\prtsvr2\\', ''
                
                # Ensure the printer name starts with a backslash if needed
                $printerName = $printerName.TrimStart('\')

                # Debugging output
                Write-Host "Selected Printer: $selectedPrinter"
                Write-Host "Printer Name to Set as Default: $printerName"
                
                # Use the existing Set-DefaultPrinter function
                Set-DefaultPrinter -PrinterName $printerName
                
                [System.Windows.Forms.MessageBox]::Show("Printer '$selectedPrinter' set as default.")
                $defaultForm.Close()
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to set '$selectedPrinter' as default printer.")
                Write-Host "Error: $_"
            }
        }
    })

    $defaultForm.Controls.AddRange(@($printerListBox, $setDefaultButton))
    $defaultForm.ShowDialog()
}

# Run the main dialog
Show-MainDialog
