$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

Add-Type -AssemblyName 'System.Windows.Forms'

Set-PSDebug -Trace 1

# trap { Write-Error "Error found: $_" }

$isAdmin = ( ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).
        IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) )
if (-Not $isAdmin) {
        [System.Windows.Forms.MessageBox]::Show("Elevetion is required!")
    # Start-Process PowerShell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"cd '$pwd'; & '$PSCommandPath';`"";
    Exit;
}

$form = New-Object System.Windows.Forms.Form -Property @{
    Text = "MsHelpViewer-Bootstrapper"
    Width = 500
    Height = 200
    AutoSize = $true
    StartPosition = "CenterScreen"
    FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::SizableToolWindow
    BackColor = [System.Drawing.Color]::WhiteSmoke
}


$layout = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
    Padding = New-Object System.Windows.Forms.Padding(4)
    Dock = [System.Windows.Forms.DockStyle]::Fill  # Make it fill the form
    CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Single
    BackColor = [System.Drawing.Color]::WhiteSmoke
    RowCount = 4
    ColumnCount = 3
}
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
# $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$layout.SuspendLayout()
$form.Controls.Add($layout)

$fontSize = 9
$font = New-Object System.Drawing.Font("Arial", $fontSize)

$verLbl = New-Object System.Windows.Forms.Label -Property @{
    Text = "Help Viewer ver:"
    # Dock = [System.Windows.Forms.DockStyle]::Fill
    # TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    # AutoSize = $true
    Font = $font
}
$layout.Controls.Add($verLbl, 0, 0)
# $layout.SetFlowBreak($verLbl, $true)

$helpVers = Get-ChildItem -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Help"
if ($helpVers.Length -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No MS Help viewer detected")
    exit 1
}

$verSel = New-Object System.Windows.Forms.ComboBox -Property @{
#    Location = New-Object System.Drawing.Point(200, 50)
    Size = New-Object System.Drawing.Size(100, 30)
    DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    Font = $font
}

foreach ($ver in $helpVers) {
    $verSel.Items.Add($ver.PSChildName)
}
$verSel.SelectedIndex = $verSel.Items.Count - 1
$layout.Controls.Add($verSel, 1, 0)

$nameLbl = New-Object System.Windows.Forms.Label -Property @{
    Text = "Catalog name:"
    # AutoSize = $true
    Font = $font
}
$layout.Controls.Add($nameLbl, 0, 1)

$nameBox = New-Object System.Windows.Forms.TextBox -Property @{
    Size = New-Object System.Drawing.Size(100, 30)
    Font = $font
    Text = "msvc"
}
$layout.Controls.Add($nameBox, 1, 1)

$dirLbl = New-Object System.Windows.Forms.Label -Property @{
    Text = "Docs dir:"
    AutoSize = $true
    Font = $font
}
$layout.Controls.Add($dirLbl, 0, 2)

$dirBox = New-Object System.Windows.Forms.TextBox -Property @{
    Size = New-Object System.Drawing.Size(300, 30)
    Font = $font
    Text = ""
}
$layout.Controls.Add($dirBox, 1, 2)

$dirBtn = New-Object System.Windows.Forms.Button -Property @{
    Text = "..."
    Size = New-Object System.Drawing.Size(25, 25)
}
$layout.Controls.Add($dirBtn, 2, 2)

$dirBtn.Add_Click({
    $name = $Script:nameBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) {
        [System.Windows.Forms.MessageBox]::Show("Empty Catalog name!")
        return;
    }
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        SelectedPath = "c:\"
        Description = "Select a target folder"
        ShowNewFolderButton = $true
        RootFolder = [System.Environment+SpecialFolder]::Desktop
    }
    $result = $FolderBrowser.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $path = $FolderBrowser.SelectedPath.TrimEnd("\")
        ${Script:dirBox}.Text = "$path\$name"
    }
})

$butLayout = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{
    Dock = [System.Windows.Forms.DockStyle]::Fill
    Padding = New-Object System.Windows.Forms.Padding(2)
    FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    WrapContents = $false
    AutoSize = $true
    BackColor = $form.BackColor
    AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
}
$layout.Controls.Add($butLayout, 1, 3)

$createBtn = New-Object System.Windows.Forms.Button -Property @{
    Text = "Create"
    # Dock = [System.Windows.Forms.DockStyle]::Left
    Size = New-Object System.Drawing.Size(50, 20)
}
$butLayout.Controls.Add($createBtn)

$createBtn.Add_Click({
    $path = $dirBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($path)) {
        [System.Windows.Forms.MessageBox]::Show("Empty path!")
        return;
    }

    :skip_dir do {
        try {
            New-Item -ItemType Directory -ErrorAction Stop -Path $path
        } catch {
            if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ResourceExists) {
                break skip_dir
            }
            [System.Windows.Forms.MessageBox]::Show("Cannot create '$path'`n$($_.Exception.Message)")
            return;
        }
    } while ($false)

    $catXmlPath = "$path\CatalogType.xml"
    try {
        '<?xml version="1.0" encoding="utf-8"?><catalogType>UserManaged</catalogType>' `
            | Set-Content -Path $catXmlPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Cannot create '$catXmlPath'`n$($_.Exception.Message)")
        return;
    }

    $cat = $nameBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($cat)) {
        [System.Windows.Forms.MessageBox]::Show("Empty Catalog name!")
        return;
    }

    $regApp = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Help\$($Script:verSel.Text)"
    $regCat = "$regApp\Catalogs\$cat"

    :skip_cat1 do {
        try {
            New-Item -ItemType RegistryKey -Path $regCat
        }
        catch {
            if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ResourceExists) {
                break skip_cat1
            }
            [System.Windows.Forms.MessageBox]::Show("Cannot create '$regCat'!`n$($_.Exception.Message)")
            return;
        }
    } while ($false)
    :skip_cat2 do {
        try {
            New-Item -ItemType RegistryKey -Path "$regCat\en-US"
        }
        catch {
            if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ResourceExists) {
                break skip_cat2
            }
            [System.Windows.Forms.MessageBox]::Show("Cannot create '$regCat\en-US'!`n$($_.Exception.Message)")
            return;
        }
    } while ($false)

    try {
        Set-ItemProperty -Path $regCat -Name "LocationPath" -Value "$path" -Type String
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Cannot write '$regCat LocationPath'!`n$($_.Exception.Message)")
        return;
    }

    try {
        $appRoot = (Get-ItemProperty -Path $regApp).AppRoot
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Cannot retrive '$regApp AppRoot'!`n$($_.Exception.Message)")
        return;
    }

    try {
        Set-ItemProperty -Path "$regCat\en-US" -Name "SeedFilePath" -Value "${appRoot}CatalogInfo\VS11_en-us.cab" -Type String
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Cannot write '$regCat\en-US SeedFilePath'!`n$($_.Exception.Message)")
        return;
    }

    try {
        Set-ItemProperty -Path "$regCat\en-US" -Name "catalogName" -Value "$cat docs" -Type String
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Cannot write '$regCat\en-US catalogName'!`n$($_.Exception.Message)")
        return;
    }

    $desktopPath = [System.Environment]::GetFolderPath('Desktop')
    $shortcutPath = Join-Path $desktopPath "$cat docs.lnk"
    $WshShell = New-Object -ComObject WScript.Shell
    $shortcut = $WshShell.CreateShortcut($shortcutPath)
    $shortcut.WorkingDirectory = $appRoot
    $shortcut.TargetPath = "$appRoot\HlpViewer.exe"
    $shortcut.Arguments = "/catalogName $cat /locale en-US"
    $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,-263"
    $shortcut.Save()
})

$closeBtn = New-Object System.Windows.Forms.Button -Property @{
    Text = "Close"
    # Dock = [System.Windows.Forms.DockStyle]::Right
    Size = New-Object System.Drawing.Size(50, 20)
    # Location = New-Object System.Drawing.Point(200,150)
}
$butLayout.Controls.Add($closeBtn)

$closeBtn.Add_Click({
    $form.Close()
})

$layout.ResumeLayout()
$form.ShowDialog()
