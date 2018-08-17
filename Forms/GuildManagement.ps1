 #CSV 

$Script     = $MyInvocation.MyCommand.Definition
$ScriptPath = Split-Path -Parent $Script

Add-Type -AssemblyName System.Drawing,System.Windows.Forms,Microsoft.VisualBasic

##### Hashtable(s)
$FormHash     = [Hashtable]::Synchronized(@{}) #Created in the Open-RunspacePool Function
$VariableHash = [Hashtable]::Synchronized(@{})
$SettingsHash = @{}
$DataHash     = @{}
$MetricsHash  = @{}

#- Setup
If ( !(Test-Path HKCU:\SOFTWARE\DnD) ) {
   $rKey = "HKCU:\SOFTWARE\DnD"
   $sLocation = "$env:USERPROFILE\DnD\"
   $iLocation = "$env:USERPROFILE\DnD\"
   $dLocation = "$env:USERPROFILE\DnD\"

   If ( !(Test-Path $sLocation) ) {New-Item $sLocation -ItemType Directory | Out-Null }
   If ( !(Test-Path $rKey) ) { New-Item -Path HKCU:\SOFTWARE\ -Name DnD | Out-Null }
   New-ItemProperty -Path $rKey -Name Import -Value $iLocation -Force   | Out-Null
   New-ItemProperty -Path $rKey -Name Settings -Value $sLocation -Force | Out-Null
   New-ItemProperty -Path $rKey -Name Data -Value $dLocation -Force     | Out-Null
   New-ItemProperty -Path $rKey -Name UserName -Value $env:USERNAME -Force  | Out-Null
   
   $wScript = New-Object -ComObject wScript.Shell
   $Shortcut = $wScript.CreateShortcut("$env:USERPROFILE\Desktop\Managment.lnk")
   $Shortcut.TargetPath = "$(Join-Path $ScriptPath GuildManagement.bat)"
   $Shortcut.WindowStyle  = 7
   $Shortcut.Save()
   
   $VariableHash.Setup = $True
}

If ($VariableHash.Setup -eq $True) {PS -PID $PID | KILL}

$SettingsHash.SettingsLocation = "$((Get-ItemProperty HKCU:\SOFTWARE\DnD).Settings)"
$SettingsHash.ImportLocation   = "$((Get-ItemProperty HKCU:\SOFTWARE\DnD).Import)"
$SettingsHash.DataLocation     = "$((Get-ItemProperty HKCU:\SOFTWARE\DnD).Data)"
$SettingsHash.UserName         = "$((Get-ItemProperty HKCU:\SOFTWARE\DnD).UserName)"

If ( Test-Path "$($SettingsHash.SettingsLocation)\GuildSettings.xml" ) { $SettingsHash.Guild = Import-Clixml -Path "$($SettingsHash.SettingsLocation)\GuildSettings.xml" }
Else {
   $SettingsHash.Guild = @{}
   $SettingsHash.Guild.Ranks = @{}
   $SettingsHash.Guild.Inactivity = @{}
}

If ($SettingsHash.SaveOnExit -ne $True) { $SettingsHash.SaveOnExit = $False }

If ( Test-Path "$($SettingsHash.DataLocation)\Data.xml" ) { $DataHash = Import-Clixml -Path "$($SettingsHash.DataLocation)\Data.xml" ; $VariableHash.DataImported = $True }
Else {$VariableHash.DataImported = $False}

If ( Test-Path "$($SettingsHash.DataLocation)\ClassMetrics.xml" ) { $MetricsHash.Classes = Import-Clixml -Path "$($SettingsHash.DataLocation)\ClassMetrics.xml" ; $VariableHash.ClassMetricsImported = $True }
Else {
   $VariableHash.ClassMetricsImported = $False
   $MetricsHash.Classes  = @{}
   $MetricsHash.Classes.ByAccount = @{}
   $MetricsHash.Classes.Totals    = @{}
}
##### Import Custom Modules
.(Join-Path $ScriptPath ..\Modules\Form_Elements.ps1)

$Form = Create-Form -Name "Form" -Size "1000,800" -ShowIcon 0 -Text "Management helper"
$Main = Create-MenuStrip -Name "Main" -Form $Form -MainMenuStrip -BackColor "#66c6caff"
   $FileMenu = Create-ToolStripMenuItem -Name "FileMenu" -Text "&File" -Form $Main -AutoSize -MenuItem -BackColor TRANSPARENT
      $FileImportCSV = Create-ToolStripMenuItem -Name "FileImportCSV" -Text "Import CSV" -Form $FileMenu -AutoSize -DropDownItem
      $FileSaveData = Create-ToolStripMenuItem -Name "FileSaveData" -Text "Save Data" -Form $FileMenu -AutoSize -DropDownItem -Available 1
      $FileLoadData = Create-ToolStripMenuItem -Name "FileLoadData" -Text "Load Data" -Form $FileMenu -AutoSize -DropDownItem -Available 0
      $FileExit = Create-ToolStripMenuItem -Name "FileExit" -Text "Quit" -Form $FileMenu -AutoSize -DropDownItem -ShortcutKeys "Control, Q"
   $ConfigurationMenu = Create-ToolStripMenuItem -Name "ConfigurationMenu" -Text "&Configure" -Form $Main -AutoSize -MenuItem -BackColor TRANSPARENT
      $ConfigurationSettings = Create-ToolStripMenuItem -Name "ConfigurationSettings" -Text "Configure Settings" -Form $ConfigurationMenu -AutoSize -DropDownItem
         $CS_Panel1 = Create-PanelBox -Name "CS_Panel1" -Size "1000,780" -Location "0,0" -Visible 0 -Form $Form -BackColor TRANSPARENT
            $CS_P1_GroupBox1 = Create-GroupBox -Name "CS_P1_GroupBox1" -Location "70,50" -Size "850,65" -Text "Import Location:" -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $CS_Panel1
               $CS_P1_GB1_RichTextBox1 = Create-RichTextBox -Name "CS_P1_GB1_RichTextBox1" -Location "8,25" -Size "655,30" -Text $SettingsHash.ImportLocation -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CS_P1_GroupBox1 -Hash $FormHash
               $CS_P1_GB1_Button1 = Create-Button -Name "CS_P1_GB1_Button1" -Location "$($CS_P1_GB1_RichTextBox1.Right+5),$($CS_P1_GB1_RichTextBox1.Top)" -Text "Browse..." -Size "175,30" -BackColor LIGHTGRAY -TextAlign BOTTOMCENTER -FontSize 14 -FlatStyle POPUP -Form $CS_P1_GroupBox1 -Hash $FormHash
            $CS_P1_GroupBox2 = Create-GroupBox -Name "CS_P1_GroupBox2" -Location "70,$($CS_P1_GroupBox1.Bottom+20)" -Size "850,65" -Text "Setting Location:" -FontSize 12 -BackColor TRANSPARENT -Form $CS_Panel1
               $CS_P1_GB2_RichTextBox1 = Create-RichTextBox -Name "CS_P1_GB2_RichTextBox1" -Location "8,25" -Size "655,30" -Text $SettingsHash.SettingsLocation -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CS_P1_GroupBox2 -Hash $FormHash
               $CS_P1_GB2_Button1 = Create-Button -Name "CS_P1_GB2_Button1" -Location "$($CS_P1_GB2_RichTextBox1.Right+5),$($CS_P1_GB2_RichTextBox1.Top)" -Text "Browse..." -Size "175,30" -BackColor LIGHTGRAY -TextAlign BOTTOMCENTER -FontSize 14 -FlatStyle POPUP -Form $CS_P1_GroupBox2 -Hash $FormHash
            $CS_P1_GroupBox3 = Create-GroupBox -Name "CS_P1_GroupBox3" -Location "70,$($CS_P1_GroupBox2.Bottom+20)" -Size "850,65" -Text "Data Location:" -FontSize 12 -BackColor TRANSPARENT -Form $CS_Panel1
               $CS_P1_GB3_RichTextBox1 = Create-RichTextBox -Name "CS_P1_GB3_RichTextBox1" -Location "8,25" -Size "655,30" -Text $SettingsHash.DataLocation -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CS_P1_GroupBox3 -Hash $FormHash
               $CS_P1_GB3_Button1 = Create-Button -Name "CS_P1_GB3_Button1" -Location "$($CS_P1_GB3_RichTextBox1.Right+5),$($CS_P1_GB3_RichTextBox1.Top)" -Text "Browse..." -Size "175,30" -BackColor LIGHTGRAY -TextAlign BOTTOMCENTER -FontSize 14 -FlatStyle POPUP -Form $CS_P1_GroupBox3 -Hash $FormHash
            $CS_P1_GroupBox4 = Create-GroupBox -Name "CS_P1_GroupBox4" -Location "70,$($CS_P1_GroupBox3.Bottom+20)" -Size "250,65" -Text "User name:" -FontSize 12 -BackColor TRANSPARENT -Form $CS_Panel1
               $CS_P1_GB4_RichTextBox1 = Create-RichTextBox -Name "CS_P1_GB4_RichTextBox1" -Location "8,25" -Size "236,30" -Text $SettingsHash.UserName -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CS_P1_GroupBox4 -Hash $FormHash
            $CS_P1_CheckBox1 = Create-CheckBox -Name "CS_P1_CheckBox1" -Location "70,$($CS_P1_GroupBox4.Bottom+20)" -AutoSize 1 -Text "Enable data autosave on application exit" -Checked $SettingsHash.SaveOnExit -FontSize 14 -Form $CS_Panel1
            $CS_P1_Button1 = Create-Button -Name "CS_P1_Button1" -Location "$($CS_Panel1.Width-150),$($CS_Panel1.Bottom-55)" -Text "Save Data" -Size "125,30" -BackColor LIGHTGRAY -TextAlign BOTTOMCENTER -FontSize 14 -FlatStyle POPUP -Form $CS_Panel1 -Hash $FormHash
      $ConfigurationGuild = Create-ToolStripMenuItem -Name "ConfigurationGuild" -Text "Configure Guild Settings" -Form $ConfigurationMenu -AutoSize -DropDownItem
         $CG_Panel1 = Create-PanelBox -Name "CR_Panel1" -Size "1000,780" -Location "0,0" -Visible 0 -Form $Form -BackColor TRANSPARENT
            $CG_P1_GroupBox1 = Create-GroupBox -Name "CG_P1_GroupBox1" -Location "30,50" -Size "600,260" -Text "Ranks:" -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $CG_Panel1
               $CG_P1_GB1_Label1 = Create-Label -Name "CG_P1_GB1_Label1" -Location "15,25" -Size "25,30" -Text "1:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox1
               $CG_P1_GB1_Label2 = Create-Label -Name "CG_P1_GB1_Label2" -Location "$($CG_P1_GB1_Label1.Left),$($CG_P1_GB1_Label1.Bottom+3)" -Size "25,30" -Text "2:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox1
               $CG_P1_GB1_Label3 = Create-Label -Name "CG_P1_GB1_Label3" -Location "$($CG_P1_GB1_Label2.Left),$($CG_P1_GB1_Label2.Bottom+3)" -Size "25,30" -Text "3:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox1
               $CG_P1_GB1_Label4 = Create-Label -Name "CG_P1_GB1_Label4" -Location "$($CG_P1_GB1_Label3.Left),$($CG_P1_GB1_Label3.Bottom+3)" -Size "25,30" -Text "4:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox1
               $CG_P1_GB1_Label5 = Create-Label -Name "CG_P1_GB1_Label5" -Location "$($CG_P1_GB1_Label4.Left),$($CG_P1_GB1_Label4.Bottom+3)" -Size "25,30" -Text "5:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox1
               $CG_P1_GB1_Label6 = Create-Label -Name "CG_P1_GB1_Label6" -Location "$($CG_P1_GB1_Label5.Left),$($CG_P1_GB1_Label5.Bottom+3)" -Size "25,30" -Text "6:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox1
               $CG_P1_GB1_Label7 = Create-Label -Name "CG_P1_GB1_Label7" -Location "$($CG_P1_GB1_Label6.Left),$($CG_P1_GB1_Label6.Bottom+3)" -Size "25,30" -Text "7:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox1
               $CG_P1_GB1_RichTextBox1 = Create-RichTextBox -Name "CG_P1_GB1_RichTextBox1" -Location "$($CG_P1_GB1_Label1.Right+5),$($CG_P1_GB1_Label1.Top)" -Size "535,30" -Text $SettingsHash.Guild.Ranks.Rank1 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox1 -Hash $FormHash
               $CG_P1_GB1_RichTextBox2 = Create-RichTextBox -Name "CG_P1_GB1_RichTextBox2" -Location "$($CG_P1_GB1_Label2.Right+5),$($CG_P1_GB1_Label2.Top)" -Size "535,30" -Text $SettingsHash.Guild.Ranks.Rank2 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox1 -Hash $FormHash
               $CG_P1_GB1_RichTextBox3 = Create-RichTextBox -Name "CG_P1_GB1_RichTextBox3" -Location "$($CG_P1_GB1_Label3.Right+5),$($CG_P1_GB1_Label3.Top)" -Size "535,30" -Text $SettingsHash.Guild.Ranks.Rank3 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox1 -Hash $FormHash
               $CG_P1_GB1_RichTextBox4 = Create-RichTextBox -Name "CG_P1_GB1_RichTextBox4" -Location "$($CG_P1_GB1_Label4.Right+5),$($CG_P1_GB1_Label4.Top)" -Size "535,30" -Text $SettingsHash.Guild.Ranks.Rank4 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox1 -Hash $FormHash
               $CG_P1_GB1_RichTextBox5 = Create-RichTextBox -Name "CG_P1_GB1_RichTextBox5" -Location "$($CG_P1_GB1_Label5.Right+5),$($CG_P1_GB1_Label5.Top)" -Size "535,30" -Text $SettingsHash.Guild.Ranks.Rank5 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox1 -Hash $FormHash
               $CG_P1_GB1_RichTextBox6 = Create-RichTextBox -Name "CG_P1_GB1_RichTextBox6" -Location "$($CG_P1_GB1_Label6.Right+5),$($CG_P1_GB1_Label6.Top)" -Size "535,30" -Text $SettingsHash.Guild.Ranks.Rank6 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox1 -Hash $FormHash
               $CG_P1_GB1_RichTextBox7 = Create-RichTextBox -Name "CG_P1_GB1_RichTextBox7" -Location "$($CG_P1_GB1_Label7.Right+5),$($CG_P1_GB1_Label7.Top)" -Size "535,30" -Text $SettingsHash.Guild.Ranks.Rank7 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox1 -Hash $FormHash
            $CG_P1_GroupBox2 = Create-GroupBox -Name "CG_P1_GroupBox2" -Location "$($CG_P1_GroupBox1.Right+15),$($CG_P1_GroupBox1.Top)" -Size "320,260" -Text "Inactivity (Days):" -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $CG_Panel1
               $CG_P1_GB2_Label1 = Create-Label -Name "CG_P1_GB2_Label1" -Location "15,25" -Size "25,30" -Text "1:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox2
               $CG_P1_GB2_Label2 = Create-Label -Name "CG_P1_GB2_Label2" -Location "$($CG_P1_GB2_Label1.Left),$($CG_P1_GB2_Label1.Bottom+3)" -Size "25,30" -Text "2:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox2
               $CG_P1_GB2_Label3 = Create-Label -Name "CG_P1_GB2_Label3" -Location "$($CG_P1_GB2_Label2.Left),$($CG_P1_GB2_Label2.Bottom+3)" -Size "25,30" -Text "3:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox2
               $CG_P1_GB2_Label4 = Create-Label -Name "CG_P1_GB2_Label4" -Location "$($CG_P1_GB2_Label3.Left),$($CG_P1_GB2_Label3.Bottom+3)" -Size "25,30" -Text "4:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox2
               $CG_P1_GB2_Label5 = Create-Label -Name "CG_P1_GB2_Label5" -Location "$($CG_P1_GB2_Label4.Left),$($CG_P1_GB2_Label4.Bottom+3)" -Size "25,30" -Text "5:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox2
               $CG_P1_GB2_Label6 = Create-Label -Name "CG_P1_GB2_Label6" -Location "$($CG_P1_GB2_Label5.Left),$($CG_P1_GB2_Label5.Bottom+3)" -Size "25,30" -Text "6:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox2
               $CG_P1_GB2_Label7 = Create-Label -Name "CG_P1_GB2_Label7" -Location "$($CG_P1_GB2_Label6.Left),$($CG_P1_GB2_Label6.Bottom+3)" -Size "25,30" -Text "7:" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $CG_P1_GroupBox2
               $CG_P1_GB2_RichTextBox1 = Create-RichTextBox -Name "CG_P1_GB2_RichTextBox1" -Location "$($CG_P1_GB2_Label1.Right+5),$($CG_P1_GB2_Label1.Top)" -Size "40,30" -Text $SettingsHash.Guild.Inactivity.Rank1 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox2 -Hash $FormHash
               $CG_P1_GB2_RichTextBox2 = Create-RichTextBox -Name "CG_P1_GB2_RichTextBox2" -Location "$($CG_P1_GB2_Label2.Right+5),$($CG_P1_GB2_Label2.Top)" -Size "40,30" -Text $SettingsHash.Guild.Inactivity.Rank2 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox2 -Hash $FormHash
               $CG_P1_GB2_RichTextBox3 = Create-RichTextBox -Name "CG_P1_GB2_RichTextBox3" -Location "$($CG_P1_GB2_Label3.Right+5),$($CG_P1_GB2_Label3.Top)" -Size "40,30" -Text $SettingsHash.Guild.Inactivity.Rank3 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox2 -Hash $FormHash
               $CG_P1_GB2_RichTextBox4 = Create-RichTextBox -Name "CG_P1_GB2_RichTextBox4" -Location "$($CG_P1_GB2_Label4.Right+5),$($CG_P1_GB2_Label4.Top)" -Size "40,30" -Text $SettingsHash.Guild.Inactivity.Rank4 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox2 -Hash $FormHash
               $CG_P1_GB2_RichTextBox5 = Create-RichTextBox -Name "CG_P1_GB2_RichTextBox5" -Location "$($CG_P1_GB2_Label5.Right+5),$($CG_P1_GB2_Label5.Top)" -Size "40,30" -Text $SettingsHash.Guild.Inactivity.Rank5 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox2 -Hash $FormHash
               $CG_P1_GB2_RichTextBox6 = Create-RichTextBox -Name "CG_P1_GB2_RichTextBox6" -Location "$($CG_P1_GB2_Label6.Right+5),$($CG_P1_GB2_Label6.Top)" -Size "40,30" -Text $SettingsHash.Guild.Inactivity.Rank6 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox2 -Hash $FormHash
               $CG_P1_GB2_RichTextBox7 = Create-RichTextBox -Name "CG_P1_GB2_RichTextBox7" -Location "$($CG_P1_GB2_Label7.Right+5),$($CG_P1_GB2_Label7.Top)" -Size "40,30" -Text $SettingsHash.Guild.Inactivity.Rank7 -FontSize 14 -BackColor WHITE -ForeColor BLACK -WordWrap 0 -MultiLine 0 -Form $CG_P1_GroupBox2 -Hash $FormHash
            $CG_P1_Button1 = Create-Button -Name "CG_P1_Button1" -Location "$($CG_Panel1.Width-150),$($CG_Panel1.Bottom-55)" -Text "Save Data" -Size "125,30" -BackColor LIGHTGRAY -TextAlign BOTTOMCENTER -FontSize 14 -FlatStyle POPUP -Form $CG_Panel1 -Hash $FormHash
   $ToolMenu = Create-ToolStripMenuItem -Name "ToolMenu" -Text "&Tools" -Form $Main -AutoSize -MenuItem -BackColor TRANSPARENT
      $RaffleTool = Create-ToolStripMenuItem -Name "RaffleTool" -Text "Raffle System" -Form $ToolMenu -AutoSize -DropDownItem -Available 0
      $Contributions = Create-ToolStripMenuItem -Name "Contributions" -Text "Contributions" -Form $ToolMenu -AutoSize -DropDownItem -Available 0
   $ViewMenu = Create-ToolStripMenuItem -Name "ViewMenu" -Text "&View" -Form $Main -AutoSize -MenuItem -BackColor TRANSPARENT
      $ViewAccounts = Create-ToolStripMenuItem -Name "ViewAccounts" -Text "Accounts" -Form $ViewMenu -AutoSize -DropDownItem
         $VA_Panel1 = Create-PanelBox -Name "VA_Panel1" -Size "1000,780" -Location "0,0" -Visible 0 -Form $Form -BackColor TRANSPARENT
            $VA_P1_GroupBox1 = Create-GroupBox -Name "VA_P1_GroupBox1" -Location "30,50" -Size "930,335" -Text "Accounts:" -FontSize 12 -BackColor TRANSPARENT -Form $VA_Panel1
               $VA_P1_GB1_CheckBox1 = Create-CheckBox -Name "VA_P1_GB1_CheckBox1" -Location "15,25" -AutoSize 1 -Text "Filter out all `"Non Members`" " -Form $VA_P1_GroupBox1
               $VA_P1_GB1_ListBox1 = Create-ListBox -Name "VA_P1_GB1_ListBox1" -Location "$($VA_P1_GB1_CheckBox1.Left),$($VA_P1_GB1_CheckBox1.Bottom+5)" -Size "220,280" -BackColor "#FFF0F0F0" -FontSize 14 -Sorted 1 -Form $VA_P1_GroupBox1
               $VA_P1_GB1_GroupBox1 = Create-GroupBox -Name "VA_P1_GB1_GroupBox1" -Location "$($VA_P1_GB1_CheckBox1.Right+20),$($VA_P1_GB1_CheckBox1.Top-2)" -Size "640,300" -Text "Details:" -FontSize 12 -BackColor TRANSPARENT -Form $VA_P1_GroupBox1
                  $VA_P1_GB1_GB1_Label1 = Create-Label -Name "VA_P1_GB1_GB1_Label1" -Location "10,25" -AutoSize 1 -Text "Rank:" -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label01 = Create-Label -Name "VA_P1_GB1_GB1_Label01" -Location "$($VA_P1_GB1_GB1_Label1.Right),$($VA_P1_GB1_GB1_Label1.Top-3)" -AutoSize 1 -FontSize 14 -TextAlign BOTTOMLEFT -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label2 = Create-Label -Name "VA_P1_GB1_GB1_Label2" -Location "$($VA_P1_GB1_GB1_Label1.Left),$($VA_P1_GB1_GB1_Label1.Bottom)" -AutoSize 1 -Text "Join Date:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label02 = Create-Label -Name "VA_P1_GB1_GB1_Label02" -Location "$($VA_P1_GB1_GB1_Label2.Right),$($VA_P1_GB1_GB1_Label2.Top-1)" -AutoSize 1 -FontSize 14 -TextAlign BOTTOMLEFT -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label3 = Create-Label -Name "VA_P1_GB1_GB1_Label3" -Location "$($VA_P1_GB1_GB1_Label1.Left),$($VA_P1_GB1_GB1_Label2.Bottom)" -AutoSize 1 -Text "Last Logon:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label03 = Create-Label -Name "VA_P1_GB1_GB1_Label03" -Location "$($VA_P1_GB1_GB1_Label3.Right),$($VA_P1_GB1_GB1_Label3.Top-1)" -AutoSize 1 -FontSize 14 -TextAlign BOTTOMLEFT -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label4 = Create-Label -Name "VA_P1_GB1_GB1_Label4" -Location "$($VA_P1_GB1_GB1_Label1.Left),$($VA_P1_GB1_GB1_Label3.Bottom)" -AutoSize 1 -Text "Last Rank Change:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label04 = Create-Label -Name "VA_P1_GB1_GB1_Label04" -Location "$($VA_P1_GB1_GB1_Label4.Right),$($VA_P1_GB1_GB1_Label4.Top-1)" -AutoSize 1 -FontSize 14 -TextAlign BOTTOMLEFT -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label5 = Create-Label -Name "VA_P1_GB1_GB1_Label5" -Location "$($VA_P1_GB1_GB1_Label1.Left),$($VA_P1_GB1_GB1_Label4.Bottom)" -AutoSize 1 -Text "Notes:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label6 = Create-Label -Name "VA_P1_GB1_GB1_Label6" -Location "$($VA_P1_GB1_GroupBox1.Width-90),4" -AutoSize 1 -Text "Status:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_Label06 = Create-Label -Name "VA_P1_GB1_GB1_Label06" -Location "$($VA_P1_GB1_GroupBox1.Width-120),$($VA_P1_GB1_GB1_Label6.Bottom+3)" -Width 120 -FontSize 14 -TextAlign BOTTOMCENTER -Form $VA_P1_GB1_GroupBox1
                  $VA_P1_GB1_GB1_RichTextBox1 = Create-RichTextBox -Name "VA_P1_GB1_GB1_RichTextBox1" -Location "$($VA_P1_GB1_GB1_Label1.Left+10),$($VA_P1_GB1_GB1_Label5.Bottom)" -Size "620,155" -ScrBar VERTICAL -BackColor "#FFD0D0D0" -FontSize 13 -ForeColor BLACK -MultiLine 1 -WordWrap 1 -Form $VA_P1_GB1_GroupBox1
            $VA_P1_GroupBox2 = Create-GroupBox -Name "VA_P1_GroupBox2" -Location "$($VA_P1_GroupBox1.Left),$($VA_P1_GroupBox1.Bottom+20)" -Size "930,335" -Text "Characters:" -FontSize 12 -BackColor TRANSPARENT -Form $VA_Panel1
               $VA_P1_GB2_CheckBox1 = Create-CheckBox -Name "VA_P1_GB2_CheckBox1" -Location "15,25" -AutoSize 1 -Text "Filter out all `"Non Members`" " -Form $VA_P1_GroupBox2
               $VA_P1_GB2_ListBox1 = Create-ListBox -Name "VA_P1_GB2_ListBox1" -Location "$($VA_P1_GB2_CheckBox1.Left),$($VA_P1_GB2_CheckBox1.Bottom+5)" -Size "220,280" -BackColor "#FFF0F0F0" -FontSize 14 -Sorted 1 -Form $VA_P1_GroupBox2
               $VA_P1_GB2_GroupBox1 = Create-GroupBox -Name "VA_P1_GB2_GroupBox1" -Location "$($VA_P1_GB2_CheckBox1.Right+20),$($VA_P1_GB2_CheckBox1.Top-2)" -Size "640,300" -Text "Details:" -FontSize 12 -BackColor TRANSPARENT -Form $VA_P1_GroupBox2
                  $VA_P1_GB2_GB1_Label1 = Create-Label -Name "VA_P1_GB2_GB1_Label1" -Location "10,25" -AutoSize 1 -Text "Rank:" -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label01 = Create-Label -Name "VA_P1_GB2_GB1_Label01" -Location "$($VA_P1_GB2_GB1_Label1.Right),$($VA_P1_GB2_GB1_Label1.Top-3)" -AutoSize 1 -FontSize 14 -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label2 = Create-Label -Name "VA_P1_GB2_GB1_Label2" -Location "$($VA_P1_GB2_GB1_Label1.Left),$($VA_P1_GB2_GB1_Label1.Bottom)" -AutoSize 1 -Text "Character:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label02 = Create-Label -Name "VA_P1_GB2_GB1_Label02" -Location "$($VA_P1_GB2_GB1_Label2.Right),$($VA_P1_GB2_GB1_Label2.Top-2)" -AutoSize 1 -FontSize 14 -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label3 = Create-Label -Name "VA_P1_GB2_GB1_Label3" -Location "$($VA_P1_GB2_GB1_Label1.Left),$($VA_P1_GB2_GB1_Label2.Bottom)" -AutoSize 1 -Text "Last Logon:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label03 = Create-Label -Name "VA_P1_GB2_GB1_Label03" -Location "$($VA_P1_GB2_GB1_Label3.Right),$($VA_P1_GB2_GB1_Label3.Top-2)" -AutoSize 1 -FontSize 14 -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label4 = Create-Label -Name "VA_P1_GB2_GB1_Label4" -Location "$($VA_P1_GB2_GB1_Label1.Left),$($VA_P1_GB2_GB1_Label3.Bottom)" -AutoSize 1 -Text "Public Comment:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label04 = Create-Label -Name "VA_P1_GB2_GB1_Label04" -Location "$($VA_P1_GB2_GB1_Label4.Right),$($VA_P1_GB2_GB1_Label4.Top+4)" -Width 485 -FontSize 14 -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label5 = Create-Label -Name "VA_P1_GB2_GB1_Label5" -Location "$($VA_P1_GB2_GB1_Label1.Left),$($VA_P1_GB2_GB1_Label4.Bottom)" -AutoSize 1 -Text "Officer Comment:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label05 = Create-Label -Name "VA_P1_GB2_GB1_Label05" -Location "$($VA_P1_GB2_GB1_Label5.Right),$($VA_P1_GB2_GB1_Label5.Top+4)" -Width 485 -FontSize 14 -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label6 = Create-Label -Name "VA_P1_GB2_GB1_Label6" -Location "$($VA_P1_GB2_GB1_Label1.Left),$($VA_P1_GB2_GB1_Label5.Bottom)" -AutoSize 1 -Text "Notes:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label7 = Create-Label -Name "VA_P1_GB2_GB1_Label7" -Location "$($VA_P1_GB2_GroupBox1.Width-90),4" -AutoSize 1 -Text "Status:" -TextAlign BOTTOMLEFT -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_Label07 = Create-Label -Name "VA_P1_GB2_GB1_Label07" -Location "$($VA_P1_GB2_GroupBox1.Width-120),$($VA_P1_GB2_GB1_Label7.Bottom+3)" -Width 120 -FontSize 14 -TextAlign BOTTOMCENTER -Form $VA_P1_GB2_GroupBox1
                  $VA_P1_GB2_GB1_RichTextBox1 = Create-RichTextBox -Name "VA_P1_GB2_GB1_RichTextBox1" -Location "$($VA_P1_GB2_GB1_Label1.Left+10),$($VA_P1_GB2_GB1_Label6.Bottom)" -Size "620,131" -ScrBar VERTICAL -BackColor "#FFD0D0D0" -FontSize 13 -ForeColor BLACK -MultiLine 1 -WordWrap 1 -Form $VA_P1_GB2_GroupBox1            
      $ViewClasses = Create-ToolStripMenuItem -Name "ViewClasses" -Text "Classes" -Form $ViewMenu -AutoSize -DropDownItem
         $VC_Panel1 = Create-PanelBox -Name "VC_Panel1" -Size "1000,780" -Location "0,0" -Visible 0 -Form $Form -BackColor TRANSPARENT
            $VC_P1_GroupBox1 = Create-GroupBox -Name "VC_P1_GroupBox1" -Location "30,50" -Size "450,290" -Text "Totals:" -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VC_Panel1
               $VC_P1_GB1_Label1 = Create-Label -Name "VC_P1_GB1_Label1" -Location "15,25" -Size "380,30" -Text "Control Wizard(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label2 = Create-Label -Name "VC_P1_GB1_Label2" -Location "$($VC_P1_GB1_Label1.Left),$($VC_P1_GB1_Label1.Bottom+3)" -Size "380,30" -Text "Devoted Cleric(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label3 = Create-Label -Name "VC_P1_GB1_Label3" -Location "$($VC_P1_GB1_Label2.Left),$($VC_P1_GB1_Label2.Bottom+3)" -Size "380,30" -Text "Great Weapons Fighter(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label4 = Create-Label -Name "VC_P1_GB1_Label4" -Location "$($VC_P1_GB1_Label3.Left),$($VC_P1_GB1_Label3.Bottom+3)" -Size "380,30" -Text "Guardian Fighter(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label5 = Create-Label -Name "VC_P1_GB1_Label5" -Location "$($VC_P1_GB1_Label4.Left),$($VC_P1_GB1_Label4.Bottom+3)" -Size "380,30" -Text "Hunter Ranger(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label6 = Create-Label -Name "VC_P1_GB1_Label6" -Location "$($VC_P1_GB1_Label5.Left),$($VC_P1_GB1_Label5.Bottom+3)" -Size "380,30" -Text "Oathbound Paladin(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label7 = Create-Label -Name "VC_P1_GB1_Label7" -Location "$($VC_P1_GB1_Label6.Left),$($VC_P1_GB1_Label6.Bottom+3)" -Size "380,30" -Text "Scourge Warlock(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label8 = Create-Label -Name "VC_P1_GB1_Label8" -Location "$($VC_P1_GB1_Label7.Left),$($VC_P1_GB1_Label7.Bottom+3)" -Size "380,30" -Text "Trickster Rogue(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label01 = Create-Label -Name "VC_P1_GB1_Label01" -Location "$($VC_P1_GB1_Label1.Right),$($VC_P1_GB1_Label1.Top)"       -Size "40,30" -Text $MetricsHash.Classes.Totals.CW  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label02 = Create-Label -Name "VC_P1_GB1_Label02" -Location "$($VC_P1_GB1_Label01.Left),$($VC_P1_GB1_Label01.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.Totals.DC  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label03 = Create-Label -Name "VC_P1_GB1_Label03" -Location "$($VC_P1_GB1_Label02.Left),$($VC_P1_GB1_Label02.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.Totals.GWF -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label04 = Create-Label -Name "VC_P1_GB1_Label04" -Location "$($VC_P1_GB1_Label03.Left),$($VC_P1_GB1_Label03.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.Totals.GF  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label05 = Create-Label -Name "VC_P1_GB1_Label05" -Location "$($VC_P1_GB1_Label04.Left),$($VC_P1_GB1_Label04.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.Totals.HR  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label06 = Create-Label -Name "VC_P1_GB1_Label06" -Location "$($VC_P1_GB1_Label05.Left),$($VC_P1_GB1_Label05.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.Totals.OP  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label07 = Create-Label -Name "VC_P1_GB1_Label07" -Location "$($VC_P1_GB1_Label06.Left),$($VC_P1_GB1_Label06.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.Totals.SW  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox1
               $VC_P1_GB1_Label08 = Create-Label -Name "VC_P1_GB1_Label08" -Location "$($VC_P1_GB1_Label07.Left),$($VC_P1_GB1_Label07.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.Totals.TR  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox1
            $VC_P1_GroupBox2 = Create-GroupBox -Name "VC_P1_GroupBox2" -Location "$($VC_P1_GroupBox1.Right+25),$($VC_P1_GroupBox1.Top)" -Size "450,290" -Text "Accounts that play:" -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VC_Panel1
               $VC_P1_GB2_Label1 = Create-Label -Name "VC_P1_GB2_Label1" -Location "15,25" -Size "380,30" -Text "Control Wizard(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label2 = Create-Label -Name "VC_P1_GB2_Label2" -Location "$($VC_P1_GB2_Label1.Left),$($VC_P1_GB2_Label1.Bottom+3)" -Size "380,30" -Text "Devoted Cleric(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label3 = Create-Label -Name "VC_P1_GB2_Label3" -Location "$($VC_P1_GB2_Label2.Left),$($VC_P1_GB2_Label2.Bottom+3)" -Size "380,30" -Text "Great Weapons Fighter(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label4 = Create-Label -Name "VC_P1_GB2_Label4" -Location "$($VC_P1_GB2_Label3.Left),$($VC_P1_GB2_Label3.Bottom+3)" -Size "380,30" -Text "Guardian Fighter(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label5 = Create-Label -Name "VC_P1_GB2_Label5" -Location "$($VC_P1_GB2_Label4.Left),$($VC_P1_GB2_Label4.Bottom+3)" -Size "380,30" -Text "Hunter Ranger(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label6 = Create-Label -Name "VC_P1_GB2_Label6" -Location "$($VC_P1_GB2_Label5.Left),$($VC_P1_GB2_Label5.Bottom+3)" -Size "380,30" -Text "Oathbound Paladin(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label7 = Create-Label -Name "VC_P1_GB2_Label7" -Location "$($VC_P1_GB2_Label6.Left),$($VC_P1_GB2_Label6.Bottom+3)" -Size "380,30" -Text "Scourge Warlock(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label8 = Create-Label -Name "VC_P1_GB2_Label8" -Location "$($VC_P1_GB2_Label7.Left),$($VC_P1_GB2_Label7.Bottom+3)" -Size "380,30" -Text "Trickster Rogue(s):" -BackColor TRANSPARENT -ForeColor BLACK -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label01 = Create-Label -Name "VC_P1_GB2_Label01" -Location "$($VC_P1_GB2_Label1.Right),$($VC_P1_GB2_Label1.Top)"       -Size "40,30" -Text $MetricsHash.Classes.ByAccount.CW  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label02 = Create-Label -Name "VC_P1_GB2_Label02" -Location "$($VC_P1_GB2_Label01.Left),$($VC_P1_GB2_Label01.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.ByAccount.DC  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label03 = Create-Label -Name "VC_P1_GB2_Label03" -Location "$($VC_P1_GB2_Label02.Left),$($VC_P1_GB2_Label02.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.ByAccount.GWF -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label04 = Create-Label -Name "VC_P1_GB2_Label04" -Location "$($VC_P1_GB2_Label03.Left),$($VC_P1_GB2_Label03.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.ByAccount.GF  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label05 = Create-Label -Name "VC_P1_GB2_Label05" -Location "$($VC_P1_GB2_Label04.Left),$($VC_P1_GB2_Label04.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.ByAccount.HR  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label06 = Create-Label -Name "VC_P1_GB2_Label06" -Location "$($VC_P1_GB2_Label05.Left),$($VC_P1_GB2_Label05.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.ByAccount.OP  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label07 = Create-Label -Name "VC_P1_GB2_Label07" -Location "$($VC_P1_GB2_Label06.Left),$($VC_P1_GB2_Label06.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.ByAccount.SW  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox2
               $VC_P1_GB2_Label08 = Create-Label -Name "VC_P1_GB2_Label08" -Location "$($VC_P1_GB2_Label07.Left),$($VC_P1_GB2_Label07.Bottom+3)" -Size "40,30" -Text $MetricsHash.Classes.ByAccount.TR  -BackColor TRANSPARENT -ForeColor BLACK -TextAlign MIDDLERIGHT -FontSize 14 -Form $VC_P1_GroupBox2
            $VC_P1_GroupBox3 = Create-GroupBox -Name "VC_P1_GroupBox3" -Location "$($VC_P1_GroupBox1.Left),$($VC_P1_GroupBox1.Bottom+15)" -Size "925,355" -Text "Search by Class:" -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VC_Panel1
               $VC_P1_GB3_ComboBox1 = Create-ComboBox -Name "VC_P1_GB3_ComboBox1" -Location "15,25" -Size "250,30" -FontSize 14 -Form $VC_P1_GroupBox3
                  $VC_P1_GB3_ComboBox1.Items.AddRange(@("Control Wizard","Devoted Cleric", "Great Weapon Fighter", "Guardian Fighter", "Hunter Ranger", "Oathbound Paladin", "Scourge Warlock", "Trickster Rogue"))
               $VC_P1_GB3_ListBox1 = Create-ListBox -Name "VC_P1_GB3_ListBox1" -Location "$($VC_P1_GB3_ComboBox1.Left),$($VC_P1_GB3_ComboBox1.Bottom+27)" -Size "250,280" -BackColor "#FFF0F0F0" -FontSize 14 -Sorted 1 -Form $VC_P1_GroupBox3
               $VC_P1_GB3_ListBox2 = Create-ListBox -Name "VC_P1_GB3_ListBox1" -Location "$($VC_P1_GB3_ComboBox1.Right+15),$($VC_P1_GB3_ComboBox1.Top)" -Size "220,320" -BackColor "#FFF0F0F0" -FontSize 14 -Sorted 1 -Form $VC_P1_GroupBox3
               $VC_P1_GB3_GroupBox1 = Create-GroupBox -Name "VC_P1_GB3_GroupBox1" -Location "$($VC_P1_GB3_ListBox2.Right+15),$($VC_P1_GB3_ListBox2.Top -10)" -Size "395,327" -Text "Details:" -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VC_P1_GroupBox3
                  $VC_P1_GB3_GB1_Label1 = Create-Label -Name "VC_P1_GB3_GB1_Label1" -Location "10,25" -AutoSize 1 -Text "Rank:" -Form $VC_P1_GB3_GroupBox1
                  $VC_P1_GB3_GB1_Label01 = Create-Label -Name "VC_P1_GB3_GB1_Label01" -Location "$($VC_P1_GB3_GB1_Label1.Right),$($VC_P1_GB3_GB1_Label1.Top-2)" -AutoSize 1 -TextAlign BOTTOMLEFT -Form $VC_P1_GB3_GroupBox1
                  $VC_P1_GB3_GB1_Label2 = Create-Label -Name "VC_P1_GB3_GB1_Label2" -Location "$($VC_P1_GB3_GB1_Label1.Left),$($VC_P1_GB3_GB1_Label1.Bottom)" -AutoSize 1 -Text "Level:" -TextAlign BOTTOMLEFT -Form $VC_P1_GB3_GroupBox1
                  $VC_P1_GB3_GB1_Label02 = Create-Label -Name "VC_P1_GB3_GB1_Label02" -Location "$($VC_P1_GB3_GB1_Label2.Right),$($VC_P1_GB3_GB1_Label2.Top)" -AutoSize 1 -TextAlign BOTTOMLEFT -Form $VC_P1_GB3_GroupBox1
                  $VC_P1_GB3_GB1_Label3 = Create-Label -Name "VC_P1_GB3_GB1_Label3" -Location "$($VC_P1_GB3_GB1_Label2.Left),$($VC_P1_GB3_GB1_Label2.Bottom)" -AutoSize 1 -Text "Member Comment:" -TextAlign BOTTOMLEFT -Form $VC_P1_GB3_GroupBox1
                  $VC_P1_GB3_GB1_Label03 = Create-Label -Name "VC_P1_GB3_GB1_Label03" -Location "$($VC_P1_GB3_GB1_Label3.Left+25),$($VC_P1_GB3_GB1_Label3.Bottom)" -Width 335 -TextAlign BOTTOMLEFT -Form $VC_P1_GB3_GroupBox1
                  $VC_P1_GB3_GB1_Label4 = Create-Label -Name "VC_P1_GB3_GB1_Label4" -Location "$($VC_P1_GB3_GB1_Label3.Left),$($VC_P1_GB3_GB1_Label03.Bottom)" -AutoSize 1 -Text "Officer Comment:" -TextAlign BOTTOMLEFT -Form $VC_P1_GB3_GroupBox1
                  $VC_P1_GB3_GB1_Label04 = Create-Label -Name "VC_P1_GB3_GB1_Label04" -Location "$($VC_P1_GB3_GB1_Label4.Left+25),$($VC_P1_GB3_GB1_Label4.Bottom)" -Width 335 -TextAlign BOTTOMLEFT -Form $VC_P1_GB3_GroupBox1
                  $VC_P1_GB3_GB1_Label5 = Create-Label -Name "VC_P1_GB3_GB1_Label5" -Location "$($VC_P1_GB3_GB1_Label4.Left),$($VC_P1_GB3_GB1_Label04.Bottom)" -AutoSize 1 -Text "Notes:" -TextAlign BOTTOMLEFT -Form $VC_P1_GB3_GroupBox1
                  $VC_P1_GB3_GB1_Label05 = Create-Label -Name "VC_P1_GB3_GB1_Label05" -Location "$($VC_P1_GB3_GB1_Label5.Left+25),$($VC_P1_GB3_GB1_Label4.Bottom)" -Width 335 -TextAlign BOTTOMLEFT -Form $VC_P1_GB3_GroupBox1
            $VC_P1_Button1 = Create-Button -Name "VC_P1_Button1" -Location "$($VC_Panel1.Width-150),$($VC_Panel1.Bottom-55)" -Text "Update..." -Size "125,30" -BackColor LIGHTGRAY -TextAlign BOTTOMCENTER -FontSize 14 -FlatStyle POPUP -Form $VC_Panel1 -Hash $FormHash
      $ViewRanks = Create-ToolStripMenuItem -Name "ViewRanks" -Text "Ranks" -Form $ViewMenu -AutoSize -DropDownItem
         $VR_Panel1 = Create-PanelBox -Name "VR_Panel1" -Size "1000,780" -Location "0,0" -Visible 0 -Form $Form -BackColor TRANSPARENT
            $VR_P1_GroupBox1 = Create-GroupBox -Name "VR_P1_GroupBox1" -Location "30,50" -Size "450,130" -Text $SettingsHash.Guild.Ranks.Rank7 -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VR_Panel1
               $VR_P1_GB1_DataGridView1 = Create-DataGridView -Name "VR_P1_GB1_DataGridView1" -Location "-4,20" -Size "454,108" -BackgroundColor "#FFF0F0F0" -ColumnCount 2 -AllowUserToResizeColumns 0 -ColumnHeadersHeight 30 -ReadOnly -Form $VR_P1_GroupBox1 -Hash $FormHash
                  $VR_P1_GB1_DataGridView1.RowHeadersWidth = 4
                  $VR_P1_GB1_DataGridView1.RowsDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft San Serif",11,[System.Drawing.FontStyle]::Regular)
                  $VR_P1_GB1_DataGridView1.Columns[0].Name = "Account"
                  $VR_P1_GB1_DataGridView1.Columns[0].Width = 228
                  $VR_P1_GB1_DataGridView1.Columns[1].Name = "Last Logon"
                  $VR_P1_GB1_DataGridView1.Columns[1].Width = 203
                  $VR_P1_GB1_DataGridView1.Columns[1].DefaultCellStyle.Format = 'MM/dd/yyyy hh:mm:ss tt'
            $VR_P1_GroupBox2 = Create-GroupBox -Name "VR_P1_GroupBox2" -Location "$($VR_P1_GroupBox1.Left),$($VR_P1_GroupBox1.Bottom+10)" -Size "450,150" -Text $SettingsHash.Guild.Ranks.Rank6 -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VR_Panel1
               $VR_P1_GB2_DataGridView1 = Create-DataGridView -Name "VR_P1_GB2_DataGridView1" -Location "-4,20" -Size "454,128" -BackgroundColor "#FFF0F0F0" -ColumnCount 2 -AllowUserToResizeColumns 0 -ColumnHeadersHeight 30 -ReadOnly -Form $VR_P1_GroupBox2 -Hash $FormHash
                  $VR_P1_GB2_DataGridView1.RowHeadersWidth = 4
                  $VR_P1_GB2_DataGridView1.RowsDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft San Serif",11,[System.Drawing.FontStyle]::Regular)
                  $VR_P1_GB2_DataGridView1.Columns[0].Name = "Account"
                  $VR_P1_GB2_DataGridView1.Columns[0].Width = 228
                  $VR_P1_GB2_DataGridView1.Columns[1].Name = "Last Logon"
                  $VR_P1_GB2_DataGridView1.Columns[1].Width = 203
                  $VR_P1_GB2_DataGridView1.Columns[1].DefaultCellStyle.Format = 'MM/dd/yyyy hh:mm:ss tt'
            $VR_P1_GroupBox3 = Create-GroupBox -Name "VR_P1_GroupBox3" -Location "$($VR_P1_GroupBox1.Left),$($VR_P1_GroupBox2.Bottom+10)" -Size "450,180" -Text $SettingsHash.Guild.Ranks.Rank5 -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VR_Panel1
               $VR_P1_GB3_DataGridView1 = Create-DataGridView -Name "VR_P1_GB3_DataGridView1" -Location "-4,20" -Size "454,158" -BackgroundColor "#FFF0F0F0" -ColumnCount 2 -AllowUserToResizeColumns 0 -ColumnHeadersHeight 30 -ReadOnly -Form $VR_P1_GroupBox3 -Hash $FormHash
                  $VR_P1_GB3_DataGridView1.RowHeadersWidth = 4
                  $VR_P1_GB3_DataGridView1.RowsDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft San Serif",11,[System.Drawing.FontStyle]::Regular)
                  $VR_P1_GB3_DataGridView1.Columns[0].Name = "Account"
                  $VR_P1_GB3_DataGridView1.Columns[0].Width = 228
                  $VR_P1_GB3_DataGridView1.Columns[1].Name = "Last Logon"
                  $VR_P1_GB3_DataGridView1.Columns[1].Width = 203
                  $VR_P1_GB3_DataGridView1.Columns[1].DefaultCellStyle.Format = 'MM/dd/yyyy hh:mm:ss tt'
            $VR_P1_GroupBox4 = Create-GroupBox -Name "VR_P1_GroupBox4" -Location "$($VR_P1_GroupBox1.Left),$($VR_P1_GroupBox3.Bottom+10)" -Size "450,200" -Text $SettingsHash.Guild.Ranks.Rank4 -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VR_Panel1
               $VR_P1_GB4_DataGridView1 = Create-DataGridView -Name "VR_P1_GB4_DataGridView1" -Location "-4,20" -Size "454,178" -BackgroundColor "#FFF0F0F0" -ColumnCount 2 -AllowUserToResizeColumns 0 -ColumnHeadersHeight 30 -ReadOnly -Form $VR_P1_GroupBox4 -Hash $FormHash
                  $VR_P1_GB4_DataGridView1.RowHeadersWidth = 4
                  $VR_P1_GB4_DataGridView1.RowsDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft San Serif",11,[System.Drawing.FontStyle]::Regular)
                  $VR_P1_GB4_DataGridView1.Columns[0].Name = "Account"
                  $VR_P1_GB4_DataGridView1.Columns[0].Width = 228
                  $VR_P1_GB4_DataGridView1.Columns[1].Name = "Last Logon"
                  $VR_P1_GB4_DataGridView1.Columns[1].Width = 203
                  $VR_P1_GB4_DataGridView1.Columns[1].DefaultCellStyle.Format = 'MM/dd/yyyy hh:mm:ss tt'
            $VR_P1_GroupBox5 = Create-GroupBox -Name "VR_P1_GroupBox5" -Location "$($VR_P1_GroupBox1.Right+25),$($VR_P1_GroupBox1.Top)"   -Size "450,260" -Text $SettingsHash.Guild.Ranks.Rank3 -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VR_Panel1
               $VR_P1_GB5_DataGridView1 = Create-DataGridView -Name "VR_P1_GB5_DataGridView1" -Location "-4,20" -Size "454,238" -BackgroundColor "#FFF0F0F0" -ColumnCount 2 -AllowUserToResizeColumns 0 -ColumnHeadersHeight 30 -ReadOnly -Form $VR_P1_GroupBox5 -Hash $FormHash
                  $VR_P1_GB5_DataGridView1.RowHeadersWidth = 4
                  $VR_P1_GB5_DataGridView1.RowsDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft San Serif",11,[System.Drawing.FontStyle]::Regular)
                  $VR_P1_GB5_DataGridView1.Columns[0].Name = "Account"
                  $VR_P1_GB5_DataGridView1.Columns[0].Width = 228
                  $VR_P1_GB5_DataGridView1.Columns[1].Name = "Last Logon"
                  $VR_P1_GB5_DataGridView1.Columns[1].Width = 203
                  $VR_P1_GB5_DataGridView1.Columns[1].ValueType = "DateTime"
                  $VR_P1_GB5_DataGridView1.Columns[1].DefaultCellStyle.Format = 'MM/dd/yyyy hh:mm:ss tt'
            $VR_P1_GroupBox6 = Create-GroupBox -Name "VR_P1_GroupBox6" -Location "$($VR_P1_GroupBox5.Left),$($VR_P1_GroupBox5.Bottom+10)" -Size "450,220" -Text $SettingsHash.Guild.Ranks.Rank2 -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VR_Panel1
               $VR_P1_GB6_DataGridView1 = Create-DataGridView -Name "VR_P1_GB6_DataGridView1" -Location "-4,20" -Size "454,198" -BackgroundColor "#FFF0F0F0" -ColumnCount 2 -AllowUserToResizeColumns 0 -ColumnHeadersHeight 30 -ReadOnly -Form $VR_P1_GroupBox6 -Hash $FormHash
                  $VR_P1_GB6_DataGridView1.RowHeadersWidth = 4
                  $VR_P1_GB6_DataGridView1.RowsDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft San Serif",11,[System.Drawing.FontStyle]::Regular)
                  $VR_P1_GB6_DataGridView1.Columns[0].Name = "Account"
                  $VR_P1_GB6_DataGridView1.Columns[0].Width = 228
                  $VR_P1_GB6_DataGridView1.Columns[1].Name = "Last Logon"
                  $VR_P1_GB6_DataGridView1.Columns[1].Width = 203
                  $VR_P1_GB6_DataGridView1.Columns[1].DefaultCellStyle.Format = 'MM/dd/yyyy hh:mm:ss tt'
            $VR_P1_GroupBox7 = Create-GroupBox -Name "VR_P1_GroupBox7" -Location "$($VR_P1_GroupBox5.Left),$($VR_P1_GroupBox6.Bottom+10)" -Size "450,190" -Text $SettingsHash.Guild.Ranks.Rank1 -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VR_Panel1
               $VR_P1_GB7_DataGridView1 = Create-DataGridView -Name "VR_P1_GB7_DataGridView1" -Location "-4,20" -Size "454,168" -BackgroundColor "#FFF0F0F0" -ColumnCount 2 -AllowUserToResizeColumns 0 -ColumnHeadersHeight 30 -ReadOnly -Form $VR_P1_GroupBox7 -Hash $FormHash
                  $VR_P1_GB7_DataGridView1.RowHeadersWidth = 4
                  $VR_P1_GB7_DataGridView1.RowsDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft San Serif",11,[System.Drawing.FontStyle]::Regular)
                  $VR_P1_GB7_DataGridView1.Columns[0].Name = "Account"
                  $VR_P1_GB7_DataGridView1.Columns[0].Width = 228
                  $VR_P1_GB7_DataGridView1.Columns[1].Name = "Last Logon"
                  $VR_P1_GB7_DataGridView1.Columns[1].Width = 203
                  $VR_P1_GB7_DataGridView1.Columns[1].DefaultCellStyle.Format = 'MM/dd/yyyy hh:mm:ss tt'
      $ViewInactiveAccounts = Create-ToolStripMenuItem -Name "InactiveAccounts" -Text "Inactivity" -Form $ViewMenu -AutoSize -DropDownItem
         $VIA_Panel1 = Create-PanelBox -Name "VIA_Panel1" -Size "1000,780" -Location "0,0" -Visible 0 -Form $Form -BackColor TRANSPARENT
            $VIA_P1_GroupBox1 = Create-GroupBox -Name "VIA_P1_GroupBox1" -Location "30,40" -Size "930,700" -Text "Inactive Characters:" -FontSize 12 -AutoSize 0 -BackColor TRANSPARENT -Form $VIA_Panel1
               $VIA_P1_GB1_DataGridView1 = Create-DataGridView -Name "VIA_P1_GB1_DataGridView1" -Location "-4,20" -Size "934,678" -BackgroundColor "#FFF0F0F0" -ColumnCount 5 -AllowUserToOrderColumns 1 -ColumnHeadersHeight 30 -ReadOnly -Form $VIA_P1_GroupBox1
                  $VIA_P1_GB1_DataGridView1.RowHeadersWidth = 4
                  $VIA_P1_GB1_DataGridView1.RowsDefaultCellStyle.Font = New-Object System.Drawing.Font("Microsoft San Serif",11,[System.Drawing.FontStyle]::Regular)
                  $VIA_P1_GB1_DataGridView1.Columns[0].Name = "Name"
                  $VIA_P1_GB1_DataGridView1.Columns[0].Width = 160
                  $VIA_P1_GB1_DataGridView1.Columns[1].Name = "Account"
                  $VIA_P1_GB1_DataGridView1.Columns[1].Width = 158
                  $VIA_P1_GB1_DataGridView1.Columns[2].Name = "Rank"
                  $VIA_P1_GB1_DataGridView1.Columns[2].Width = 200
                  $VIA_P1_GB1_DataGridView1.Columns[3].Name = "Last Logon"
                  $VIA_P1_GB1_DataGridView1.Columns[3].Width = 177
                  $VIA_P1_GB1_DataGridView1.Columns[3].DefaultCellStyle.Format = 'MM/dd/yyyy hh:mm:ss tt'
                  $VIA_P1_GB1_DataGridView1.Columns[4].Name = "Officer's Comment"
                  $VIA_P1_GB1_DataGridView1.Columns[4].Width = 216
   $HelpMenu = Create-ToolStripMenuItem -Name "HelpMenu" -Text "&Help" -Form $Main -AutoSize -MenuItem -BackColor TRANSPARENT -Available 0
      $HelpAbout = Create-ToolStripMenuItem -Name "HelpAbout" -Text "About" -Form $HelpMenu -AutoSize -DropDownItem
      $HelpManual = Create-ToolStripMenuItem -Name "HelpManual" -Text "Manual" -Form $HelpMenu -AutoSize -DropDownItem
    
    
    
      
#####-- Startup
$Startup = {
   If ($VariableHash.DataImported -eq $True) {
      ForEach ( $Account in $DataHash.Keys ) {
         For ( $i = 7 ; $i -gt 0 ; $i-- ) {
            If ($DataHash.($Account).Rank -like $SettingsHash.Guild.Ranks.("Rank$i")  -AND $DataHash.($Account).Status -like "Member") {
               $FormHash.("VR_P1_GB$(8-$i)_DataGridView1").Rows.Add($Account,[System.DateTime]$DataHash.($Account).LastLogon)
            }
         }        
      }
      Inactivity-Update
   }
   Else {
      $ViewMenu.Available = $False
      $ToolMenu.Available = $False
   }
   If ($VariableHash.ClassMetricsImported -eq $True) {
   }
}


#####-- Functions
.(Join-Path $ScriptPath ..\SupportFiles\Functions.ps1)

Function ViewTools-Toggle ([Int32]$Panel) {
   Switch ($Panel) {
      1 {$VIA_Panel1.Visible = $True }
      2 {$VIA_Panel1.Visible = $False}
   }
}

#####-- Event scriptblocks
.(Join-Path $ScriptPath ..\SupportFiles\EventScripts.ps1)

#####-- Events
$Form.Add_Load($Startup)
$Form.Add_Closing({
   If ($SettingsHash.SaveOnExit -eq $True) {
      Export-Clixml -Path "$($SettingsHash.DataLocation)\Data.xml" -InputObject $DataHash
   }
   #PS -PID $PID | KILL
})

#- FileMenu
$FileImportCSV[0].Add_Click($FileImportCSV_Click)
$FileSaveData[0].Add_Click($FileSaveData_Click)
$FileExit[0].Add_Click({$Form.Close()})

#- ConfigurationSettings
$ConfigurationSettings[0].Add_Click($ConfigurationSettings_Click)
$ConfigurationGuild[0].Add_Click($ConfigurationGuild_Click)

$CS_P1_GB1_Button1.Add_Click($CS_P1_GB1_Button1_Click)
$CS_P1_GB2_Button1.Add_Click($CS_P1_GB2_Button1_Click)
$CS_P1_GB3_Button1.Add_Click($CS_P1_GB3_Button1_Click)
$CS_P1_CheckBox1.Add_CheckedChanged($CS_P1_CheckBox1_CheckedChanged)
$CS_P1_Button1.Add_Click($CS_P1_Button1_Click)
$CG_P1_Button1.Add_Click($CG_P1_Button1_Click)

#- View
$ViewClasses[0].Add_Click($ViewClasses_Click)

$VC_P1_GB3_ComboBox1.Add_SelectedIndexChanged($VC_P1_GB3_ComboBox1_SelectedIndexChanged)
$VC_P1_GB3_ListBox1.Add_SelectedValueChanged($VC_P1_GB3_ListBox1_SelectedValueChanged)
$VC_P1_GB3_ListBox2.Add_SelectedValueChanged($VC_P1_GB3_ListBox2_SelectedValueChanged)
$VC_P1_Button1.Add_Click($VC_P1_Button1_Click)

$ViewAccounts[0].Add_Click($ViewAccounts_Click)
$VA_P1_GB1_CheckBox1.Add_CheckedChanged($VA_P1_GB1_CheckBox1_CheckedChanged)
$VA_P1_GB1_ListBox1.Add_SelectedIndexChanged($VA_P1_GB1_ListBox1_SelectedIndexChanged)
$VA_P1_GB1_GB1_RichTextBox1.Add_Leave($VA_P1_GB1_GB1_RichTextBox1_Leave)
$VA_P1_GB2_ListBox1.Add_SelectedIndexChanged($VA_P1_GB2_ListBox1_SelectedIndexChanged)
$VA_P1_GB2_GB1_RichTextBox1.Add_Leave($VA_P1_GB2_GB1_RichTextBox1_Leave)
$VA_P1_GB2_CheckBox1.Add_CheckedChanged($VA_P1_GB2_CheckBox1_CheckedChanged)

$ViewRanks[0].Add_Click($ViewRanks_Click)

$ViewInactiveAccounts[0].Add_Click($ViewInactiveAccounts_Click)

#####-- End
$Form.ShowDialog()