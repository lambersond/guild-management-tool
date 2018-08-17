####              ####
# File DropDown Opts #
####              ####
$FileImportCSV_Click = {
   If ($DataHash.Count -gt 0) {
      $Answer = [Microsoft.VisualBasic.Interaction]::MsgBox("Do you wish to merge the new import with the old date?",'YesNo',"Import CSV")
      If ( $Answer -like "Yes" ) { $VariableHash.MergeDataHash = $True }
      Else { $VariableHash.MergeDataHash = $False }
   }

   $Location = New-Object System.Windows.Forms.OpenFileDialog
   $Location.ShowHelp = $True
   If ($Location.ShowDialog() -like "OK") {
      $VariableHash.ImportedCSV = Import-Csv -Path $location.FileName
      $VariableHash.DataHashOld = $DataHash.Clone()
      $DataHash.Clear()

      For ( $i = 0 ; $i -lt $VariableHash.ImportedCSV.Count ; $i++) {
         If ( !($DataHash.ContainsKey("$($VariableHash.ImportedCSV[$i].'Account Handle')")) ) {
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')") = New-Object PSObject @{
               Rank = $VariableHash.ImportedCSV[$i].'Guild Rank'
               JoinDate = $VariableHash.ImportedCSV[$i].'Join Date'
               LastLogon = $VariableHash.ImportedCSV[$i].'Last Active Date'
               RankDate = $VariableHash.ImportedCSV[$i].'Rank Change Date'
               Classes = @("$($VariableHash.ImportedCSV[$i].'Class')")
               MemberClasses = @("$($VariableHash.ImportedCSV[$i].'Class')")
               Status = "Member"
               Notes = ""
            }
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters = @{}
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')") = New-Object PSObject @{
               PublicComment = $VariableHash.ImportedCSV[$i].'Public Comment'
               PublicCommentLastEdit = $VariableHash.ImportedCSV[$i].'Public Comment Last Edit Date'
               OfficerComment = $VariableHash.ImportedCSV[$i].'Officer Comment'
               OfficerCommentAuthor = $VariableHash.ImportedCSV[$i].'Officer Comment Author'
               OfficerCommentLastEdit = $VariableHash.ImportedCSV[$i].'Officer Comment Last Edit Date'
               Level = $VariableHash.ImportedCSV[$i].'Level'
               LastLogon = $VariableHash.ImportedCSV[$i].'Last Active Date'
               Class = $VariableHash.ImportedCSV[$i].'Class'
               CharacterRank = $VariableHash.ImportedCSV[$i].'Guild Rank'
               CharacterRankDate = $VariableHash.ImportedCSV[$i].'Rank Change Date'
               TotalCharacters = 1
               Status = "Member"
               Notes = ""
            }
         }
         ElseIf ( !($DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.ContainsKey("$($VariableHash.ImportedCSV[$i].'Character Name')")) ) {
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')") = New-Object PSObject @{
               PublicComment = $VariableHash.ImportedCSV[$i].'Public Comment'
               PublicCommentLastEdit = $VariableHash.ImportedCSV[$i].'Public Comment Last Edit Date'
               OfficerComment = $VariableHash.ImportedCSV[$i].'Officer Comment'
               OfficerCommentAuthor = $VariableHash.ImportedCSV[$i].'Officer Comment Author'
               OfficerCommentLastEdit = $VariableHash.ImportedCSV[$i].'Officer Comment Last Edit Date'
               Level = $VariableHash.ImportedCSV[$i].'Level'
               LastLogon = $VariableHash.ImportedCSV[$i].'Last Active Date'
               Class = $VariableHash.ImportedCSV[$i].'Class'
               CharacterRank = $VariableHash.ImportedCSV[$i].'Guild Rank'
               CharacterRankDate = $VariableHash.ImportedCSV[$i].'Rank Change Date'
               Status = "Member"
               Notes = ""
            }
            
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").TotalCharacters += 1
            If ($DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Classes -notcontains $VariableHash.ImportedCSV[$i].'Class' ) {
               $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Classes += $VariableHash.ImportedCSV[$i].'Class'
            }
            
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").MemberClasses += $VariableHash.ImportedCSV[$i].'Class'
            
            $GuildRank = 0
            $CharacterRank = 0

            Switch ($VariableHash.ImportedCSV[$i].'Guild Rank') {
               "$($SettingsHash.Guild.Ranks.Rank7)" { $CharacterRank = 7 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank6)" { $CharacterRank = 6 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank5)" { $CharacterRank = 5 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank4)" { $CharacterRank = 4 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank3)" { $CharacterRank = 3 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank2)" { $CharacterRank = 2 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank1)" { $CharacterRank = 1 ; Break }
            }
            Switch ($DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Rank) {
               "$($SettingsHash.Guild.Ranks.Rank7)" { $GuildRank = 7 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank6)" { $GuildRank = 6 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank5)" { $GuildRank = 5 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank4)" { $GuildRank = 4 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank3)" { $GuildRank = 3 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank2)" { $GuildRank = 2 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank1)" { $GuildRank = 1 ; Break }
            }
            If ($CharacterRank -gt $GuildRank) { $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Rank = $VariableHash.ImportedCSV[$i].'Guild Rank' }
            
            If ((New-TimeSpan -Start $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").LastLogon -End $VariableHash.ImportedCSV[$i].'Last Active Date').Seconds -gt 0) {
               $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").LastLogon = $VariableHash.ImportedCSV[$i].'Last Active Date'
            }

            If ((New-TimeSpan -Start $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").JoinDate -End $VariableHash.ImportedCSV[$i].'Join Date').Seconds -lt 0) {
               $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").JoinDate = $VariableHash.ImportedCSV[$i].'Join Date'
            }
         }
         Else {
            $GuildRank = 0
            $CharacterRank = 0

            Switch ($VariableHash.ImportedCSV[$i].'Guild Rank') {
               "$($SettingsHash.Guild.Ranks.Rank7)" { $CharacterRank = 7 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank6)" { $CharacterRank = 6 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank5)" { $CharacterRank = 5 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank4)" { $CharacterRank = 4 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank3)" { $CharacterRank = 3 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank2)" { $CharacterRank = 2 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank1)" { $CharacterRank = 1 ; Break }
            }
            Switch ($DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Rank) {
               "$($SettingsHash.Guild.Ranks.Rank7)" { $GuildRank = 7 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank6)" { $GuildRank = 6 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank5)" { $GuildRank = 5 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank4)" { $GuildRank = 4 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank3)" { $GuildRank = 3 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank2)" { $GuildRank = 2 ; Break }
               "$($SettingsHash.Guild.Ranks.Rank1)" { $GuildRank = 1 ; Break }
            }
            If ($CharacterRank -gt $GuildRank) { $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Rank = $VariableHash.ImportedCSV[$i].'Guild Rank' }
            
            If ((New-TimeSpan -Start $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").LastLogon -End $VariableHash.ImportedCSV[$i].'Last Active Date').Seconds -gt 0) {
               $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").LastLogon = $VariableHash.ImportedCSV[$i].'Last Active Date'
            }
            
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')").PublicComment = $VariableHash.ImportedCSV[$i].'Public Comment'
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')").PublicCommentLastEdit = $VariableHash.ImportedCSV[$i].'Public Comment Last Edit Date'
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')").OfficerComment = $VariableHash.ImportedCSV[$i].'Officer Comment'
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')").OfficerCommentAuthor = $VariableHash.ImportedCSV[$i].'Officer Comment Author'
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')").OfficerCommentLastEdit = $VariableHash.ImportedCSV[$i].'Officer Comment Last Edit Date'
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')").Level = $VariableHash.ImportedCSV[$i].'Level'
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')").CharacterRank = $VariableHash.ImportedCSV[$i].'Guild Rank'
            $DataHash.("$($VariableHash.ImportedCSV[$i].'Account Handle')").Characters.("$($VariableHash.ImportedCSV[$i].'Character Name')").CharacterRankDate = $VariableHash.ImportedCSV[$i].'Rank Change Date'
         }
      }

      If ($VariableHash.MergeDataHash -eq $True) {
         ForEach ($Account in $VariableHash.DataHashOld.Keys){
            If ( !($DataHash.ContainsKey($Account)) ) {
               $DataHash.($Account) = $VariableHash.DataHashOld.($Account)
               $DataHash.($Account).Status = "Non Member"
               $DataHash.($Account).MemberClasses = @()
               
               ForEach ($Character in $DataHash.($Account).Characters.Keys ) {
                  $DataHash.($Account).Characters.($Character).Status = "Non Member"
               }
            }
            Else {
               ForEach ($Character in $VariableHash.DataHashOld.($Account).Characters.Keys ) {
                  If ( !($DataHash.($Account).Characters.ContainsKey($Character)) ) {
                     $DataHash.($Account).Characters.($Character) = $VariableHash.DataHashOld.($Account).Characters.($Character)
                     $DataHash.($Account).Characters.($Character).Status = "Non Member"
                     $temp = [array]::indexof([Array]$DataHash.($Account).memberclasses,$VariableHash.DataHashOld.($Account).Characters.($Character).Class)
                     $DataHash.($Account).MemberClasses[$temp] = $Null
                  }
               }
            }
         }
      }
      $VariableHash.DataHashOld = $Null
   }
   $ViewMenu.Available = $True
   $ToolMenu.Available = $True
}
$FileSaveData_Click = {
   Export-Clixml -Path "$($SettingsHash.DataLocation)\Data.xml" -InputObject $DataHash
}

####              ####
# Configure Settings #
####              ####
$ConfigurationSettings_Click = {
   ViewVisuals-Toggle 5
   ConfigurationVisuals-Toggle 1
}
$CS_P1_GB1_Button1_Click = {
   $Location = New-Object System.Windows.Forms.FolderBrowserDialog
   $Location.SelectedPath = $SettingsHash.ImportLocation
   $Location.ShowHelp = $True
   If ( $Location.ShowDialog() -like "OK" ) { $CS_P1_GB1_RichTextBox1.Text = $location.SelectedPath }
}
$CS_P1_GB2_Button1_Click = {
   $Location = New-Object System.Windows.Forms.FolderBrowserDialog
   $Location.SelectedPath = $SettingsHash.SettingsLocation
   $Location.ShowHelp = $True
   If ( $Location.ShowDialog() -like "OK" ) { $CS_P1_GB2_RichTextBox1.Text = $location.SelectedPath }
}
$CS_P1_GB3_Button1_Click = {
   $Location = New-Object System.Windows.Forms.FolderBrowserDialog
   $Location.SelectedPath = $SettingsHash.DataLocation
   $Location.ShowHelp = $True
   If ( $Location.ShowDialog() -like "OK" ) { $CS_P1_GB3_RichTextBox1.Text = $location.SelectedPath }
}
$CS_P1_Button1_Click = {
   $SettingsHash.ImportLocation   = $CS_P1_GB1_RichTextBox1.Text
   $SettingsHash.SettingsLocation = $CS_P1_GB2_RichTextBox1.Text
   $SettingsHash.DataLocation     = $CS_P1_GB3_RichTextBox1.Text
   $SettingsHash.UserName         = $CS_P1_GB4_RichTextBox1.Text

   Set-ItemProperty -Path HKCU:\SOFTWARE\DnD -Name Import -Value $SettingsHash.ImportLocation -Force   | Out-Null
   Set-ItemProperty -Path HKCU:\SOFTWARE\DnD -Name Settings -Value $SettingsHash.SettingsLocation -Force | Out-Null
   Set-ItemProperty -Path HKCU:\SOFTWARE\DnD -Name Data -Value $SettingsHash.DataLocation -Force     | Out-Null
   Set-ItemProperty -Path HKCU:\SOFTWARE\DnD -Name UserName -Value $SettingsHash.UserName -Force  | Out-Null
}
$CS_P1_CheckBox1_CheckedChanged = {
   If($This.Checked){$SettingsHash.SaveOnExit = $True}
   Else {$SettingsHash.SaveOnExit = $False}
}

####           ####
# Configure Guild #
####           ####
$ConfigurationGuild_Click = {
   ViewVisuals-Toggle 5
   ConfigurationVisuals-Toggle 2
}
$CG_P1_Button1_Click = {
   $SettingsHash.Guild.Ranks.Rank1 = $CG_P1_GB1_RichTextBox1.Text
   $SettingsHash.Guild.Ranks.Rank2 = $CG_P1_GB1_RichTextBox2.Text
   $SettingsHash.Guild.Ranks.Rank3 = $CG_P1_GB1_RichTextBox3.Text
   $SettingsHash.Guild.Ranks.Rank4 = $CG_P1_GB1_RichTextBox4.Text
   $SettingsHash.Guild.Ranks.Rank5 = $CG_P1_GB1_RichTextBox5.Text
   $SettingsHash.Guild.Ranks.Rank6 = $CG_P1_GB1_RichTextBox6.Text
   $SettingsHash.Guild.Ranks.Rank7 = $CG_P1_GB1_RichTextBox7.Text
   
   $SettingsHash.Guild.Inactivity.Rank1 = $CG_P1_GB2_RichTextBox1.Text
   $SettingsHash.Guild.Inactivity.Rank2 = $CG_P1_GB2_RichTextBox2.Text
   $SettingsHash.Guild.Inactivity.Rank3 = $CG_P1_GB2_RichTextBox3.Text
   $SettingsHash.Guild.Inactivity.Rank4 = $CG_P1_GB2_RichTextBox4.Text
   $SettingsHash.Guild.Inactivity.Rank5 = $CG_P1_GB2_RichTextBox5.Text
   $SettingsHash.Guild.Inactivity.Rank6 = $CG_P1_GB2_RichTextBox6.Text
   $SettingsHash.Guild.Inactivity.Rank7 = $CG_P1_GB2_RichTextBox7.Text

   Export-Clixml -Path "$($SettingsHash.SettingsLocation)\GuildSettings.xml" -InputObject $SettingsHash.Guild
}

####        ####
# View Classes #
####        ####
$ViewClasses_Click = {
   ConfigurationVisuals-Toggle 3
   ViewVisuals-Toggle 1
}
$VC_P1_GB3_ComboBox1_SelectedIndexChanged = {
   $VC_P1_GB3_ListBox1.Items.Clear()
   ForEach ($Key in $DataHash.Keys) { If ($DataHash.($Key).Classes -contains $This.Text -AND $DataHash.($Key).Status -like "Member") {$VC_P1_GB3_ListBox1.Items.Add($Key)} }
}
$VC_P1_GB3_ListBox1_SelectedValueChanged = {
   $VC_P1_GB3_ListBox2.Items.Clear()
   ForEach ($Character in $DataHash.($This.Text).Characters.Keys) {
      If ($DataHash.($This.Text).Characters.($Character).Class -contains $VC_P1_GB3_ComboBox1.Text -AND $DataHash.($This.Text).Characters.($Character).Status -like "Member") {$VC_P1_GB3_ListBox2.Items.Add($Character)}
   }
}
$VC_P1_GB3_ListBox2_SelectedValueChanged = {
   $VC_P1_GB3_GB1_Label01.Text = $DataHash.($VC_P1_GB3_ListBox1.Text).Characters.($VC_P1_GB3_ListBox2.Text).CharacterRank
   $VC_P1_GB3_GB1_Label02.Text = $DataHash.($VC_P1_GB3_ListBox1.Text).Characters.($VC_P1_GB3_ListBox2.Text).Level
   $VC_P1_GB3_GB1_Label03.Text = $DataHash.($VC_P1_GB3_ListBox1.Text).Characters.($VC_P1_GB3_ListBox2.Text).PublicComment
   $VC_P1_GB3_GB1_Label4.Location  = "$($VC_P1_GB3_GB1_Label3.Left),$($VC_P1_GB3_GB1_Label03.Bottom)"
   $VC_P1_GB3_GB1_Label04.Location = "$($VC_P1_GB3_GB1_Label4.Left+25),$($VC_P1_GB3_GB1_Label4.Bottom)"
   $VC_P1_GB3_GB1_Label04.Text = $DataHash.($VC_P1_GB3_ListBox1.Text).Characters.($VC_P1_GB3_ListBox2.Text).OfficerComment
   $VC_P1_GB3_GB1_Label5.Location  = "$($VC_P1_GB3_GB1_Label4.Left),$($VC_P1_GB3_GB1_Label04.Bottom)"
   $VC_P1_GB3_GB1_Label05.Location = "$($VC_P1_GB3_GB1_Label5.Left+25),$($VC_P1_GB3_GB1_Label5.Bottom)"
   $VC_P1_GB3_GB1_Label05.Text = $DataHash.($VC_P1_GB3_ListBox1.Text).Characters.($VC_P1_GB3_ListBox2.Text).Notes
}
$VC_P1_Button1_Click = {
   ViewClasses-Update
}

####         ####
# View Accounts #
####         ####
$ViewAccounts_Click = {
   ConfigurationVisuals-Toggle 3
   ViewVisuals-Toggle 2
}
$VA_P1_GB1_CheckBox1_CheckedChanged = {
   $VA_P1_GB1_ListBox1.Items.Clear()

   If ($This.Checked) {
      ForEach ( $Account in $DataHash.Keys ) {
         If ($DataHash.($Account).Status -like "Member") { $VA_P1_GB1_ListBox1.Items.Add($Account) }
      }
   }
   Else {
      ForEach ( $Account in $DataHash.Keys ) {
         $VA_P1_GB1_ListBox1.Items.Add($Account)
      }
   }
}
$VA_P1_GB1_ListBox1_SelectedIndexChanged = {
   $VA_P1_GB1_GB1_Label01.Text = $DataHash.($This.Text).Rank
   $VA_P1_GB1_GB1_Label02.Text = $DataHash.($This.Text).JoinDate
   $VA_P1_GB1_GB1_Label03.Text = $DataHash.($This.Text).LastLogon
   $VA_P1_GB1_GB1_Label04.Text = $DataHash.($This.Text).RankDate
   $VA_P1_GB1_GB1_Label06.Text = $DataHash.($This.Text).Status
   If ($DataHash.($This.Text).Status -like "Member") {$VA_P1_GB1_GB1_Label06.ForeColor = "FORESTGREEN"}
   Else {$VA_P1_GB1_GB1_Label06.ForeColor = "RED"}
   $VA_P1_GB1_GB1_RichTextBox1.Text = $DataHash.($This.Text).Notes

   $VA_P1_GB2_ListBox1.Items.Clear()
   ForEach ($Character in $DataHash.($VA_P1_GB1_ListBox1.Text).Characters.Keys) {
      $VA_P1_GB2_ListBox1.Items.Add($Character)
   }
}
$VA_P1_GB1_GB1_RichTextBox1_Leave = {
   $DataHash.($VA_P1_GB1_ListBox1.Text).Notes = $VA_P1_GB1_GB1_RichTextBox1.Text
}
$VA_P1_GB2_ListBox1_SelectedIndexChanged = {
   $VariableHash.ViewAccount_CharacterAccount = $VA_P1_GB1_ListBox1.Text

   $VA_P1_GB2_GB1_Label01.Text = $DataHash.($VA_P1_GB1_ListBox1.Text).Rank
   $VA_P1_GB2_GB1_Label02.Text = "Level $($DataHash.($VA_P1_GB1_ListBox1.Text).Characters.($This.Text).Level), $($DataHash.($VA_P1_GB1_ListBox1.Text).Characters.($This.Text).Class)"
   $VA_P1_GB2_GB1_Label03.Text = $DataHash.($VA_P1_GB1_ListBox1.Text).Characters.($This.Text).LastLogon
   $VA_P1_GB2_GB1_Label04.Text = $DataHash.($VA_P1_GB1_ListBox1.Text).Characters.($This.Text).PublicComment
   $VA_P1_GB2_GB1_Label05.Text = $DataHash.($VA_P1_GB1_ListBox1.Text).Characters.($This.Text).OfficerComment
   $VA_P1_GB2_GB1_Label07.Text = $DataHash.($VA_P1_GB1_ListBox1.Text).Characters.($This.Text).Status
   If ($DataHash.($VA_P1_GB1_ListBox1.Text).Characters.($This.Text).Status -like "Member") {$VA_P1_GB2_GB1_Label07.ForeColor = "FORESTGREEN"}
   Else {$VA_P1_GB2_GB1_Label07.ForeColor = "RED"}

   $VA_P1_GB2_GB1_RichTextBox1.Text = $DataHash.($VA_P1_GB1_ListBox1.Text).Characters.($This.Text).Notes
}
$VA_P1_GB2_GB1_RichTextBox1_Leave = {
   $DataHash.($VariableHash.ViewAccount_CharacterAccount).Characters.($VA_P1_GB2_ListBox1.Text).Notes = $VA_P1_GB2_GB1_RichTextBox1.Text
}
$VA_P1_GB2_CheckBox1_CheckedChanged = {
   $VA_P1_GB2_ListBox1.Items.Clear()

   If ($This.Checked) {
      ForEach ( $Character in $DataHash.($VA_P1_GB1_ListBox1.Text).Characters.Keys ) {
         If ($DataHash.($VA_P1_GB1_ListBox1.Text).Characters.($Character).Status -like "Member") { $VA_P1_GB2_ListBox1.Items.Add($Character) }
      }
   }
   Else {
      ForEach ( $Character in $DataHash.($VA_P1_GB1_ListBox1.Text).Characters.Keys ) {
         $VA_P1_GB2_ListBox1.Items.Add($Character)
      }
   }
}

####      ####
# View Ranks #
####      ####
$ViewRanks_Click = {
   ConfigurationVisuals-Toggle 3
   ViewVisuals-Toggle 3
}

####           ####
# View Inactivity #
####           ####
$ViewInactiveAccounts_Click = {
   ConfigurationVisuals-Toggle 3
   ViewVisuals-Toggle 4
}
