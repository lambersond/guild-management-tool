####     ####
# Functions #
####     ####
Function ConfigurationVisuals-Toggle ([Int32]$Panel) {
   Switch ($Panel) {
      1 {$CS_Panel1.Visible = $True  ; $CG_Panel1.Visible = $False}
      2 {$CS_Panel1.Visible = $False ; $CG_Panel1.Visible = $True}
      3 {$CS_Panel1.Visible = $False ; $CG_Panel1.Visible = $False}
   }
}
Function ViewVisuals-Toggle ([Int32]$Panel) {
   Switch ($Panel) {
      1 {$VA_Panel1.Visible = $False ; $VR_Panel1.Visible = $False ; $VIA_Panel1.Visible = $False ; $VC_Panel1.Visible = $True }
      2 {
         $VC_Panel1.Visible = $False
         $VR_Panel1.Visible = $False
         $VIA_Panel1.Visible = $False
         $VA_Panel1.Visible = $True
         
         $VA_P1_GB1_ListBox1.Items.Clear()
         If ($VA_P1_GB1_CheckBox1.Checked) {
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
      3 {$VC_Panel1.Visible = $False ; $VA_Panel1.Visible = $False ; $VIA_Panel1.Visible = $False ; $VR_Panel1.Visible = $True }
      4 {$VC_Panel1.Visible = $False ; $VA_Panel1.Visible = $False ; $VR_Panel1.Visible = $False ; $VIA_Panel1.Visible = $True  }
      5 {$VC_Panel1.Visible = $False ; $VA_Panel1.Visible = $False ; $VIA_Panel1.Visible = $False ; $VR_Panel1.Visible = $False }
   }
}
Function Inactivity-Update {
   ForEach ($Account in $DataHash.Keys) {
      If ($DataHash.($Account).Status -like "Member") {
         ForEach ($Character in $DataHash.($Account).Characters.Keys) {
            If ($DataHash.($Account).Characters.($Character).Status -like "Member") {
               $InactiveTime = -1
               For ( $i = 7 ; $i -gt 0 ; $i-- ) {
                  If ($DataHash.($Account).Rank -like $SettingsHash.Guild.Ranks.("Rank$i")) { $InactiveTime = $SettingsHash.Guild.Inactivity.("Rank$i") ; $i = 0 }
               }
               If ( (New-TimeSpan -Start $DataHash.($Account).Characters.($Character).LastLogon -End ([System.DateTime]::Now)).Days -ge $InactiveTime ) {
                  $VIA_P1_GB1_DataGridView1.rows.Add($Character,$Account,$DataHash.($Account).Characters.($Character).CharacterRank,[System.DateTime]$DataHash.($Account).Characters.($Character).LastLogon,$DataHash.($Account).Characters.($Character).OfficerComment)
               }
            }
         }
      }
   }
}
Function ViewClasses-Update-OLD { #OLD- backup
   $MetricsHash.Classes.Totals.CW  = 0
   $MetricsHash.Classes.Totals.DC  = 0
   $MetricsHash.Classes.Totals.GWF = 0
   $MetricsHash.Classes.Totals.GF  = 0
   $MetricsHash.Classes.Totals.HR  = 0
   $MetricsHash.Classes.Totals.OP  = 0
   $MetricsHash.Classes.Totals.SW  = 0
   $MetricsHash.Classes.Totals.TR  = 0
   
   $MetricsHash.Classes.ByAccount.CW  = 0
   $MetricsHash.Classes.ByAccount.DC  = 0
   $MetricsHash.Classes.ByAccount.GWF = 0
   $MetricsHash.Classes.ByAccount.GF  = 0
   $MetricsHash.Classes.ByAccount.HR  = 0
   $MetricsHash.Classes.ByAccount.OP  = 0
   $MetricsHash.Classes.ByAccount.SW  = 0
   $MetricsHash.Classes.ByAccount.TR  = 0
   
   ForEach ( $account in $DataHash.Keys ) {
      $DataHash.("$account").Classes | % {
         Switch ($_) {
            "Control Wizard" {$MetricsHash.Classes.ByAccount.CW++;Break}
            "Devoted Cleric" {$MetricsHash.Classes.ByAccount.DC++;Break}
            "Great Weapon Fighter" {$MetricsHash.Classes.ByAccount.GWF++;Break}
            "Guardian Fighter" {$MetricsHash.Classes.ByAccount.GF++;Break}
            "Hunter Ranger" {$MetricsHash.Classes.ByAccount.HR++;Break}
            "Oathbound Paladin" {$MetricsHash.Classes.ByAccount.OP++;Break}
            "Scourge Warlock" {$MetricsHash.Classes.ByAccount.SW++;Break}
            "Trickster Rogue" {$MetricsHash.Classes.ByAccount.TR++;Break}
         }
      }
      $DataHash.("$account").Characters.Keys | % {
         Switch ($DataHash.("$account").Characters.($_).Class) {
            "Control Wizard" {$MetricsHash.Classes.Totals.CW++;Break}
            "Devoted Cleric" {$MetricsHash.Classes.Totals.DC++;Break}
            "Great Weapon Fighter" {$MetricsHash.Classes.Totals.GWF++;Break}
            "Guardian Fighter" {$MetricsHash.Classes.Totals.GF++;Break}
            "Hunter Ranger" {$MetricsHash.Classes.Totals.HR++;Break}
            "Oathbound Paladin" {$MetricsHash.Classes.Totals.OP++;Break}
            "Scourge Warlock" {$MetricsHash.Classes.Totals.SW++;Break}
            "Trickster Rogue" {$MetricsHash.Classes.Totals.TR++;Break}
         }
      }
   }
   
   $VC_P1_GB1_Label01.Text = $MetricsHash.Classes.Totals.CW
   $VC_P1_GB1_Label02.Text = $MetricsHash.Classes.Totals.DC
   $VC_P1_GB1_Label03.Text = $MetricsHash.Classes.Totals.GWF
   $VC_P1_GB1_Label04.Text = $MetricsHash.Classes.Totals.GF
   $VC_P1_GB1_Label05.Text = $MetricsHash.Classes.Totals.HR
   $VC_P1_GB1_Label06.Text = $MetricsHash.Classes.Totals.OP
   $VC_P1_GB1_Label07.Text = $MetricsHash.Classes.Totals.SW
   $VC_P1_GB1_Label08.Text = $MetricsHash.Classes.Totals.TR
   
   $VC_P1_GB2_Label01.Text = $MetricsHash.Classes.ByAccount.CW
   $VC_P1_GB2_Label02.Text = $MetricsHash.Classes.ByAccount.DC
   $VC_P1_GB2_Label03.Text = $MetricsHash.Classes.ByAccount.GWF
   $VC_P1_GB2_Label04.Text = $MetricsHash.Classes.ByAccount.GF
   $VC_P1_GB2_Label05.Text = $MetricsHash.Classes.ByAccount.HR
   $VC_P1_GB2_Label06.Text = $MetricsHash.Classes.ByAccount.OP
   $VC_P1_GB2_Label07.Text = $MetricsHash.Classes.ByAccount.SW
   $VC_P1_GB2_Label08.Text = $MetricsHash.Classes.ByAccount.TR

   Export-Clixml -Path "$($SettingsHash.DataLocation)\ClassMetrics.xml" -InputObject $MetricsHash.Classes
}

Function ViewClasses-Update { 
   $MetricsHash.Classes.Totals.CW  = 0
   $MetricsHash.Classes.Totals.DC  = 0
   $MetricsHash.Classes.Totals.GWF = 0
   $MetricsHash.Classes.Totals.GF  = 0
   $MetricsHash.Classes.Totals.HR  = 0
   $MetricsHash.Classes.Totals.OP  = 0
   $MetricsHash.Classes.Totals.SW  = 0
   $MetricsHash.Classes.Totals.TR  = 0
   
   $MetricsHash.Classes.ByAccount.CW  = 0
   $MetricsHash.Classes.ByAccount.DC  = 0
   $MetricsHash.Classes.ByAccount.GWF = 0
   $MetricsHash.Classes.ByAccount.GF  = 0
   $MetricsHash.Classes.ByAccount.HR  = 0
   $MetricsHash.Classes.ByAccount.OP  = 0
   $MetricsHash.Classes.ByAccount.SW  = 0
   $MetricsHash.Classes.ByAccount.TR  = 0
   
   ForEach ( $account in $DataHash.Keys ) {
      $VariableHash.MembersClasses = @()
      ForEach ( $Class in $DataHash.("$account").MemberClasses) {
         If ($VariableHash.MembersClasses -notcontains $Class) {$VariableHash.MembersClasses += $Class}
      }
      $VariableHash.MembersClasses | % {
         Switch ($_) {
            "Control Wizard" {$MetricsHash.Classes.ByAccount.CW++;Break}
            "Devoted Cleric" {$MetricsHash.Classes.ByAccount.DC++;Break}
            "Great Weapon Fighter" {$MetricsHash.Classes.ByAccount.GWF++;Break}
            "Guardian Fighter" {$MetricsHash.Classes.ByAccount.GF++;Break}
            "Hunter Ranger" {$MetricsHash.Classes.ByAccount.HR++;Break}
            "Oathbound Paladin" {$MetricsHash.Classes.ByAccount.OP++;Break}
            "Scourge Warlock" {$MetricsHash.Classes.ByAccount.SW++;Break}
            "Trickster Rogue" {$MetricsHash.Classes.ByAccount.TR++;Break}
         }
      }
      $DataHash.("$account").Characters.Keys | % {
         If ($DataHash.($Account).Characters.($_).Status -like "Member") {
            Switch ($DataHash.($Account).Characters.($_).Class) {
               "Control Wizard" {$MetricsHash.Classes.Totals.CW++;Break}
               "Devoted Cleric" {$MetricsHash.Classes.Totals.DC++;Break}
               "Great Weapon Fighter" {$MetricsHash.Classes.Totals.GWF++;Break}
               "Guardian Fighter" {$MetricsHash.Classes.Totals.GF++;Break}
               "Hunter Ranger" {$MetricsHash.Classes.Totals.HR++;Break}
               "Oathbound Paladin" {$MetricsHash.Classes.Totals.OP++;Break}
               "Scourge Warlock" {$MetricsHash.Classes.Totals.SW++;Break}
               "Trickster Rogue" {$MetricsHash.Classes.Totals.TR++;Break}
            }
         }
      }
   }
   
   $VC_P1_GB1_Label01.Text = $MetricsHash.Classes.Totals.CW
   $VC_P1_GB1_Label02.Text = $MetricsHash.Classes.Totals.DC
   $VC_P1_GB1_Label03.Text = $MetricsHash.Classes.Totals.GWF
   $VC_P1_GB1_Label04.Text = $MetricsHash.Classes.Totals.GF
   $VC_P1_GB1_Label05.Text = $MetricsHash.Classes.Totals.HR
   $VC_P1_GB1_Label06.Text = $MetricsHash.Classes.Totals.OP
   $VC_P1_GB1_Label07.Text = $MetricsHash.Classes.Totals.SW
   $VC_P1_GB1_Label08.Text = $MetricsHash.Classes.Totals.TR
   
   $VC_P1_GB2_Label01.Text = $MetricsHash.Classes.ByAccount.CW
   $VC_P1_GB2_Label02.Text = $MetricsHash.Classes.ByAccount.DC
   $VC_P1_GB2_Label03.Text = $MetricsHash.Classes.ByAccount.GWF
   $VC_P1_GB2_Label04.Text = $MetricsHash.Classes.ByAccount.GF
   $VC_P1_GB2_Label05.Text = $MetricsHash.Classes.ByAccount.HR
   $VC_P1_GB2_Label06.Text = $MetricsHash.Classes.ByAccount.OP
   $VC_P1_GB2_Label07.Text = $MetricsHash.Classes.ByAccount.SW
   $VC_P1_GB2_Label08.Text = $MetricsHash.Classes.ByAccount.TR

   Export-Clixml -Path "$($SettingsHash.DataLocation)\ClassMetrics.xml" -InputObject $MetricsHash.Classes
}