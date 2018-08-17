 <########################################################################################
Title: Form-Elements
Arthur(s): D. Lamberson & J. Wills
Contributors: M. Lewark
Created: 05 October 2015
Last Edited: 25 October 2016
Version: 1.8.1
#########################################################################################
Description: Account controls for local users and minimal control over domain users.

Version  Date     User              Updates
1.0.0    08OCT15  Wills/Lamberson   - Initial Release
1.0.1    28OCT15  Wills             - Fixed Syntax to comply with POSH formatting
                                    - Added more parameters for more creation flexibility 
1.1.0    10NOV15  Wills             - Added Create-Form
                                    - enable ability for Textbox, Labels, Buttons to be 
                                      added to a form by passing the $Form variable as -
                                      form <form name> 
                                    - Enabled ability for RadioButton and CheckBox to be 
                                      added to GroupBox by passing the $Group variable as 
                                      a -Group <group name>
                                    - Standarized Parameter Names for all Functions
                                    - Added Additional Properties for all functions.
1.1.1    06JAN16  Wills             - Added ListBox and ComboBox
                                    - Changed default font pt to 12
1.2.0    04FEB16  Lamberson         - Fixed the Boolean parameters that were being 
                                      labeled as String
                                    - Create-Richtextbox now creates a rich textbox 
                                      instead of just a testbox
1.3.0    04FEB16  Lamberson         - Added Create-Picturebox
1.4.0    05FEB16  Lamberson         - Transcribed over to RCI with minor formating
1.5.0    22MAR16  Lamberson         - Added Create-ProgressBar, Create-PanelBox,
                                      Create-MenuStrip, and Create-ToolStripMenuItem
                                    - Added more options to already existing functions
                                      such as Enabled, Visible, Font, BackColor, Hash,
                                      ReadOnly, Location, and Size.
1.5.1    28APR16  Lamberson         - Added ThreeState switch and Autosize boolean 
                                      to Create-Checkbox.
1.6.0    03MAY16  Lamberson         - Added Open-FolderBrowser function
1.7.0    26MAY16  Lamberson         - Added Create-TabControl, Create-TabPage, and
                                      Create-DataGridView functions
1.7.1    17AUG16  Lamberson         - Added parameter 'AllowDrop' for RichTextBox & 
                                      TextBox. This enables Drag-drop to be configured
                                    - Added Maximum and Minimum sizes for Forms
                                    - Added MaxLength for RichTextBoxes
                                    - Added OpenFile Function
1.8.0    14SEP16  Lamberson         - Added TextAlign to Create-Button
                                    - Added TextAlign to Create-Radio & Create-Checkbox
                                    - Added Create-TreeView
                                    - Added Create-TreeNode
1.8.1    25OCT16  Lamberson         - Added AutoCheck to Create-Checkbox
                                    - Added BackgroundImageLayout to Create-Form
                                    - Added Create-ObjectiveChecklist
                                    - Added ShowIcon to Create-Form
                                    - Added Available to Create-ToolStripMenuItem
#>

Function Create-PictureBox {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]$xLoc,
      [Parameter(Mandatory=$True)]  [String]$yLoc,
      [Parameter(Mandatory=$True)]  [System.Drawing.Image]$Picture,
      [Parameter(Mandatory=$False)] [String]$SizeMode = "StretchImage",
      [Parameter(Mandatory=$True)]  [Array]$Form
   )

   $PictureBox           = New-Object System.Windows.Forms.PictureBox
   $PictureBox.Image     = $Picture
   $PictureBox.Location  = "$xLoc,$yLoc"
   $PictureBox.SizeMode  = $SizeMode
   $PictureBox.Size      = "$($Picture.Width),$($Picture.Height)"
   $PictureBox.BackColor = "Transparent"
   $PictureBox

   ForEach ( $i in $Form ) { $i.Controls.Add($PictureBox) }
}
Function Create-ToolStripMenuItem {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled      = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible      = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Available    = $True,
      [Parameter(Mandatory=$False)] [String]  $TextAlign    = "MiddleCenter",
      [Parameter(Mandatory=$False)] [String]  $Text         = "",
      [Parameter(Mandatory=$False)] [String]  $ForeColor    = "BLACK",
      [Parameter(Mandatory=$False)] [String]  $BackColor    = "Control",
      [Parameter(Mandatory=$False)] [String]  $Font         = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize     = "11",
      [Parameter(Mandatory=$False)] [String]  $FontStyle    = "Regular",
      [Parameter(Mandatory=$False)] [String]  $ShortcutKeys = "",
      [Parameter(Mandatory=$False)] [String]  $Alignment    = "Left",
      [Parameter(Mandatory=$False)] [Switch]  $AutoSize,
      [Parameter(Mandatory=$True)]            $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash,
      [Parameter(Mandatory=$False)] [Switch]  $MenuItem,
      [Parameter(Mandatory=$False)] [Switch]  $DropDownItem 
   )
   
   If ( $PSBoundParameters.ContainsKey('MenuItem') ) { $MenuItem = $True ; $DropDownItem = $False }
   ElseIf ( $PSBoundParameters.ContainsKey('DropDownItem') ) { $DropDownItem = $True ; $MenuItem = $False }
   Else { Write-Host "Must use either the `"MenuItem`" or `"DropDownItem`" switch" ; Return }

   $ToolStripMenuItem      = New-Object System.Windows.Forms.ToolStripMenuItem
   $ToolStripMenuItem.Text = $Text
   $ToolStripMenuItem.Name = $Name
   If ($AutoSize) { $ToolStripMenuItem.AutoSize = $True } Else { $ToolStripMenuItem.AutoSize = $False }
   $ToolStripMenuItem.TextAlign    = $TextAlign
   If ( $ShortcutKeys -ne "" ) { $ToolStripMenuItem.ShortcutKeys = $ShortcutKeys }
   $ToolStripMenuItem.Enabled      = $Enabled
   $ToolStripMenuItem.Visible      = $Visible
   $ToolStripMenuItem.Alignment    = $Alignment
   $ToolStripMenuItem.Available    = $Available
   $ToolStripMenuItem.ForeColor    = $ForeColor
   $ToolStripMenuItem.BackColor    = $BackColor
   $ToolStripMenuItem.Font         = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $ToolStripMenuItem

  If ( $MenuItem ) {
      $ErrorActionPreference = 'Stop'
      Try { [System.Windows.Forms.MenuStrip]$Form.Items.Add($ToolStripMenuItem) } Catch {}
      $ErrorActionPreference = 'Continue'
   }
   If ( $DropDownItem ) { [System.Windows.Forms.ToolStripMenuItem]$Form.DropDownItems.Add($ToolStripMenuItem) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($ToolStripMenuItem.Name) = $ToolStripMenuItem } }
}
Function Create-MenuStrip {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $Location  = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width     = 295,
      [Parameter(Mandatory=$False)] [String]  $Height    = 25,
      [Parameter(Mandatory=$False)] [String]  $Size      = "$Width,$Height",
      [Parameter(Mandatory=$False)] [Boolean] $Enabled   = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible   = $True,
      [Parameter(Mandatory=$False)] [String]  $Text      = "",
      [Parameter(Mandatory=$False)] [String]  $Dock      = "Top",
      [Parameter(Mandatory=$False)] [String]  $ForeColor = "BLACK",
      [Parameter(Mandatory=$False)] [String]  $BackColor = "Control",
      [Parameter(Mandatory=$False)] [String]  $Font      = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize  = "11",
      [Parameter(Mandatory=$False)] [String]  $FontStyle = "Regular",
      [Parameter(Mandatory=$False)] [Switch]  $MainMenuStrip,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $MenuStrip           = New-Object System.Windows.Forms.MenuStrip
   $MenuStrip.Text      = $Text
   $MenuStrip.Name      = $Name
   $MenuStrip.Location  = $Location
   $MenuStrip.Size      = $Size
   $MenuStrip.Dock      = $Dock
   $MenuStrip.Enabled   = $Enabled
   $MenuStrip.Visible   = $Visible
   $MenuStrip.ForeColor = $ForeColor
   $MenuStrip.BackColor = $BackColor
   $MenuStrip.Font      = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $MenuStrip

   ForEach ( $i in $Form ) { $i.Controls.Add($MenuStrip) ; If ($MainMenuStrip) {$i.MainMenuStrip = $MenuStrip} }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($MenuStrip.Name) = $MenuStrip } }
}
Function Create-TabControl {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc       = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc       = 0,
      [Parameter(Mandatory=$False)] [String]  $Location   = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width      = 295,
      [Parameter(Mandatory=$False)] [String]  $Height     = 25,
      [Parameter(Mandatory=$False)] [String]  $Size       = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $ForeColor  = "BLACK",
      [Parameter(Mandatory=$False)] [String]  $Alignment  = "Top",
      [Parameter(Mandatory=$False)] [String]  $Appearance = "Normal",
      [Parameter(Mandatory=$False)] [String]  $Font       = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize   = "11",
      [Parameter(Mandatory=$False)] [String]  $FontStyle  = "Regular",
      [Parameter(Mandatory=$False)] [Boolean] $Enabled    = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible    = $True,
      [Parameter(Mandatory=$False)] [Boolean] $HotTrack   = $False,
      [Parameter(Mandatory=$False)] [Boolean] $Multiline  = $False,
      [Parameter(Mandatory=$True)]            $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash,
      [Parameter(Mandatory=$False)] [System.Drawing.Bitmap] $BackgroundImage = $Null
   )
   
   $TabControl            = New-Object System.Windows.Forms.TabControl
   $TabControl.Text       = $Text
   $TabControl.Name       = $Name
   $TabControl.Location   = $Location
   $TabControl.Size       = $Size
   $TabControl.Enabled    = $Enabled
   $TabControl.Visible    = $Visible
   $TabControl.Multiline  = $Multiline
   $TabControl.HotTrack   = $HotTrack
   $TabControl.Appearance = $Appearance
   $TabControl.Alignment  = $Alignment
   $TabControl.ForeColor  = $ForeColor
   $TabControl.Font       = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   If ( $BackgroundImage -ne $Null ) { $TabControl.BackgroundImage = $BackgroundImage }
   $TabControl

   ForEach ( $i in $Form ) { $i.Controls.Add($TabControl) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($TabControl.Name) = $TabControl } }
}
Function Create-TabPage {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled   = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible   = $True,
      [Parameter(Mandatory=$False)] [String]  $Text      = "",
      [Parameter(Mandatory=$False)] [String]  $ForeColor = "BLACK",
      [Parameter(Mandatory=$False)] [String]  $BackColor = "Control",
      [Parameter(Mandatory=$False)] [String]  $Font      = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize  = "11",
      [Parameter(Mandatory=$False)] [String]  $FontStyle = "Regular",
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $TabPage           = New-Object System.Windows.Forms.TabPage
   $TabPage.Text      = $Text
   $TabPage.Name      = $Name
   $TabPage.Enabled   = $Enabled
   $TabPage.Visible   = $Visible
   $TabPage.ForeColor = $ForeColor
   $TabPage.BackColor = $BackColor
   $TabPage.Font      = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $TabPage

   ForEach ( $i in $Form ) { $i.Controls.Add($TabPage) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($TabPage.Name) = $TabPage } }
}
Function Create-DataGridView {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc            = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc            = 0,
      [Parameter(Mandatory=$False)] [String]  $Location        = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width           = 295,
      [Parameter(Mandatory=$False)] [String]  $Height          = 25,
      [Parameter(Mandatory=$False)] [String]  $Size            = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $ForeColor       = "BLACK",
      [Parameter(Mandatory=$False)] [String]  $BackgroundColor = "BLACK",
      [Parameter(Mandatory=$False)] [Boolean] $Enabled         = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible         = $True,
      [Parameter(Mandatory=$False)] [Boolean] $MultiSelect     = $False,
      [Parameter(Mandatory=$False)] [Int32]   $ColumnCount              = 0,
      [Parameter(Mandatory=$False)] [Int32]   $ColumnHeadersHeight      = 21,
      [Parameter(Mandatory=$False)] [Boolean] $AllowUserToAddRows       = $False,
      [Parameter(Mandatory=$False)] [Boolean] $AllowUserToDeleteRows    = $False,
      [Parameter(Mandatory=$False)] [Boolean] $AllowUserToResizeRows    = $False,
      [Parameter(Mandatory=$False)] [Boolean] $AllowUserToOrderColumns  = $True,
      [Parameter(Mandatory=$False)] [Boolean] $AllowUserToResizeColumns = $True,
      [Parameter(Mandatory=$False)] [Switch]  $ReadOnly,
      [Parameter(Mandatory=$True)]            $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )
   
   $DataGridView                     = New-Object System.Windows.Forms.DataGridView
   $DataGridView.Text                = $Text
   $DataGridView.Name                = $Name
   $DataGridView.Location            = $Location
   $DataGridView.ColumnCount         = $ColumnCount
   $DataGridView.ColumnHeadersHeight = $ColumnHeadersHeight
   $DataGridView.Size                = $Size
   $DataGridView.Enabled             = $Enabled
   $DataGridView.MultiSelect         = $MultiSelect
   $DataGridView.Visible             = $Visible
   $DataGridView.ForeColor           = $ForeColor
   $DataGridView.BackgroundColor     = $BackgroundColor
   $DataGridView.AllowUserToAddRows       = $AllowUserToAddRows
   $DataGridView.AllowUserToDeleteRows    = $AllowUserToDeleteRows
   $DataGridView.AllowUserToResizeRows    = $AllowUserToResizeRows
   $DataGridView.AllowUserToOrderColumns  = $AllowUserToOrderColumns
   $DataGridView.AllowUserToResizeColumns = $AllowUserToResizeColumns
   
   If ( $ReadOnly ) { $DataGridView.ReadOnly = $True }
   $DataGridView

   ForEach ( $i in $Form ) { $i.Controls.Add($DataGridView) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($DataGridView.Name) = $DataGridView } }
}
Function Create-ProgressBar {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc     = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc     = 0,
      [Parameter(Mandatory=$False)] [String]  $Location = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width    = 295,
      [Parameter(Mandatory=$False)] [String]  $Height   = 25,
      [Parameter(Mandatory=$False)] [String]  $Size     = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $Step     = 1,
      [Parameter(Mandatory=$False)] [String]  $Minimum  = 0,
      [Parameter(Mandatory=$False)] [String]  $Maximum  = 100,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled  = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible  = $True,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $ProgressBar          = New-Object System.Windows.Forms.ProgressBar
   $ProgressBar.Location = $Location
   $ProgressBar.Size     = $Size
   $ProgressBar.Name     = $Name
   $ProgressBar.Step     = $Step
   $ProgressBar.Maximum  = $Maximum
   $ProgressBar.Minimum  = $Minimum
   $ProgressBar.Enabled  = $Enabled
   $ProgressBar.Visible  = $Visible
   $ProgressBar

   ForEach ( $i in $Form ) { $i.Controls.Add($ProgressBar) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($ProgressBar.Name) = $ProgressBar } }
}
Function Create-RichTextBox {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc        = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc        = 0,
      [Parameter(Mandatory=$False)] [String]  $Location    = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width       = 295,
      [Parameter(Mandatory=$False)] [String]  $Height      = 30,
      [Parameter(Mandatory=$False)] [String]  $Size        = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $MaxLength   = 2147483647,
      [Parameter(Mandatory=$False)] [Boolean] $MultiLine   = $False,
      [Parameter(Mandatory=$False)] [Boolean] $WordWrap    = $False,
      [Parameter(Mandatory=$False)] [Boolean] $AllowDrop   = $False,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled     = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible     = $True,
      [Parameter(Mandatory=$False)] [String]  $ScrBar      = "None",
      [Parameter(Mandatory=$False)] [String]  $Text        = "",
      [Parameter(Mandatory=$False)] [String]  $ForeColor   = "GRAY",
      [Parameter(Mandatory=$False)] [String]  $BackColor   = "BLACK",
      [Parameter(Mandatory=$False)] [String]  $Font        = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize    = "11",
      [Parameter(Mandatory=$False)] [String]  $FontStyle   = "Regular",
      [Parameter(Mandatory=$False)] [String]  $BorderStyle = "Fixed3D",
      [Parameter(Mandatory=$False)] [Switch]  $ReadOnly,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $Box             = New-Object System.Windows.Forms.RichTextBox
   $Box.Location    = $Location
   $Box.Size        = $Size
   $Box.Text        = $Text
   $Box.Name        = $Name
   $Box.AllowDrop   = $AllowDrop
   $Box.MaxLength   = $MaxLength
   $Box.Multiline   = $MultiLine
   $Box.WordWrap    = $WordWrap
   $Box.ScrollBars  = $ScrBar
   $Box.Enabled     = $Enabled
   $Box.Visible     = $Visible
   $Box.ForeColor   = $ForeColor
   $Box.BackColor   = $BackColor
   $Box.BorderStyle = $BorderStyle
   $Box.Font        = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   If ( $ReadOnly ) { $Box.ReadOnly = $True }
   $Box

   ForEach ( $i in $Form ) { $i.Controls.Add($Box) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($Box.Name) = $Box } }
}
Function Create-TextBox {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc        = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc        = 0,
      [Parameter(Mandatory=$False)] [String]  $Location    = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width       = 295,
      [Parameter(Mandatory=$False)] [String]  $Height      = 25,
      [Parameter(Mandatory=$False)] [String]  $Size        = "$Width,$Height",
      [Parameter(Mandatory=$False)] [Boolean] $MultiLine   = $False,
      [Parameter(Mandatory=$False)] [Boolean] $WordWrap    = $False,
      [Parameter(Mandatory=$False)] [Boolean] $AllowDrop   = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible     = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled     = $True,
      [Parameter(Mandatory=$False)] [String]  $ScrBar      = "None",
      [Parameter(Mandatory=$False)] [String]  $Text        = "",
      [Parameter(Mandatory=$False)] [String]  $TextAlign   = "Left",
      [Parameter(Mandatory=$False)] [String]  $Tag         = "",
      [Parameter(Mandatory=$False)] [String]  $ForeColor   = "Black",
      [Parameter(Mandatory=$False)] [String]  $BackColor   = "Control",
      [Parameter(Mandatory=$False)] [String]  $Font        = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize    = "9",
      [Parameter(Mandatory=$False)] [String]  $FontStyle   = "Regular",
      [Parameter(Mandatory=$False)] [String]  $BorderStyle = "Fixed3D",
      [Parameter(Mandatory=$False)] [Int32]   $MaxLength   = 32767,
      [Parameter(Mandatory=$False)] [Switch]  $PasswordChar,
      [Parameter(Mandatory=$False)] [Switch]  $ReadOnly,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $Box             = New-Object System.Windows.Forms.TextBox
   $Box.TabIndex    = 0
   $Box.Location    = $Location
   $Box.Size        = $Size
   $Box.Text        = $Text
   $Box.Tag         = $Tag
   $Box.Name        = $Name
   $Box.AllowDrop   = $AllowDrop
   $Box.Multiline   = $MultiLine
   $Box.WordWrap    = $WordWrap
   $Box.ScrollBars  = $ScrBar
   $Box.Enabled     = $Enabled
   $Box.Visible     = $Visible
   $Box.TextAlign   = $TextAlign
   $Box.BorderStyle = $BorderStyle
   $Box.MaxLength   = $MaxLength
   $Box.ForeColor   = $ForeColor
   $Box.BackColor   = $BackColor
   $Box.Font        = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   If ( $PasswordChar ) { $Box.PasswordChar = "•" }
   If ( $ReadOnly ) { $Box.ReadOnly = $True }
   $Box

   ForEach ( $i in $Form ) { $i.Controls.Add($Box) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($Box.Name) = $Box } }
}
Function Create-Label {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $Location  = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width     = 295,
      [Parameter(Mandatory=$False)] [String]  $Height    = 20,
      [Parameter(Mandatory=$False)] [String]  $Size      = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $Text      = "",
      [Parameter(Mandatory=$False)] [String]  $ForeColor = "Black",
      [Parameter(Mandatory=$False)] [String]  $TextAlign = "MiddleLeft",
      [Parameter(Mandatory=$False)] [String]  $BackColor = "Transparent",
      [Parameter(Mandatory=$False)] [Boolean] $AutoSize  = $False,
      [Parameter(Mandatory=$False)] [Boolean] $Visible   = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled   = $True,
      [Parameter(Mandatory=$False)] [String]  $Font      = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize  = "12",
      [Parameter(Mandatory=$False)] [String]  $FontStyle = "Regular",
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $Label           = New-Object System.Windows.Forms.Label
   $Label.Location  = $Location
   $Label.Size      = $Size
   $Label.Name      = $Name
   $Label.ForeColor = $ForeColor
   $Label.BackColor = $BackColor
   $Label.Font      = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $Label.AutoSize  = $AutoSize
   $Label.Visible   = $Visible
   $Label.Enabled   = $Enabled
   $Label.TextAlign = $TextAlign
   $Label.Text      = $Text
   $Label

   ForEach ( $i in $Form ) { $i.Controls.Add($Label) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($Label.Name) = $Label } }
}
Function Create-Button {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc       = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc       = 0,
      [Parameter(Mandatory=$False)] [String]  $Location   = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width      = 50,
      [Parameter(Mandatory=$False)] [String]  $Height     = 30,
      [Parameter(Mandatory=$False)] [String]  $Size       = "$Width,$Height",
      [Parameter(Mandatory=$False)] [Boolean] $AutoSize   = $False,
      [Parameter(Mandatory=$False)] [Boolean] $Visible    = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled    = $True,
      [Parameter(Mandatory=$False)] [String]  $Text       = "",
      [Parameter(Mandatory=$False)] [String]  $TextAlign  = "MiddleCenter",
      [Parameter(Mandatory=$False)] [String]  $Font       = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize   = "12",
      [Parameter(Mandatory=$False)] [String]  $FontStyle  = "Regular",
      [Parameter(Mandatory=$False)] [String]  $BackColor  = "WHITE",
      [Parameter(Mandatory=$False)] [String]  $ForeColor  = "BLACK",
      [Parameter(Mandatory=$False)] [String]  $FlatStyle  = "Standard",
      [Parameter(Mandatory=$False)] [String]  $BorderSize = 1,
      [Parameter(Mandatory=$False)]   $Margin     = 6,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $Button            = New-Object System.Windows.Forms.Button
   $Button.Location   = $Location
   $Button.Size       = $Size
   $Button.Name       = $Name
   $Button.Text       = $Text
   $Button.TextAlign  = $TextAlign
   $Button.AutoSize   = $AutoSize
   $Button.Enabled    = $Enabled
   $Button.Font       = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $Button.BackColor  = $BackColor
   $Button.ForeColor  = $ForeColor
   $Button.FlatStyle  = $FlatStyle
   $Button.Visible    = $Visible
   $Button.Margin     = $Margin
   $Button.FlatAppearance.BorderSize = $BorderSize
   $Button

   ForEach ( $i in $Form ) { $i.Controls.Add($Button) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($Button.Name) = $Button } }
}
Function Create-GroupBox {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$False)] [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $Location  = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width     = 295,
      [Parameter(Mandatory=$False)] [String]  $Height    = 26,
      [Parameter(Mandatory=$False)] [String]  $Size      = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $Text      = "",
      [Parameter(Mandatory=$False)] [String]  $Font      = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize  = "12",
      [Parameter(Mandatory=$False)] [String]  $FontStyle = "Regular",
      [Parameter(Mandatory=$False)] [String]  $ForeColor = "Black",
      [Parameter(Mandatory=$False)] [String]  $BackColor = "Control",
      [Parameter(Mandatory=$False)] [Boolean] $AutoSize  = $False,
      [Parameter(Mandatory=$False)] [Boolean] $Visible   = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled   = $True,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $GroupBox           = New-Object System.Windows.Forms.GroupBox
   $GroupBox.Location  = $Location
   $GroupBox.Size      = $Size
   $GroupBox.Text      = $Text
   $GroupBox.Name      = $Name
   $GroupBox.Font      = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $GroupBox.AutoSize  = $AutoSize
   $GroupBox.Enabled   = $Enabled
   $GroupBox.Visible   = $Visible
   $GroupBox.BackColor = $BackColor
   $GroupBox.ForeColor = $ForeColor
   $GroupBox

   ForEach ( $i in $Form ) { $i.Controls.Add($GroupBox) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($GroupBox.Name) = $GroupBox } }
}
Function Create-PanelBox {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$False)] [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc        = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc        = 0,
      [Parameter(Mandatory=$False)] [String]  $Location    = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width       = 295,
      [Parameter(Mandatory=$False)] [String]  $Height      = 26,
      [Parameter(Mandatory=$False)] [String]  $Size        = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $Font        = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize    = "12",
      [Parameter(Mandatory=$False)] [String]  $FontStyle   = "Regular",
      [Parameter(Mandatory=$False)] [String]  $ForeColor   = "Black",
      [Parameter(Mandatory=$False)] [String]  $BackColor   = "Control",
      [Parameter(Mandatory=$False)] [String]  $BorderStyle = "None",
      [Parameter(Mandatory=$False)] [Boolean] $Enabled     = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible     = $True,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $PanelBox             = New-Object System.Windows.Forms.Panel
   $PanelBox.Location    = $Location
   $PanelBox.Size        = $Size
   $PanelBox.Name        = $Name
   $PanelBox.Enabled     = $Enabled
   $PanelBox.Visible     = $Visible
   $PanelBox.BorderStyle = $BorderStyle
   $PanelBox.Font        = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $PanelBox.BackColor   = $BackColor
   $PanelBox.ForeColor   = $ForeColor
   $PanelBox

   ForEach ( $i in $Form ) { $i.Controls.Add($PanelBox) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($PanelBox.Name) = $PanelBox } }
}
Function Create-Radio {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $Location  = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width     = 80,
      [Parameter(Mandatory=$False)] [String]  $Height    = 18,
      [Parameter(Mandatory=$False)] [String]  $Size      = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $Font      = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize  = "12",
      [Parameter(Mandatory=$False)] [String]  $FontStyle = "Regular",
      [Parameter(Mandatory=$False)] [String]  $BackColor = "Control",
      [Parameter(Mandatory=$False)] [String]  $ForeColor = "Black",
      [Parameter(Mandatory=$False)] [String]  $Text      = "",
      [Parameter(Mandatory=$False)] [String]  $TextAlign = "MiddleLeft",
      [Parameter(Mandatory=$False)] [Boolean] $Visible   = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled   = $True,
      [Parameter(Mandatory=$False)] [Switch]  $Checked,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $RadioButton           = New-Object System.Windows.Forms.RadioButton
   $RadioButton.Location  = $Location
   $RadioButton.Size      = $Size
   $RadioButton.Text      = $Text
   $RadioButton.TextAlign = $TextAlign
   $RadioButton.Name      = $Name
   $RadioButton.Checked   = $False
   $RadioButton.Visible   = $Visible
   $RadioButton.Enabled   = $Enabled
   $RadioButton.BackColor = $BackColor
   $RadioButton.ForeColor = $ForeColor
   $RadioButton.Font      = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   If ( $Checked ) { $RadioButton.Checked = $True }
   $RadioButton

   ForEach ( $i in $Form ) { $i.Controls.Add($RadioButton) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($RadioButton.Name) = $RadioButton } }
}
Function Create-CheckBox {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc      = 0,
      [Parameter(Mandatory=$False)] [String]  $Location  = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width     = 200,
      [Parameter(Mandatory=$False)] [String]  $Height    = 30,
      [Parameter(Mandatory=$False)] [String]  $Size      = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $Text      = "",
      [Parameter(Mandatory=$False)] [String]  $TextAlign = "MiddleLeft",
      [Parameter(Mandatory=$False)] [Boolean] $Visible   = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled   = $True,
      [Parameter(Mandatory=$False)] [Boolean] $AutoCheck = $True,
      [Parameter(Mandatory=$False)] [Boolean] $AutoSize  = $False,
      [Parameter(Mandatory=$False)] [String]  $BackColor = "Control",
      [Parameter(Mandatory=$False)] [String]  $ForeColor = "Black",
      [Parameter(Mandatory=$False)] [String]  $Font      = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize  = "12",
      [Parameter(Mandatory=$False)] [String]  $FontStyle = "Regular",
      [Parameter(Mandatory=$False)] [Switch]  $Checked,
      [Parameter(Mandatory=$False)] [Switch]  $ThreeState,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $CheckBox           = New-Object System.Windows.Forms.CheckBox
   $CheckBox.Location  = $Location
   $CheckBox.Size      = $Size
   $CheckBox.Text      = $Text
   $CheckBox.TextAlign = $TextAlign
   $CheckBox.Name      = $Name
   $CheckBox.Checked   = $False
   $CheckBox.Visible   = $Visible
   $CheckBox.Enabled   = $Enabled
   $CheckBox.BackColor = $BackColor
   $CheckBox.AutoSize  = $AutoSize
   $CheckBox.AutoCheck = $AutoCheck
   $CheckBox.ForeColor = $ForeColor
   $CheckBox.Font      = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   If ( $Checked ) { $CheckBox.Checked = $True }
   If ( $ThreeState ) { $CheckBox.ThreeState = $True }
   $CheckBox

   ForEach ( $i in $Form ) { $i.Controls.Add($CheckBox) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($CheckBox.Name) = $CheckBox } }
}
Function Create-ListBox {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc          = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc          = 0,
      [Parameter(Mandatory=$False)] [String]  $Location      = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width         = 100,
      [Parameter(Mandatory=$False)] [String]  $Height        = 100,
      [Parameter(Mandatory=$False)] [String]  $Size          = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $Font          = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize      = "12",
      [Parameter(Mandatory=$False)] [String]  $FontStyle     = "Regular",
      [Parameter(Mandatory=$False)] [String]  $ForeColor     = "BLACK",
      [Parameter(Mandatory=$False)] [String]  $BackColor     = "WHITE",
      [Parameter(Mandatory=$False)] [String]  $Text          = "",
      [Parameter(Mandatory=$False)] [String]  $SelectionMode = "One",
      [Parameter(Mandatory=$False)] [Boolean] $Visible       = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled       = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Sorted        = $True,
      [Parameter(Mandatory=$False)] [Switch]  $HorizontalScrollbar,
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $ListBox               = New-Object System.Windows.Forms.ListBox
   $ListBox.Location      = $Location
   $ListBox.Size          = $Size
   $ListBox.Name          = $Name
   $ListBox.Text          = $Text
   $ListBox.Visible       = $Visible
   $ListBox.ForeColor     = $ForeColor
   $ListBox.BackColor     = $BackColor
   $ListBox.SelectionMode = $SelectionMode
   $ListBox.Enabled       = $Enabled
   $ListBox.Sorted        = $Sorted
   $ListBox.Font          = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   If ($HorizontalScrollbar ) {$ListBox.HorizontalScrollbar = $True }
   $ListBox

   ForEach ( $i in $Form ) { $i.Controls.Add($ListBox) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($ListBox.Name) = $ListBox } }
}
Function Create-ComboBox {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc          = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc          = 0,
      [Parameter(Mandatory=$False)] [String]  $Location      = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width         = 200,
      [Parameter(Mandatory=$False)] [String]  $Height        = 30,
      [Parameter(Mandatory=$False)] [String]  $Size          = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $DropDownStyle = "DropDownList",
      [Parameter(Mandatory=$False)] [Boolean] $AutoSize      = $False,
      [Parameter(Mandatory=$False)] [Boolean] $Visible       = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Enabled       = $True,
      [Parameter(Mandatory=$False)] [String]  $Font          = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize      = "12",
      [Parameter(Mandatory=$False)] [String]  $FontStyle     = "Regular",
      [Parameter(Mandatory=$True)]  [Array]   $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )

   $ComboBox               = New-Object System.Windows.Forms.ComboBox
   $ComboBox.DropDownStyle = $DropDownStyle
   $ComboBox.Location      = $Location
   $ComboBox.Size          = $Size
   $ComboBox.Name          = $Name
   $ComboBox.Font          = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $ComboBox.AutoSize      = $AutoSize
   $ComboBox.Visible       = $Visible
   $ComboBox.Enabled       = $Enabled
   $ComboBox
   
   ForEach ( $i in $Form ) { $i.Controls.Add($ComboBox) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($ComboBox.Name) = $ComboBox } }
}
Function Create-TreeView {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $xLoc        = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc        = 0,
      [Parameter(Mandatory=$False)] [String]  $Location    = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $Width       = 295,
      [Parameter(Mandatory=$False)] [String]  $Height      = 25,
      [Parameter(Mandatory=$False)] [String]  $Size        = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $BorderStyle = "Fixed3D",
      [Parameter(Mandatory=$False)] [String]  $Font        = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize    = "11",
      [Parameter(Mandatory=$False)] [String]  $FontStyle   = "Regular",
      [Parameter(Mandatory=$False)] [String]  $ForeColor   = "WindowText",
      [Parameter(Mandatory=$False)] [String]  $BackColor   = "Window",
      [Parameter(Mandatory=$False)] [String]  $ItemHeight  = "16",
      [Parameter(Mandatory=$False)] [Boolean] $Enabled     = $True,
      [Parameter(Mandatory=$False)] [Boolean] $Visible     = $True,
      [Parameter(Mandatory=$False)] [Boolean] $HotTracking = $False,
      [Parameter(Mandatory=$False)] [Boolean] $LabelEdit   = $False,
      [Parameter(Mandatory=$False)] [Boolean] $CheckBoxes  = $False,
      [Parameter(Mandatory=$True)]            $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash
   )
   
   $TreeView             = New-Object System.Windows.Forms.TreeView
   $TreeView.Text        = $Text
   $TreeView.Name        = $Name
   $TreeView.Location    = $Location
   $TreeView.Size        = $Size
   $TreeView.BorderStyle = $BorderStyle
   $TreeView.ForeColor   = $ForeColor
   $TreeView.BackColor   = $BackColor
   $TreeView.ItemHeight  = $ItemHeight
   $TreeView.Enabled     = $Enabled
   $TreeView.Visible     = $Visible
   $TreeView.HotTracking = $HotTracking
   $TreeView.LabelEdit   = $LabelEdit
   $TreeView.CheckBoxes  = $CheckBoxes
   $TreeView.Font       = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $TreeView

   ForEach ( $i in $Form ) { $i.Controls.Add($TreeView) }
   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($TreeView.Name) = $TreeView } }
}
Function Create-TreeNode {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $ForeColor   = "WindowText",
      [Parameter(Mandatory=$False)] [String]  $BackColor   = "Window",
      [Parameter(Mandatory=$False)] [String]  $Font        = "Microsoft San Serif",
      [Parameter(Mandatory=$False)] [String]  $FontSize    = "11",
      [Parameter(Mandatory=$False)] [String]  $FontStyle   = "Regular",
      [Parameter(Mandatory=$False)] [String]  $Text        = "",
      [Parameter(Mandatory=$False)] [String]  $ToolTipText = "",
      [Parameter(Mandatory=$False)] [Boolean] $Checked     = $False,
      [Parameter(Mandatory=$True)]            $Form,
      [Parameter(Mandatory=$False)] [Array]   $Hash,
      [Parameter(Mandatory=$False)] [Switch]  $Tree,
      [Parameter(Mandatory=$False)] [Switch]  $Node 
   )
   
   $TreeNode             = New-Object System.Windows.Forms.TreeNode
   $TreeNode.Text        = $Text
   $TreeNode.Name        = $Name
   $TreeNode.ForeColor   = $ForeColor
   $TreeNode.BackColor   = $BackColor
   $TreeNode.Text        = $Text
   $TreeNode.ToolTipText = $ToolTipText
   $TreeNode.Checked     = $Checked
   $TreeNode.NodeFont    = New-Object System.Drawing.Font($Font,$FontSize,[System.Drawing.FontStyle]::$FontStyle)
   $TreeNode
   
  If ( $Tree ) { Try { [System.Windows.Forms.TreeView]$Form.Nodes.Add($TreeNode) } Catch {} }
  If ( $Node ) { [System.Windows.Forms.TreeNode]$Form.Nodes.Add($TreeNode) }

   If ( $Hash -ne $Null ) { ForEach ( $i in $Hash ) { $i.$($TreeNode.Name) = $TreeNode } }
}
Function Create-Form {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [String]  $Name,
      [Parameter(Mandatory=$False)] [String]  $Text,
      [Parameter(Mandatory=$False)] [String]  $xLoc                  = 0,
      [Parameter(Mandatory=$False)] [String]  $yLoc                  = 0,
      [Parameter(Mandatory=$False)] [String]  $Location              = "$xLoc,$yLoc",
      [Parameter(Mandatory=$False)] [String]  $MinWidth              = 0,
      [Parameter(Mandatory=$False)] [String]  $MinHeight             = 0,
      [Parameter(Mandatory=$False)] [String]  $MinimumSize           = "$MinWidth,$MinHeight",
      [Parameter(Mandatory=$False)] [String]  $Width                 = 100,
      [Parameter(Mandatory=$False)] [String]  $Height                = 100,
      [Parameter(Mandatory=$False)] [String]  $Size                  = "$Width,$Height",
      [Parameter(Mandatory=$False)] [String]  $MaxWidth              = 100000,
      [Parameter(Mandatory=$False)] [String]  $MaxHeight             = 100000,
      [Parameter(Mandatory=$False)] [String]  $MaximumSize           = "$MaxWidth,$MaxHeight",
      [Parameter(Mandatory=$False)] [String]  $Color,
      [Parameter(Mandatory=$False)] [Boolean] $ShowInTaskBar         = $True,
      [Parameter(Mandatory=$False)] [Boolean] $ControlBox            = $True,
      [Parameter(Mandatory=$False)] [Boolean] $ShowIcon              = $True,
      [Parameter(Mandatory=$False)] [Boolean] $AutoSize              = $False,
      [Parameter(Mandatory=$False)] [String]  $FormBorderStyle       = "Fixed3D",
      [Parameter(Mandatory=$False)] [String]  $StartPosition         = "CenterScreen",
      [Parameter(Mandatory=$False)] [String]  $BackgroundImageLayout = "Tile",
      [Parameter(Mandatory=$False)] [System.Drawing.Icon]   $Icon            = $Null,
      [Parameter(Mandatory=$False)] [System.Drawing.Bitmap] $BackgroundImage = $Null
   )

   $Form                       = New-Object System.Windows.Forms.Form
   $Form.Location              = $Location
   $Form.MinimumSize           = $MinimumSize
   $Form.Size                  = $Size
   $Form.MaximumSize           = $MaximumSize
   $Form.Text                  = $Text
   $Form.Name                  = $Name
   $Form.AutoSize              = $AutoSize
   $Form.ShowIcon              = $ShowIcon
   $Form.BackColor             = $Color
   $Form.FormBorderStyle       = $FormBorderStyle
   $Form.MaximizeBox           = $False
   $Form.StartPosition         = $StartPosition
   $Form.ControlBox            = $ControlBox
   $Form.ShowInTaskbar         = $ShowInTaskBar
   $Form.BackgroundImageLayout = $BackgroundImageLayout
   If ( $Icon -ne $Null ) { $Form.Icon = $Icon }
   If ( $BackgroundImage -ne $Null ) { $Form.BackgroundImage = $BackgroundImage }

   $Form
}
Function Open-FolderBrowser {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$False)] [String] $Description = "",
      [Parameter(Mandatory=$False,
                 HelpMessage='32-bit flag that specifies internal features: &H0001, &H002, &H0004, &H0008, &H0010, &H0020, &H1000, &H2000, &H4000')]
                 [String]
                 $Option = "&H4000",
      [Parameter(Mandatory=$False,
                 HelpMessage='Use one of the values: 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21')]
                 [Int32]
                 $Root = 0,
      [Parameter(Mandatory=$False,
                 HelpMessage='Submit a string with the folder path (such as C:\Born) ')]
                 [String]
                 $Folder = ""
   )

   If ($Folder -eq "") { $Location = $Root }
   ElseIf ( !(Test-Path $Folder) ) { $Location = $Root }
   Else { $Location = $Folder }

   $SB = {
      $wShell = New-Object -ComObject Wscript.Shell
      Start-Sleep -Milliseconds 250
      $wShell.AppActivate('Browse For Files or Folders')
   }

   Start-Job -Name FileBrowser -ScriptBlock $SB | Out-Null
   $temp = ((New-Object -ComObject Shell.Application).BrowseForFolder(0,$Description,$Option,$Location)).Self #.Path
   Remove-Job -Name FileBrowser -Force | Out-Null
   Return $temp
}
Function OpenFile {
	$SelectOpenForm = New-Object System.Windows.Forms.OpenFileDialog
	$SelectOpenForm.Filter = "All Files (*.*)|*.*"
	$SelectOpenForm.InitialDirectory = ".\"
	$SelectOpenForm.Title = "Select a File to Open"
	$GetKey = $SelectOpenForm.ShowDialog()
	If ($GetKey -eq "OK") { $InputFileName = $SelectOpenForm.FileName }
    Return $InputFileName
}
Function Create-ObjectiveChecklist {
   [CmdletBinding()]
   param
   (
      [Parameter(Mandatory=$True)]  [Int]   $Max,
      [Parameter(Mandatory=$False)] [String]$Title     = "",
      [Parameter(Mandatory=$False)] [String]$Name      = "ObjectiveCheckbox",
      [Parameter(Mandatory=$False)] [String]$BackColor = "BLACK",
      [Parameter(Mandatory=$False)] [String]$ForeColor = "WHITE"
   )
   $FormHash.ObjectiveForm = Create-Form -Name "Form" -Text $Title -Size "10,10" -AutoSize 1 -Color BLACK -ShowIcon 0
      $FormHash.("$($Name)1") = Create-CheckBox -Name "$($Name)1" -Location "10,10" -AutoSize 1 `
      -TextAlign BOTTOMLEFT -AutoCheck 0 -BackColor TRANSPARENT -ForeColor $ForeColor -Form $FormHash.ObjectiveForm
      For ( $i = 2 ; $i -le $Max ; $i++ ) {
         $FormHash.("$($Name)$i") = Create-CheckBox -Name "$($Name)$i" `
         -Location "$($FormHash.("$($Name)$($i-1)").Left),$($FormHash.("$($Name)$($i-1)").Bottom+6)" `
         -AutoSize 1 -TextAlign BOTTOMLEFT -AutoCheck 0 -ForeColor $ForeColor -BackColor TRANSPARENT -Form $FormHash.ObjectiveForm
      }
   $FormHash.ObjectiveForm.Height = $FormHash.("ObjectiveCheckbox$Max").Bottom + 45
}