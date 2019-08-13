Namespace DataEntryFormSample
	Partial Public Class MainForm
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.components = New System.ComponentModel.Container()
			Me.spreadsheetControl1 = New DevExpress.XtraSpreadsheet.SpreadsheetControl()
			Me.spreadsheetBarController1 = New DevExpress.XtraSpreadsheet.UI.SpreadsheetBarController(Me.components)
			Me.bindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
			Me.panel1 = New System.Windows.Forms.Panel()
			Me.dataNavigator1 = New DevExpress.XtraEditors.DataNavigator()
			Me.spreadsheetBindingManager1 = New DataEntryFormSample.SpreadsheetBindingManager(Me.components)
			DirectCast(Me.spreadsheetBarController1, System.ComponentModel.ISupportInitialize).BeginInit()
			DirectCast(Me.bindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.panel1.SuspendLayout()
			Me.SuspendLayout()
			' 
			' spreadsheetControl1
			' 
			Me.spreadsheetControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
			Me.spreadsheetControl1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.spreadsheetControl1.Location = New System.Drawing.Point(0, 0)
			Me.spreadsheetControl1.Name = "spreadsheetControl1"
			Me.spreadsheetControl1.Options.Behavior.Selection.HideSelection = True
			Me.spreadsheetControl1.Options.Behavior.Selection.MoveActiveCellMode = DevExpress.XtraSpreadsheet.MoveActiveCellModeOnEnterPress.None
			Me.spreadsheetControl1.Options.HorizontalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden
			Me.spreadsheetControl1.Options.TabSelector.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetElementVisibility.Hidden
			Me.spreadsheetControl1.Options.VerticalScrollbar.Visibility = DevExpress.XtraSpreadsheet.SpreadsheetScrollbarVisibility.Hidden
			Me.spreadsheetControl1.Size = New System.Drawing.Size(898, 619)
			Me.spreadsheetControl1.TabIndex = 0
			Me.spreadsheetControl1.Text = "spreadsheetControl1"
'INSTANT VB NOTE: The following InitializeComponent event wireup was converted to a 'Handles' clause:
'ORIGINAL LINE: this.spreadsheetControl1.CustomCellEdit += new DevExpress.XtraSpreadsheet.SpreadsheetCustomCellEditEventHandler(this.spreadsheetControl1_CustomCellEdit);
'INSTANT VB NOTE: The following InitializeComponent event wireup was converted to a 'Handles' clause:
'ORIGINAL LINE: this.spreadsheetControl1.SelectionChanged += new System.EventHandler(this.SpreadsheetControl1_SelectionChanged);
'INSTANT VB NOTE: The following InitializeComponent event wireup was converted to a 'Handles' clause:
'ORIGINAL LINE: this.spreadsheetControl1.ProtectionWarning += new System.ComponentModel.HandledEventHandler(this.spreadsheetControl1_ProtectionWarning);
			' 
			' spreadsheetBarController1
			' 
			Me.spreadsheetBarController1.Control = Me.spreadsheetControl1
			' 
			' bindingSource1
			' 
			Me.bindingSource1.AllowNew = False
			' 
			' panel1
			' 
			Me.panel1.Controls.Add(Me.dataNavigator1)
			Me.panel1.Dock = System.Windows.Forms.DockStyle.Bottom
			Me.panel1.Location = New System.Drawing.Point(0, 619)
			Me.panel1.Name = "panel1"
			Me.panel1.Size = New System.Drawing.Size(898, 19)
			Me.panel1.TabIndex = 4
			' 
			' dataNavigator1
			' 
			Me.dataNavigator1.Buttons.Append.Enabled = False
			Me.dataNavigator1.Buttons.Append.Visible = False
			Me.dataNavigator1.Buttons.CancelEdit.Enabled = False
			Me.dataNavigator1.Buttons.CancelEdit.Visible = False
			Me.dataNavigator1.Buttons.EndEdit.Visible = False
			Me.dataNavigator1.Buttons.NextPage.Visible = False
			Me.dataNavigator1.Buttons.PrevPage.Visible = False
			Me.dataNavigator1.Buttons.Remove.Visible = False
			Me.dataNavigator1.DataSource = Me.bindingSource1
			Me.dataNavigator1.Dock = System.Windows.Forms.DockStyle.Left
			Me.dataNavigator1.Location = New System.Drawing.Point(0, 0)
			Me.dataNavigator1.Name = "dataNavigator1"
			Me.dataNavigator1.Size = New System.Drawing.Size(143, 19)
			Me.dataNavigator1.TabIndex = 3
			Me.dataNavigator1.Text = "dataNavigator1"
			Me.dataNavigator1.TextLocation = DevExpress.XtraEditors.NavigatorButtonsTextLocation.Center
			' 
			' spreadsheetBindingManager1
			' 
			Me.spreadsheetBindingManager1.Control = Me.spreadsheetControl1
			Me.spreadsheetBindingManager1.DataSource = Nothing
			' 
			' MainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(898, 638)
			Me.Controls.Add(Me.spreadsheetControl1)
			Me.Controls.Add(Me.panel1)
			Me.Name = "MainForm"
			Me.ShowIcon = False
			Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
			Me.Text = "Spreadsheet Data Entry Form Sample"
			DirectCast(Me.spreadsheetBarController1, System.ComponentModel.ISupportInitialize).EndInit()
			DirectCast(Me.bindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.panel1.ResumeLayout(False)
			Me.ResumeLayout(False)

		End Sub

		#End Region

		Private WithEvents spreadsheetControl1 As DevExpress.XtraSpreadsheet.SpreadsheetControl
		Private spreadsheetBarController1 As DevExpress.XtraSpreadsheet.UI.SpreadsheetBarController
		Private bindingSource1 As System.Windows.Forms.BindingSource
		Private panel1 As System.Windows.Forms.Panel
		Private dataNavigator1 As DevExpress.XtraEditors.DataNavigator
		Private spreadsheetBindingManager1 As SpreadsheetBindingManager
	End Class
End Namespace

