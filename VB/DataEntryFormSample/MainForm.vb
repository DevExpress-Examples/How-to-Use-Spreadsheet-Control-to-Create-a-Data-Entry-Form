Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Repository
Imports DevExpress.XtraSpreadsheet

Namespace DataEntryFormSample
	Partial Public Class MainForm
		Inherits XtraForm

		Private ReadOnly payrollData As New List(Of PayrollModel)()
		' ...

		Public Sub New()
			InitializeComponent()
			InitializePayrollData()
			LoadDocumentTemplate()
			BindCustomEditors()
			BindDataSource()
		End Sub

		Private Sub LoadDocumentTemplate()
			spreadsheetControl1.LoadDocument("PayrollCalculatorTemplate.xlsx")
			spreadsheetControl1.Document.History.IsEnabled = False
		End Sub

		Private Sub BindCustomEditors()
			Dim sheet = spreadsheetControl1.ActiveWorksheet
			' Assign custom editors to worksheet cells.
			sheet.CustomCellInplaceEditors.Add(sheet("D8"), CustomCellInplaceEditorType.Custom, "RegularHoursWorked")
			sheet.CustomCellInplaceEditors.Add(sheet("D10"), CustomCellInplaceEditorType.Custom, "VacationHours")
			sheet.CustomCellInplaceEditors.Add(sheet("D12"), CustomCellInplaceEditorType.Custom, "SickHours")
			sheet.CustomCellInplaceEditors.Add(sheet("D14"), CustomCellInplaceEditorType.Custom, "OvertimeHours")
			sheet.CustomCellInplaceEditors.Add(sheet("D16"), CustomCellInplaceEditorType.Custom, "OvertimeRate")
			sheet.CustomCellInplaceEditors.Add(sheet("D22"), CustomCellInplaceEditorType.Custom, "OtherDeduction")
		End Sub

		' Activate a cell editor when a user selects an editable cell.
		Private Sub SpreadsheetControl1_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles spreadsheetControl1.SelectionChanged
			Dim sheet = spreadsheetControl1.ActiveWorksheet
			If sheet IsNot Nothing Then
				Dim editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection)
				If editors.Count = 1 Then
					spreadsheetControl1.OpenCellEditor(CellEditorMode.Edit)
				End If
			End If
		End Sub

		' Display a custom in-place editor (a spin editor) for a cell.
		Private Sub spreadsheetControl1_CustomCellEdit(ByVal sender As Object, ByVal e As SpreadsheetCustomCellEditEventArgs) Handles spreadsheetControl1.CustomCellEdit
			If e.ValueObject.IsText Then
				e.RepositoryItem = CreateCustomEditor(e.ValueObject.TextValue)
			End If
		End Sub

		Private Function CreateCustomEditor(ByVal tag As String) As RepositoryItem
			Select Case tag
				Case "RegularHoursWorked"
					Return CreateSpinEdit(0, 184, 1)
				Case "VacationHours"
					Return CreateSpinEdit(0, 184, 1)
				Case "SickHours"
					Return CreateSpinEdit(0, 184, 1)
				Case "OvertimeHours"
					Return CreateSpinEdit(0, 100, 1)
				Case "OvertimeRate"
					Return CreateSpinEdit(0, 50, 1)
				Case "OtherDeduction"
					Return CreateSpinEdit(0, 100, 1)
				Case Else
					Return Nothing
			End Select
		End Function

		private RepositoryItemSpinEdit CreateSpinEditFunction(minValue As Integer, maxValue As Integer, increment As Integer) New RepositoryItemSpinEdit With {.AutoHeight = False, .BorderStyle = BorderStyles.NoBorder, .MinValue = minValue, .MaxValue = maxValue, .Increment = increment, .IsFloatValue = False}

		' Suppress the protection warning message.
		private void spreadsheetControl1_ProtectionWarningSub(sender As Object, e As HandledEventArgs) e.Handled = True


		' Generate sample data for a payroll.
		private void InitializePayrollData()
			payrollData.Add(New PayrollModel() With {.EmployeeName = "Linda Brown", .HourlyWage = 10.0, .RegularHoursWorked = 40, .VacationHours = 5, .SickHours = 1, .OvertimeHours = 0, .OvertimeRate = 15.0, .OtherDeduction = 20.0, .TaxStatus = 1, .FederalAllowance = 4, .StateTax = 0.023, .FederalIncomeTax = 0.28, .SocialSecurityTax = 0.063, .MedicareTax = 0.0145, .InsuranceDeduction = 20.0, .OtherRegularDeduction = 40.0})

			payrollData.Add(New PayrollModel() With {.EmployeeName = "Kate Smith", .HourlyWage = 11.0, .RegularHoursWorked = 45, .VacationHours = 5, .SickHours = 0, .OvertimeHours = 3, .OvertimeRate = 20.0, .OtherDeduction = 20.0, .TaxStatus = 1, .FederalAllowance = 4, .StateTax = 0.0245, .FederalIncomeTax = 0.276, .SocialSecurityTax = 0.061, .MedicareTax = 0.015, .InsuranceDeduction = 20.0, .OtherRegularDeduction = 42.0})

			payrollData.Add(New PayrollModel() With {.EmployeeName = "Nick Taylor", .HourlyWage = 15.0, .RegularHoursWorked = 40, .VacationHours = 6, .SickHours = 2, .OvertimeHours = 6, .OvertimeRate = 40.0, .OtherDeduction = 21.0, .TaxStatus = 2, .FederalAllowance = 3, .StateTax = 0.0301, .FederalIncomeTax = 0.2702, .SocialSecurityTax = 0.068, .MedicareTax = 0.015, .InsuranceDeduction = 22.0, .OtherRegularDeduction = 39.0})

			payrollData.Add(New PayrollModel() With {.EmployeeName = "Tommy Dickson", .HourlyWage = 20.0, .RegularHoursWorked = 40, .VacationHours = 0, .SickHours = 0, .OvertimeHours = 3, .OvertimeRate = 45.0, .OtherDeduction = 12.46, .TaxStatus = 3, .FederalAllowance = 4, .StateTax = 0.045, .FederalIncomeTax = 0.2904, .SocialSecurityTax = 0.084, .MedicareTax = 0.0143, .InsuranceDeduction = 41.4, .OtherRegularDeduction = 24.3})

			payrollData.Add(New PayrollModel() With {.EmployeeName = "Emmy Milton", .HourlyWage = 32.0, .RegularHoursWorked = 45, .VacationHours = 0, .SickHours = 0, .OvertimeHours = 5, .OvertimeRate = 40.0, .OtherDeduction = 0.0, .TaxStatus = 2, .FederalAllowance = 3, .StateTax = 0.025, .FederalIncomeTax = 0.28, .SocialSecurityTax = 0.064, .MedicareTax = 0.0143, .InsuranceDeduction = 19.34, .OtherRegularDeduction = 25.0})

		private void BindDataSource()
			spreadsheetBindingManager1.SheetName = "Payroll Calculator"

			' Bind data source properties to spreadsheet cells. 
			spreadsheetBindingManager1.AddBinding("EmployeeName", "C3")
			spreadsheetBindingManager1.AddBinding("HourlyWage", "D6")
			spreadsheetBindingManager1.AddBinding("RegularHoursWorked", "D8")
			spreadsheetBindingManager1.AddBinding("VacationHours", "D10")
			spreadsheetBindingManager1.AddBinding("SickHours", "D12")
			spreadsheetBindingManager1.AddBinding("OvertimeHours", "D14")
			spreadsheetBindingManager1.AddBinding("OvertimeRate", "D16")
			spreadsheetBindingManager1.AddBinding("OtherDeduction", "D22")
			spreadsheetBindingManager1.AddBinding("TaxStatus", "I4")
			spreadsheetBindingManager1.AddBinding("FederalAllowance", "I6")
			spreadsheetBindingManager1.AddBinding("StateTax", "I8")
			spreadsheetBindingManager1.AddBinding("FederalIncomeTax", "I10")
			spreadsheetBindingManager1.AddBinding("SocialSecurityTax", "I12")
			spreadsheetBindingManager1.AddBinding("MedicareTax", "I14")
			spreadsheetBindingManager1.AddBinding("InsuranceDeduction", "I20")
			spreadsheetBindingManager1.AddBinding("OtherRegularDeduction", "I22")

			' Bind the list of PayrollModel objects to the BindingSource.  
			bindingSource1.DataSource = payrollData
			' Attach the BindingSource to the SpreadsheetBindingManager and DataNavigator control.  
			spreadsheetBindingManager1.DataSource = bindingSource1
			dataNavigator1.DataSource = bindingSource1
	End Class
End Namespace
