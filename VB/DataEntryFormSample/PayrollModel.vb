Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading.Tasks

Namespace DataEntryFormSample
	''' <summary>
	''' An entity class that exposes basic properties for a payroll. 
	''' </summary>
	Public Class PayrollModel
		Implements INotifyPropertyChanged

'INSTANT VB NOTE: The variable employeeName was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private employeeName_Renamed As String
'INSTANT VB NOTE: The variable hourlyWage was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private hourlyWage_Renamed As Double
'INSTANT VB NOTE: The variable regularHoursWorked was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private regularHoursWorked_Renamed As Double
'INSTANT VB NOTE: The variable vacationHours was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private vacationHours_Renamed As Double
'INSTANT VB NOTE: The variable sickHours was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private sickHours_Renamed As Double
'INSTANT VB NOTE: The variable overtimeHours was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private overtimeHours_Renamed As Double
'INSTANT VB NOTE: The variable overtimeRate was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private overtimeRate_Renamed As Double
'INSTANT VB NOTE: The variable otherDeduction was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private otherDeduction_Renamed As Double
'INSTANT VB NOTE: The variable taxStatus was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private taxStatus_Renamed As Integer
'INSTANT VB NOTE: The variable federalAllowance was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private federalAllowance_Renamed As Integer
'INSTANT VB NOTE: The variable stateTax was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private stateTax_Renamed As Double
'INSTANT VB NOTE: The variable federalIncomeTax was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private federalIncomeTax_Renamed As Double
'INSTANT VB NOTE: The variable socialSecurityTax was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private socialSecurityTax_Renamed As Double
'INSTANT VB NOTE: The variable medicareTax was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private medicareTax_Renamed As Double
'INSTANT VB NOTE: The variable insuranceDeduction was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private insuranceDeduction_Renamed As Double
'INSTANT VB NOTE: The variable otherRegularDeduction was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private otherRegularDeduction_Renamed As Double

		Public Property EmployeeName() As String
			Function(get) employeeName_Renamed
			Set(ByVal value As String)
				If employeeName_Renamed <> value Then
					employeeName_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property HourlyWage() As Double
			Function(get) hourlyWage_Renamed
			Set(ByVal value As Double)
				If hourlyWage_Renamed <> value Then
					hourlyWage_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property RegularHoursWorked() As Double
			Function(get) regularHoursWorked_Renamed
			Set(ByVal value As Double)
				If regularHoursWorked_Renamed <> value Then
					regularHoursWorked_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property VacationHours() As Double
			Function(get) vacationHours_Renamed
			Set(ByVal value As Double)
				If vacationHours_Renamed <> value Then
					vacationHours_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property SickHours() As Double
			Function(get) sickHours_Renamed
			Set(ByVal value As Double)
				If sickHours_Renamed <> value Then
					sickHours_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property OvertimeHours() As Double
			Function(get) overtimeHours_Renamed
			Set(ByVal value As Double)
				If overtimeHours_Renamed <> value Then
					overtimeHours_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property OvertimeRate() As Double
			Function(get) overtimeRate_Renamed
			Set(ByVal value As Double)
				If overtimeRate_Renamed <> value Then
					overtimeRate_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property OtherDeduction() As Double
			Function(get) otherDeduction_Renamed
			Set(ByVal value As Double)
				If otherDeduction_Renamed <> value Then
					otherDeduction_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property TaxStatus() As Integer
			Function(get) taxStatus_Renamed
			Set(ByVal value As Integer)
				If taxStatus_Renamed <> value Then
					taxStatus_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property FederalAllowance() As Integer
			Function(get) federalAllowance_Renamed
			Set(ByVal value As Integer)
				If federalAllowance_Renamed <> value Then
					federalAllowance_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property StateTax() As Double
			Function(get) stateTax_Renamed
			Set(ByVal value As Double)
				If stateTax_Renamed <> value Then
					stateTax_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property FederalIncomeTax() As Double
			Function(get) federalIncomeTax_Renamed
			Set(ByVal value As Double)
				If federalIncomeTax_Renamed <> value Then
					federalIncomeTax_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property SocialSecurityTax() As Double
			Function(get) socialSecurityTax_Renamed
			Set(ByVal value As Double)
				If socialSecurityTax_Renamed <> value Then
					socialSecurityTax_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property MedicareTax() As Double
			Function(get) medicareTax_Renamed
			Set(ByVal value As Double)
				If medicareTax_Renamed <> value Then
					medicareTax_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property InsuranceDeduction() As Double
			Function(get) insuranceDeduction_Renamed
			Set(ByVal value As Double)
				If insuranceDeduction_Renamed <> value Then
					insuranceDeduction_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Property OtherRegularDeduction() As Double
			Function(get) otherRegularDeduction_Renamed
			Set(ByVal value As Double)
				If otherRegularDeduction_Renamed <> value Then
					otherRegularDeduction_Renamed = value
					NotifyPropertyChanged()
				End If
			End Set
		End Property

		Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

		' This method is called by the Set accessor of each property.  
		' The CallerMemberName attribute applied to the optional propertyName parameter  
		' causes the property name of the caller to be substituted as an argument.
		Private Sub NotifyPropertyChanged(Optional <CallerMemberName> ByVal propertyName As String = "")
			PropertyChanged?.Invoke(Me, New PropertyChangedEventArgs(propertyName))
		End Sub
	End Class
End Namespace
