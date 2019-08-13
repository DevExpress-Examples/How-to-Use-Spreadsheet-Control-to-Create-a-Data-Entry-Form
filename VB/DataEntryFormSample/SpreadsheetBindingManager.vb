Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace DataEntryFormSample
	''' <summary>
	''' This class is used to bind the data source object's properties to spreadsheet cells. 
	''' </summary>
	Partial Public Class SpreadsheetBindingManager
		Inherits Component

'INSTANT VB NOTE: The variable control was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private control_Renamed As SpreadsheetControl
'INSTANT VB NOTE: The variable dataSource was renamed since Visual Basic does not allow variables and other class members to have the same name:
		Private dataSource_Renamed As Object
		Private currentItem As Object
		Private bindingManager As BindingManagerBase
		Private ReadOnly cellBindings As New Dictionary(Of String, String)()
		Private ReadOnly propertyDescriptors As New PropertyDescriptorCollection(Nothing)

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(ByVal container As IContainer)
			container.Add(Me)
			InitializeComponent()
		End Sub

		Public Property Control() As SpreadsheetControl
			Function(get) control_Renamed
			Set(ByVal value As SpreadsheetControl)
				If Not ReferenceEquals(control_Renamed, value) Then
					If control_Renamed IsNot Nothing Then
						RemoveHandler control_Renamed.CellValueChanged, AddressOf SpreadsheetControl_CellValueChanged
					End If
					control_Renamed = value
					If control_Renamed IsNot Nothing Then
						AddHandler control_Renamed.CellValueChanged, AddressOf SpreadsheetControl_CellValueChanged
					End If
				End If
			End Set
		End Property

		<Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
		Public Property SheetName() As String

		Public Property DataSource() As Object
			Function(get) dataSource_Renamed
			Set(ByVal value As Object)
				If Not ReferenceEquals(dataSource_Renamed, value) Then
					Detach()
					dataSource_Renamed = value
					Attach()
				End If
			End Set
		End Property

		''' <summary>
		''' Creates a binding between the data source object's property and a cell.
		''' </summary>
		''' <param name="propertyName">Specifies the data source property name.</param>
		''' <param name="cellReference">Specifies a cell reference in the A1 reference style.</param>
		Public Sub AddBinding(ByVal propertyName As String, ByVal cellReference As String)
			If cellBindings.ContainsKey(propertyName) Then
				Throw New ArgumentException($"A binding to the {propertyName} property already exists")
			End If
			If TypeOf dataSource_Renamed Is ITypedList typedList Then
				Dim dataSourceProperties As PropertyDescriptorCollection = typedList.GetItemProperties(Nothing)
				Dim propertyDescriptor As PropertyDescriptor = dataSourceProperties(propertyName)
				If propertyDescriptor Is Nothing Then
					Throw New InvalidOperationException($"Unknown { propertyName } property")
				End If
				If currentItem IsNot Nothing Then
					propertyDescriptor.AddValueChanged(currentItem, AddressOf OnPropertyChanged)
				End If
				propertyDescriptors.Add(propertyDescriptor)
			End If
			cellBindings.Add(propertyName, cellReference)
		End Sub

		''' <summary>
		''' Removes a binding for a data source property.
		''' </summary>
		''' <param name="propertyName">Specifies the data source property name.</param>
		Public Sub RemoveBinding(ByVal propertyName As String)
			If cellBindings.ContainsKey(propertyName) Then
				Dim propertyDescriptor As PropertyDescriptor = propertyDescriptors(propertyName)
				If currentItem IsNot Nothing Then
					propertyDescriptor.RemoveValueChanged(currentItem, AddressOf OnPropertyChanged)
				End If
				propertyDescriptors.Remove(propertyDescriptor)
				cellBindings.Remove(propertyName)
			End If
		End Sub

		''' <summary>
		''' Removes all bindings.
		''' </summary>
		Public Sub ClearBindings()
			UnsubscribePropertyChanged()
			propertyDescriptors.Clear()
			cellBindings.Clear()
		End Sub

		''' <summary>
		''' Retrieves the binding manager and property descriptors, and subscribes to the data source and data members events.
		''' </summary>
		Private Sub Attach()
			If TypeOf dataSource_Renamed Is ICurrencyManagerProvider provider Then
				bindingManager = provider.CurrencyManager
				AddHandler bindingManager.CurrentChanged, AddressOf BindingManager_CurrentChanged
				currentItem = bindingManager.Current
			End If
			If TypeOf dataSource_Renamed Is ITypedList typedList Then
				Dim dataSourceProperties As PropertyDescriptorCollection = typedList.GetItemProperties(Nothing)
				For Each propertyName As String In cellBindings.Keys
					Dim propertyDescriptor As PropertyDescriptor = dataSourceProperties(propertyName)
					If propertyDescriptor Is Nothing Then
						Throw New InvalidOperationException($"Unable to get a property descriptor for the { propertyName } property")
					End If
					propertyDescriptors.Add(propertyDescriptor)
				Next propertyName
			End If
			PullData()
			SubscribePropertyChanged()
		End Sub

		''' <summary>
		''' Unsubscribes from the data source and data members events, and clears property descriptors. 
		''' </summary>
		Private Sub Detach()
			If dataSource_Renamed IsNot Nothing Then
				UnsubscribePropertyChanged()
				If bindingManager IsNot Nothing Then
					RemoveHandler bindingManager.CurrentChanged, AddressOf BindingManager_CurrentChanged
					bindingManager = Nothing
					currentItem = Nothing
				End If
				propertyDescriptors.Clear()
			End If
		End Sub

		Private Sub BindingManager_CurrentChanged(ByVal sender As Object, ByVal e As EventArgs)
			' Update the data entry form when the current item changes.
			DeactivateCellEditor(CellEditorEnterValueMode.ActiveCell)
			control?.BeginUpdate()
			Try
				UnsubscribePropertyChanged()
				currentItem = bindingManager.Current
				PullData()
				SubscribePropertyChanged()
			Finally
				control?.EndUpdate()
				ActivateCellEditor()
			End Try
		End Sub

		Private Sub UnsubscribePropertyChanged()
			If currentItem IsNot Nothing Then
				For Each propertyDescriptor As PropertyDescriptor In propertyDescriptors
					propertyDescriptor.RemoveValueChanged(currentItem, AddressOf OnPropertyChanged)
				Next propertyDescriptor
			End If
		End Sub
		Private Sub SubscribePropertyChanged()
			If currentItem IsNot Nothing Then
				For Each propertyDescriptor As PropertyDescriptor In propertyDescriptors
					propertyDescriptor.AddValueChanged(currentItem, AddressOf OnPropertyChanged)
				Next propertyDescriptor
			End If
		End Sub

		Private Sub OnPropertyChanged(ByVal sender As Object, ByVal eventArgs As EventArgs)
			' Update the bound cell's value when the corresponding data source property's value changes.
			Dim propertyDescriptor As PropertyDescriptor = TryCast(sender, PropertyDescriptor)
			If propertyDescriptor IsNot Nothing AndAlso currentItem IsNot Nothing Then
				Dim reference As String = Nothing
				If cellBindings.TryGetValue(propertyDescriptor.Name, reference) Then
					SetCellValue(reference, CellValue.FromObject(propertyDescriptor.GetValue(currentItem)))
				End If
			End If
		End Sub

		' Pull data from the data source (update values of all bound cells).
		Private Sub PullData()
			If currentItem IsNot Nothing Then
				For Each propertyDescriptor As PropertyDescriptor In propertyDescriptors
					Dim reference As String = cellBindings(propertyDescriptor.Name)
					SetCellValue(reference, CellValue.FromObject(propertyDescriptor.GetValue(currentItem)))
				Next propertyDescriptor
			End If
		End Sub

		Private Sub SpreadsheetControl_CellValueChanged(ByVal sender As Object, ByVal e As SpreadsheetCellEventArgs)
			' Update the data source property's value when the bound cell's value changes.
			If e.SheetName = SheetName Then
				Dim reference As String = e.Cell.GetReferenceA1()
				Dim propertyName As String = cellBindings.SingleOrDefault(Function(p) p.Value = reference).Key
				If Not String.IsNullOrEmpty(propertyName) Then
					Dim propertyDescriptor As PropertyDescriptor = propertyDescriptors(propertyName)
					If propertyDescriptor IsNot Nothing AndAlso currentItem IsNot Nothing Then
						propertyDescriptor.SetValue(currentItem, e.Value.ToObject())
					End If
				End If
			End If
		End Sub

		Private Sheet As Function(Worksheet)If(control_Renamed IsNot Nothing AndAlso control_Renamed.Document.Worksheets.Contains(SheetName), control_Renamed.Document.Worksheets(SheetName), Nothing)

		Private Sub SetCellValue(ByVal reference As String, ByVal value As CellValue)
			If Sheet IsNot Nothing Then
				If reference = Sheet.Selection.GetReferenceA1() Then
					DeactivateCellEditor()
				End If
				Sheet(reference).Value = value
				If reference = Sheet.Selection.GetReferenceA1() Then
					ActivateCellEditor()
				End If
			End If
		End Sub

		Private Sub ActivateCellEditor()
'INSTANT VB NOTE: The variable sheet was renamed since Visual Basic does not handle local variables named the same as class members well:
			Dim sheet_Renamed = Sheet
			If sheet_Renamed IsNot Nothing Then
				Dim editors = sheet_Renamed.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet_Renamed.Selection)
				If editors.Count = 1 Then
					control_Renamed.OpenCellEditor(CellEditorMode.Edit)
				End If
			End If
		End Sub

		Private Sub DeactivateCellEditor(Optional ByVal mode As CellEditorEnterValueMode = CellEditorEnterValueMode.Cancel)
			If control_Renamed IsNot Nothing AndAlso control_Renamed.IsCellEditorActive Then
				control_Renamed.CloseCellEditor(mode)
			End If
		End Sub
	End Class
End Namespace
