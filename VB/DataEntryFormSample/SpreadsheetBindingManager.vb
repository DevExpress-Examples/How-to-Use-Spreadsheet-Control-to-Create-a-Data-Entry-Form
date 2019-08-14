Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Linq
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet

Namespace DataEntryFormSample
    ' This class is used to bind the data source object's properties to spreadsheet cells. 
    Partial Public Class SpreadsheetBindingManager
        Inherits Component

        Private _control As SpreadsheetControl
        Private _dataSource As Object
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
            Get
                Return _control
            End Get
            Set(ByVal value As SpreadsheetControl)
                If Not ReferenceEquals(_control, value) Then
                    If _control IsNot Nothing Then
                        RemoveHandler _control.CellValueChanged, AddressOf SpreadsheetControl_CellValueChanged
                    End If
                    _control = value
                    If _control IsNot Nothing Then
                        AddHandler _control.CellValueChanged, AddressOf SpreadsheetControl_CellValueChanged
                    End If
                End If
            End Set
        End Property

        <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
        Public Property SheetName() As String

        Public Property DataSource() As Object
            Get
                Return _dataSource
            End Get
            Set(ByVal value As Object)
                If Not ReferenceEquals(_dataSource, value) Then
                    Detach()
                    _dataSource = value
                    Attach()
                End If
            End Set
        End Property

        ' Creates a binding between the data source object's property and a cell.
        ' propertyName specifies the data source property name.
        ' cellReference specifies a cell reference in the A1 reference style.
        Public Sub AddBinding(ByVal propertyName As String, ByVal cellReference As String)
            If cellBindings.ContainsKey(propertyName) Then
                Throw New ArgumentException($"Already has binding to {propertyName} property")
            End If
            Dim typedList As ITypedList = TryCast(_dataSource, ITypedList)
            If typedList IsNot Nothing Then
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

        'Removes a binding for a data source property.
        'propertyName specifies the data source property name.</param>
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

        'Removes all bindings.
        Public Sub ClearBindings()
            UnsubscribePropertyChanged()
            propertyDescriptors.Clear()
            cellBindings.Clear()
        End Sub

        ' Retrieves the binding manager and property descriptors, and subscribes to the data source and data members events.
        Private Sub Attach()
            Dim provider As ICurrencyManagerProvider = TryCast(_dataSource, ICurrencyManagerProvider)
            If provider IsNot Nothing Then
                bindingManager = provider.CurrencyManager
                AddHandler bindingManager.CurrentChanged, AddressOf BindingManager_CurrentChanged
                currentItem = bindingManager.Current
            End If
            Dim typedList As ITypedList = TryCast(_dataSource, ITypedList)
            If typedList IsNot Nothing Then
                Dim dataSourceProperties As PropertyDescriptorCollection = typedList.GetItemProperties(Nothing)
                For Each propertyName As String In cellBindings.Keys
                    Dim propertyDescriptor As PropertyDescriptor = dataSourceProperties(propertyName)
                    If propertyDescriptor Is Nothing Then
                        Throw New InvalidOperationException($"Unable to get property descriptor for { propertyName } property")
                    End If
                    propertyDescriptors.Add(propertyDescriptor)
                Next propertyName
            End If
            PullData()
            SubscribePropertyChanged()
        End Sub

        'Unsubscribes from the data source and data members events, and clears property descriptors. 
        Private Sub Detach()
            If _dataSource IsNot Nothing Then
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
            _control?.BeginUpdate()
            Try
                'Update the data entry form when the current item changes.
                DeactivateCellEditor(CellEditorEnterValueMode.ActiveCell)
                UnsubscribePropertyChanged()
                currentItem = bindingManager.Current
                PullData()
                SubscribePropertyChanged()
                ActivateCellEditor()
            Finally
                _control?.EndUpdate()
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

            'Update the bound cell's value when the corresponding data source property's value changes.
            Dim propertyDescriptor As PropertyDescriptor = TryCast(sender, PropertyDescriptor)
            If propertyDescriptor IsNot Nothing AndAlso currentItem IsNot Nothing Then
                Dim reference As String = Nothing
                If cellBindings.TryGetValue(propertyDescriptor.Name, reference) Then
                    SetCellValue(reference, CellValue.FromObject(propertyDescriptor.GetValue(currentItem)))
                End If
            End If
        End Sub

        'Pull data from the data source (update values of all bound cells).
        Private Sub PullData()
            If currentItem IsNot Nothing Then
                For Each propertyDescriptor As PropertyDescriptor In propertyDescriptors
                    Dim reference As String = cellBindings(propertyDescriptor.Name)
                    SetCellValue(reference, CellValue.FromObject(propertyDescriptor.GetValue(currentItem)))
                Next propertyDescriptor
            End If
        End Sub

        Private Sub SpreadsheetControl_CellValueChanged(ByVal sender As Object, ByVal e As SpreadsheetCellEventArgs)

            'Update the data source property's value when the bound cell's value changes.
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

        Private ReadOnly Property Sheet() As Worksheet
            Get
                Return If(_control IsNot Nothing AndAlso _control.Document.Worksheets.Contains(SheetName), _control.Document.Worksheets(SheetName), Nothing)
            End Get
        End Property

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
            
            Dim _sheet = Sheet
            If _sheet IsNot Nothing Then
                Dim editors = _sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(_sheet.Selection)
                If editors.Count = 1 Then
                    _control.OpenCellEditor(CellEditorMode.Edit)
                End If
            End If
        End Sub

        Private Sub DeactivateCellEditor(Optional ByVal mode As CellEditorEnterValueMode = CellEditorEnterValueMode.Cancel)
            If _control IsNot Nothing AndAlso _control.IsCellEditorActive Then
                _control.CloseCellEditor(mode)
            End If
        End Sub
    End Class
End Namespace
