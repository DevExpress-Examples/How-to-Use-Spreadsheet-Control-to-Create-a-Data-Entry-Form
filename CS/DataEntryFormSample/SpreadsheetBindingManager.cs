using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataEntryFormSample {
    /// <summary>
    /// This class is used to bind the data source object's properties to spreadsheet cells. 
    /// </summary>
    public partial class SpreadsheetBindingManager : Component {
        private SpreadsheetControl control;
        private object dataSource;
        private object currentItem;
        private BindingManagerBase bindingManager;
        private readonly Dictionary<string, string> cellBindings = new Dictionary<string, string>();
        private readonly PropertyDescriptorCollection propertyDescriptors = new PropertyDescriptorCollection(null);

        public SpreadsheetBindingManager() {
            InitializeComponent();
        }

        public SpreadsheetBindingManager(IContainer container) {
            container.Add(this);
            InitializeComponent();
        }

        public SpreadsheetControl Control {
            get => control;
            set {
                if (!ReferenceEquals(control, value)) {
                    if (control != null)
                        control.CellValueChanged -= SpreadsheetControl_CellValueChanged;
                    control = value;
                    if (control != null)
                        control.CellValueChanged += SpreadsheetControl_CellValueChanged;
                }
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string SheetName { get; set; }

        public object DataSource {
            get => dataSource;
            set {
                if (!ReferenceEquals(dataSource, value)) {
                    Detach();
                    dataSource = value;
                    Attach();
                }
            }
        }

        /// <summary>
        /// Creates a binding between the data source object's property and a cell.
        /// </summary>
        /// <param name="propertyName">Specifies the data source property name.</param>
        /// <param name="cellReference">Specifies a cell reference in the A1 reference style.</param>
        public void AddBinding(string propertyName, string cellReference) {
            if (cellBindings.ContainsKey(propertyName))
                throw new ArgumentException($"A binding to the {propertyName} property already exists");
            if (dataSource is ITypedList typedList) {
                PropertyDescriptorCollection dataSourceProperties = typedList.GetItemProperties(null);
                PropertyDescriptor propertyDescriptor = dataSourceProperties[propertyName];
                if (propertyDescriptor == null)
                    throw new InvalidOperationException($"Unknown { propertyName } property");
                if (currentItem != null)
                    propertyDescriptor.AddValueChanged(currentItem, OnPropertyChanged);
                propertyDescriptors.Add(propertyDescriptor);
            }
            cellBindings.Add(propertyName, cellReference);
        }

        /// <summary>
        /// Removes a binding for a data source property.
        /// </summary>
        /// <param name="propertyName">Specifies the data source property name.</param>
        public void RemoveBinding(string propertyName) {
            if (cellBindings.ContainsKey(propertyName)) {
                PropertyDescriptor propertyDescriptor = propertyDescriptors[propertyName];
                if (currentItem != null)
                    propertyDescriptor.RemoveValueChanged(currentItem, OnPropertyChanged);
                propertyDescriptors.Remove(propertyDescriptor);
                cellBindings.Remove(propertyName);
            }
        }

        /// <summary>
        /// Removes all bindings.
        /// </summary>
        public void ClearBindings() {
            UnsubscribePropertyChanged();
            propertyDescriptors.Clear();
            cellBindings.Clear();
        }

        /// <summary>
        /// Retrieves the binding manager and property descriptors, and subscribes to the data source and data members events.
        /// </summary>
        private void Attach() {
            if (dataSource is ICurrencyManagerProvider provider) {
                bindingManager = provider.CurrencyManager;
                bindingManager.CurrentChanged += BindingManager_CurrentChanged;
                currentItem = bindingManager.Current;
            }
            if (dataSource is ITypedList typedList) {
                PropertyDescriptorCollection dataSourceProperties = typedList.GetItemProperties(null);
                foreach (string propertyName in cellBindings.Keys) {
                    PropertyDescriptor propertyDescriptor = dataSourceProperties[propertyName];
                    if (propertyDescriptor == null)
                        throw new InvalidOperationException($"Unable to get a property descriptor for the { propertyName } property");
                    propertyDescriptors.Add(propertyDescriptor);
                }
            }
            PullData();
            SubscribePropertyChanged();
        }

        /// <summary>
        /// Unsubscribes from the data source and data members events, and clears property descriptors. 
        /// </summary>
        private void Detach() {
            if (dataSource != null) {
                UnsubscribePropertyChanged();
                if (bindingManager != null) {
                    bindingManager.CurrentChanged -= BindingManager_CurrentChanged;
                    bindingManager = null;
                    currentItem = null;
                }
                propertyDescriptors.Clear();
            }
        }

        private void BindingManager_CurrentChanged(object sender, EventArgs e) {
            // Update the data entry form when the current item changes.
            DeactivateCellEditor(CellEditorEnterValueMode.ActiveCell);
            control?.BeginUpdate();
            try {
                UnsubscribePropertyChanged();
                currentItem = bindingManager.Current;
                PullData();
                SubscribePropertyChanged();
            }
            finally {
                control?.EndUpdate();
                ActivateCellEditor();
            }
        }

        private void UnsubscribePropertyChanged() {
            if (currentItem != null) {
                foreach (PropertyDescriptor propertyDescriptor in propertyDescriptors)
                    propertyDescriptor.RemoveValueChanged(currentItem, OnPropertyChanged);
            }
        }
        private void SubscribePropertyChanged() {
            if (currentItem != null) {
                foreach (PropertyDescriptor propertyDescriptor in propertyDescriptors)
                    propertyDescriptor.AddValueChanged(currentItem, OnPropertyChanged);
            }
        }

        private void OnPropertyChanged(object sender, EventArgs eventArgs) {
            // Update the bound cell's value when the corresponding data source property's value changes.
            PropertyDescriptor propertyDescriptor = sender as PropertyDescriptor;
            if (propertyDescriptor != null && currentItem != null) {
                string reference;
                if (cellBindings.TryGetValue(propertyDescriptor.Name, out reference))
                    SetCellValue(reference, CellValue.FromObject(propertyDescriptor.GetValue(currentItem)));
            }
        }

        // Pull data from the data source (update values of all bound cells).
        private void PullData() {
            if (currentItem != null) {
                foreach (PropertyDescriptor propertyDescriptor in propertyDescriptors) {
                    string reference = cellBindings[propertyDescriptor.Name];
                    SetCellValue(reference, CellValue.FromObject(propertyDescriptor.GetValue(currentItem)));
                }
            }
        }

        private void SpreadsheetControl_CellValueChanged(object sender, SpreadsheetCellEventArgs e) {
            // Update the data source property's value when the bound cell's value changes.
            if (e.SheetName == SheetName) {
                string reference = e.Cell.GetReferenceA1();
                string propertyName = cellBindings.SingleOrDefault(p => p.Value == reference).Key;
                if (!string.IsNullOrEmpty(propertyName)) {
                    PropertyDescriptor propertyDescriptor = propertyDescriptors[propertyName];
                    if (propertyDescriptor != null && currentItem != null)
                        propertyDescriptor.SetValue(currentItem, e.Value.ToObject());
                }
            }
        }

        private Worksheet Sheet => 
            control != null && control.Document.Worksheets.Contains(SheetName) ? control.Document.Worksheets[SheetName] : null;

        private void SetCellValue(string reference, CellValue value) {
            if (Sheet != null) {
                if (reference == Sheet.Selection.GetReferenceA1())
                    DeactivateCellEditor();
                Sheet[reference].Value = value;
                if (reference == Sheet.Selection.GetReferenceA1())
                    ActivateCellEditor();
            }
        }

        private void ActivateCellEditor() {
            var sheet = Sheet;
            if (sheet != null) {
                var editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection);
                if (editors.Count == 1)
                    control.OpenCellEditor(CellEditorMode.Edit);
            }
        }

        private void DeactivateCellEditor(CellEditorEnterValueMode mode = CellEditorEnterValueMode.Cancel) {
            if (control != null && control.IsCellEditorActive)
                control.CloseCellEditor(mode);
        }
    }
}
