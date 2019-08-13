using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraSpreadsheet;

namespace DataEntryFormSample {
    public partial class MainForm : XtraForm {
        readonly List<PayrollModel> payrollData = new List<PayrollModel>();
        // ...

        public MainForm() {
            InitializeComponent();
            InitializePayrollData();
            LoadDocumentTemplate();
            BindCustomEditors();
            BindDataSource();
        }

        private void LoadDocumentTemplate() {
            spreadsheetControl1.LoadDocument("PayrollCalculatorTemplate.xlsx");
            spreadsheetControl1.Document.History.IsEnabled = false;
        }

        private void BindCustomEditors() {
            var sheet = spreadsheetControl1.ActiveWorksheet;
            // Assign custom editors to worksheet cells.
            sheet.CustomCellInplaceEditors.Add(sheet["D8"], CustomCellInplaceEditorType.Custom, "RegularHoursWorked");
            sheet.CustomCellInplaceEditors.Add(sheet["D10"], CustomCellInplaceEditorType.Custom, "VacationHours");
            sheet.CustomCellInplaceEditors.Add(sheet["D12"], CustomCellInplaceEditorType.Custom, "SickHours");
            sheet.CustomCellInplaceEditors.Add(sheet["D14"], CustomCellInplaceEditorType.Custom, "OvertimeHours");
            sheet.CustomCellInplaceEditors.Add(sheet["D16"], CustomCellInplaceEditorType.Custom, "OvertimeRate");
            sheet.CustomCellInplaceEditors.Add(sheet["D22"], CustomCellInplaceEditorType.Custom, "OtherDeduction");
        }

        // Activate a cell editor when a user selects an editable cell.
        private void SpreadsheetControl1_SelectionChanged(object sender, EventArgs e) {
            var sheet = spreadsheetControl1.ActiveWorksheet;
            if (sheet != null) {
                var editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection);
                if (editors.Count == 1)
                    spreadsheetControl1.OpenCellEditor(CellEditorMode.Edit);
            }
        }

        // Display a custom in-place editor (a spin editor) for a cell.
        private void spreadsheetControl1_CustomCellEdit(object sender, SpreadsheetCustomCellEditEventArgs e)
        {
            if (e.ValueObject.IsText)
                e.RepositoryItem = CreateCustomEditor(e.ValueObject.TextValue);
        }

        private RepositoryItem CreateCustomEditor(string tag)
        {
            switch (tag)
            {
                case "RegularHoursWorked": return CreateSpinEdit(0, 184, 1);
                case "VacationHours": return CreateSpinEdit(0, 184, 1);
                case "SickHours": return CreateSpinEdit(0, 184, 1);
                case "OvertimeHours": return CreateSpinEdit(0, 100, 1);
                case "OvertimeRate": return CreateSpinEdit(0, 50, 1);
                case "OtherDeduction": return CreateSpinEdit(0, 100, 1);
                default: return null;
            }
        }

        private RepositoryItemSpinEdit CreateSpinEdit(int minValue, int maxValue, int increment) => new RepositoryItemSpinEdit
        {
            AutoHeight = false,
            BorderStyle = BorderStyles.NoBorder,
            MinValue = minValue,
            MaxValue = maxValue,
            Increment = increment,
            IsFloatValue = false
        };

        // Suppress the protection warning message.
        private void spreadsheetControl1_ProtectionWarning(object sender, HandledEventArgs e) => e.Handled = true;

        
        // Generate sample data for a payroll.
        private void InitializePayrollData() {
            payrollData.Add(new PayrollModel() {
                EmployeeName = "Linda Brown",
                HourlyWage = 10.0,
                RegularHoursWorked = 40,
                VacationHours = 5,
                SickHours = 1,
                OvertimeHours = 0,
                OvertimeRate = 15.0,
                OtherDeduction = 20.0,
                TaxStatus = 1,
                FederalAllowance = 4,
                StateTax = 0.023,
                FederalIncomeTax = 0.28,
                SocialSecurityTax = 0.063,
                MedicareTax = 0.0145,
                InsuranceDeduction = 20.0,
                OtherRegularDeduction = 40.0
            });

            payrollData.Add(new PayrollModel() {
                EmployeeName = "Kate Smith",
                HourlyWage = 11.0,
                RegularHoursWorked = 45,
                VacationHours = 5,
                SickHours = 0,
                OvertimeHours = 3,
                OvertimeRate = 20.0,
                OtherDeduction = 20.0,
                TaxStatus = 1,
                FederalAllowance = 4,
                StateTax = 0.0245,
                FederalIncomeTax = 0.276,
                SocialSecurityTax = 0.061,
                MedicareTax = 0.015,
                InsuranceDeduction = 20.0,
                OtherRegularDeduction = 42.0
            });

            payrollData.Add(new PayrollModel() {
                EmployeeName = "Nick Taylor",
                HourlyWage = 15.0,
                RegularHoursWorked = 40,
                VacationHours = 6,
                SickHours = 2,
                OvertimeHours = 6,
                OvertimeRate = 40.0,
                OtherDeduction = 21.0,
                TaxStatus = 2,
                FederalAllowance = 3,
                StateTax = 0.0301,
                FederalIncomeTax = 0.2702,
                SocialSecurityTax = 0.068,
                MedicareTax = 0.015,
                InsuranceDeduction = 22.0,
                OtherRegularDeduction = 39.0
            });

            payrollData.Add(new PayrollModel() {
                EmployeeName = "Tommy Dickson",
                HourlyWage = 20.0,
                RegularHoursWorked = 40,
                VacationHours = 0,
                SickHours = 0,
                OvertimeHours = 3,
                OvertimeRate = 45.0,
                OtherDeduction = 12.46,
                TaxStatus = 3,
                FederalAllowance = 4,
                StateTax = 0.045,
                FederalIncomeTax = 0.2904,
                SocialSecurityTax = 0.084,
                MedicareTax = 0.0143,
                InsuranceDeduction = 41.4,
                OtherRegularDeduction = 24.3
            });

            payrollData.Add(new PayrollModel() {
                EmployeeName = "Emmy Milton",
                HourlyWage = 32.0,
                RegularHoursWorked = 45,
                VacationHours = 0,
                SickHours = 0,
                OvertimeHours = 5,
                OvertimeRate = 40.0,
                OtherDeduction = 0.0,
                TaxStatus = 2,
                FederalAllowance = 3,
                StateTax = 0.025,
                FederalIncomeTax = 0.28,
                SocialSecurityTax = 0.064,
                MedicareTax = 0.0143,
                InsuranceDeduction = 19.34,
                OtherRegularDeduction = 25.0
            });
        }

        private void BindDataSource() {
            spreadsheetBindingManager1.SheetName = "Payroll Calculator";

            // Bind data source properties to spreadsheet cells. 
            spreadsheetBindingManager1.AddBinding("EmployeeName", "C3");
            spreadsheetBindingManager1.AddBinding("HourlyWage", "D6");
            spreadsheetBindingManager1.AddBinding("RegularHoursWorked", "D8");
            spreadsheetBindingManager1.AddBinding("VacationHours", "D10");
            spreadsheetBindingManager1.AddBinding("SickHours", "D12");
            spreadsheetBindingManager1.AddBinding("OvertimeHours", "D14");
            spreadsheetBindingManager1.AddBinding("OvertimeRate", "D16");
            spreadsheetBindingManager1.AddBinding("OtherDeduction", "D22");
            spreadsheetBindingManager1.AddBinding("TaxStatus", "I4");
            spreadsheetBindingManager1.AddBinding("FederalAllowance", "I6");
            spreadsheetBindingManager1.AddBinding("StateTax", "I8");
            spreadsheetBindingManager1.AddBinding("FederalIncomeTax", "I10");
            spreadsheetBindingManager1.AddBinding("SocialSecurityTax", "I12");
            spreadsheetBindingManager1.AddBinding("MedicareTax", "I14");
            spreadsheetBindingManager1.AddBinding("InsuranceDeduction", "I20");
            spreadsheetBindingManager1.AddBinding("OtherRegularDeduction", "I22");

            // Bind the list of PayrollModel objects to the BindingSource.  
            bindingSource1.DataSource = payrollData;
            // Attach the BindingSource to the SpreadsheetBindingManager and DataNavigator control.  
            spreadsheetBindingManager1.DataSource = bindingSource1;
            dataNavigator1.DataSource = bindingSource1;
        }
    }
}
