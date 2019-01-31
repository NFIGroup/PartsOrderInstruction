using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;

namespace Parts_Order_Instruction
{
    [AddIn("Report Command AddIn", Version = "1.0.0.0")]
    public class ReportCommandAddIn : IReportCommand2
    {
        #region IReportCommand Members
        static public IGlobalContext _globalContext;
        private IRecordContext _recordContext;
        public static IIncident _incidentRecord;
        IList<IReportRow> _selectedRows;
        public static ProgressForm form = new ProgressForm();
        string _srNumber = "";
        string _shippingInstruction = "";
        int _shipToSite = 0;
        int _qtyToBeCompleted;
        int _partsOrderInstructionID;
        /// <summary>
        /// 
        /// </summary>
        public bool Enabled(IList<IReportRow> rows)
        {
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Execute(IList<IReportRow> rows)
        {
            _selectedRows = rows;
            _recordContext = _globalContext.AutomationContext.CurrentWorkspace;

            _incidentRecord = (IIncident)_recordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
            System.Threading.Thread th = new System.Threading.Thread(ProcessSelectedRowInfo);
            th.Start();
            form.Show();
        }

        /// <summary>
        /// Function to populate Report Incident Fields based on Report Data
        /// </summary>
        public void ProcessSelectedRowInfo()
        {
            try
            {
                foreach (IReportRow row in _selectedRows)
                {
                    IList<IReportCell> cells = row.Cells;
                    foreach (IReportCell cell in cells)
                    {
                        if (cell.Name == "SR_NUM")
                        {
                            _srNumber = cell.Value;
                        }
                        if (cell.Name == "Ship_Site_ID")
                        {
                            _shipToSite = Convert.ToInt32(cell.Value);
                        }
                        if (cell.Name == "Qty_Per_Week")
                        {
                            _qtyToBeCompleted = Convert.ToInt32(cell.Value);
                        }
                        if (cell.Name == "PartOrderInstruction_ID")
                        {
                            _partsOrderInstructionID = Convert.ToInt32(cell.Value);
                        }
                        if (cell.Name == "Shipping_Instruction")
                        {
                            _shippingInstruction = cell.Value;
                        }
                    }
                }
                if (_srNumber != string.Empty)
                {
                    //Replace string 'SR-' wiht '00' to populate project number field
                    string projectNumber = _srNumber.Replace("SR-", "").PadLeft(5,'0');

                    IGenericObject its = _recordContext.GetWorkspaceRecord("CO$ITS") as IGenericObject;
                    string itsNUm = GetCustomObjectFieldValue("ITS_nmbr", its);

                    //Set Incident Custom Fields
                    SetIncidentField("CO", "project_number", projectNumber);
                    if (_shipToSite != 0)
                    {
                        SetIncidentField("CO", "Ship_to_site", _shipToSite.ToString());
                        SetIncidentField("CO", "Bill_to_site", _shipToSite.ToString());
                    }
                    SetIncidentField("CO", "no_of_vins", _qtyToBeCompleted.ToString());
                    SetIncidentField("CO", "retrofit_number", itsNUm+"-"+_srNumber);
                    SetIncidentField("CO", "PO_Number", itsNUm + "-" + _srNumber+"-"+ _qtyToBeCompleted.ToString());
                    SetIncidentField("CO", "shipping_instructions", _shippingInstruction);
                    SetIncidentField("CO", "PartOrderInstruction_ID", _partsOrderInstructionID.ToString());
                }

                _recordContext.RefreshWorkspace();
                form.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception: " + ex.Message);
                form.Hide();
            }
        }            
        /// <summary>
        /// Method which is called to get value of a custom field of Incident record.
        /// </summary>
        /// <param name="packageName">The name of the package.</param>
        /// <param name="fieldName">The name of the custom field.</param>
        /// <returns>Value of the field</returns>
        public string getIncidentField(string packageName, string fieldName, IIncident incidentRecord)
        {
            string value = "";
            IList<ICustomAttribute> incCustomAttributes = incidentRecord.CustomAttributes;

            foreach (ICustomAttribute val in incCustomAttributes)
            {
                if (val.PackageName == packageName)//if package name matches
                {
                    if (val.GenericField.Name == packageName + "$" + fieldName)//if field matches
                    {
                        if (val.GenericField.DataValue.Value != null)
                        {
                            value = val.GenericField.DataValue.Value.ToString();
                            break;
                        }
                    }
                }
            }
            return value;
        }

        /// <summary>
        /// Method which is use to set incident field 
        /// </summary>
        /// <param name="pkgName">package name of custom field</param>
        /// <param name="fieldName">field name</param>
        /// <param name="value">value of field</param>
        public static void SetIncidentField(string pkgName, string fieldName, string value)
        {
            IList<ICustomAttribute> incCustomAttributes = _incidentRecord.CustomAttributes;

            foreach (ICustomAttribute val in incCustomAttributes)
            {
                if (val.PackageName == pkgName)
                {
                    if (val.GenericField.Name == pkgName + "$" + fieldName)
                    {
                        switch (val.GenericField.DataType)
                        {
                            case RightNow.AddIns.Common.DataTypeEnum.BOOLEAN:
                                if (value == "1" || value.ToLower() == "true")
                                {
                                    val.GenericField.DataValue.Value = true;
                                }
                                else if (value == "0" || value.ToLower() == "false")
                                {
                                    val.GenericField.DataValue.Value = false;
                                }
                                break;
                            case RightNow.AddIns.Common.DataTypeEnum.INTEGER:
                                if (value.Trim() == "" || value.Trim() == null)
                                {
                                    val.GenericField.DataValue.Value = null;
                                }
                                else
                                {
                                    val.GenericField.DataValue.Value = Convert.ToInt32(value);
                                }
                                break;
                            case RightNow.AddIns.Common.DataTypeEnum.STRING:
                                val.GenericField.DataValue.Value = value;
                                break;
                            case RightNow.AddIns.Common.DataTypeEnum.ID:
                                val.GenericField.DataValue.Value = value;
                                break;
                        }
                    }
                }
            }
            return;
        }
        /// <summary>
        /// Method which is called to get value of a Custom object.
        /// </summary>
        /// <param name="fieldName">The name of the custom field.</param>
        /// <returns>Value of the field</returns>
        public string GetCustomObjectFieldValue(string fieldName, IGenericObject genericObject)
        {
            IList<IGenericField> fields = genericObject.GenericFields;
            if (null != fields)
            {
                foreach (IGenericField field in fields)
                {
                    if (field.Name.Equals(fieldName))
                    {
                        if (field.DataValue.Value != null)
                            return field.DataValue.Value.ToString();
                    }
                }
            }
            return "";
        }
        /// <summary>
        /// 
        /// </summary>
        public Image Image16
        {
            get
            {
                return Properties.Resources.AddIn16;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public Image Image32
        {
            get
            {
                return Properties.Resources.AddIn32;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public IList<RightNow.AddIns.Common.ReportRecordIdType> RecordTypes
        {
            get
            {
                IList<ReportRecordIdType> typeList = new List<ReportRecordIdType>();

                typeList.Add(ReportRecordIdType.Answer);
                typeList.Add(ReportRecordIdType.Chat);
                typeList.Add(ReportRecordIdType.CloudAcct2Search);
                typeList.Add(ReportRecordIdType.Contact);
                typeList.Add(ReportRecordIdType.ContactList);
                typeList.Add(ReportRecordIdType.Document);
                typeList.Add(ReportRecordIdType.Flow);
                typeList.Add(ReportRecordIdType.Incident);
                typeList.Add(ReportRecordIdType.Mailing);
                typeList.Add(ReportRecordIdType.MetaAnswer);
                typeList.Add(ReportRecordIdType.Opportunity);
                typeList.Add(ReportRecordIdType.Organization);
                typeList.Add(ReportRecordIdType.Question);
                typeList.Add(ReportRecordIdType.QueuedReport);
                typeList.Add(ReportRecordIdType.Quote);
                typeList.Add(ReportRecordIdType.QuoteProduct);
                typeList.Add(ReportRecordIdType.Report);
                typeList.Add(ReportRecordIdType.Segment);
                typeList.Add(ReportRecordIdType.Survey);
                typeList.Add(ReportRecordIdType.Task);
                typeList.Add(ReportRecordIdType.CustomObjectAll);
                return typeList;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Text
        {
            get
            {
                return "Select";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "Order Parts as per Instructions";
            }
        }

        public IList<string> CustomObjectRecordTypes
        {
            get
            {
                IList<string> typeList = new List<string>();

                typeList.Add("PartOdrInstruction");                
                return typeList;
            }

        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            _globalContext = GlobalContext;
            return true;
        }

        #endregion
    }
}
