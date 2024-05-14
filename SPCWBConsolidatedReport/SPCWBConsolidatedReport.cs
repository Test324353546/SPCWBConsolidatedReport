using Syncfusion.WinForms.DataGrid;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Resources;
using System.Windows.Forms;
using Syncfusion.XlsIO;
using System.Diagnostics;
//using System.Data.OleDb;
//using System.SPCWBConsolidatedReport;



namespace SPCWBConsolidatedReport
{
    public partial class SPCWBConsolidatedReport : Form
    {
        public string columntext = "";
        public static IniFile.ini iniObj = new IniFile.ini(AppDomain.CurrentDomain.BaseDirectory + "//SPCWB.ini");
        char[] c = { '\r', '\n' };
        ResourceManager rm;
        bool selectFlag = false;
        public string FilePath = "";
        public string FileName = "";
        ArrayList fileNames = new ArrayList();
        public string htmltextfromDatabase = "";
        string conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
        bool state = false;
        public int EmailAlertCount;
        OleDbCommand cmd;
        OleDbConnection connection;
        OleDbDataReader reader;
        ListViewItem lstviewItem = new ListViewItem();
        DataTable dtFillDatagridWithFiles = new DataTable();
        //GridCheckBoxColumn checkBoxColumn = new GridCheckBoxColumn();
        public string checkcol = "";
        string colName = "File Name";
        string colName1 = "File Path";
        string colName2 = "Modified Date";
        public int LastRowIndex;
        public string filepath = "";
        public int currentCellRowIndex;
        public int previousRowIndex;
        public int previousvalue;
        // public static string cellValue = "";
        // ArrayList arrFileName = new ArrayList();
        List<KeyValuePair<string, string>> arrFileName = new List<KeyValuePair<string, string>>();
        //  List<string> arrFileName = new List<string>();
        public string cssstyleforheadertag;
        //private string workbook;
        public string ProcName = "";
        public string PartName = "";
        public string PartNo = "";
        public string CharName = "";
        public string SGNO="";
        public string CharType = "";
        public int ChartType;
        public int SGSize;
        public int Target;
        public double USL;
        public double LSL;
        public int DataEntry;
        public string TraceCat = "";
        public string TraceType = "";
        public string Tracevalue = "";
        public string EventName = "";
        public string EventValue = "";
        public string EventValueForEvent = "";
        public string CharName1 = "";
        public string Operator = "";
        public string UserName = "";
        public string Email = "";
        public string EventCat = "";
        public string EventType = "";
        public int SGNo;
        OleDbConnection con = new OleDbConnection();
        public DateTime SGDate;
        public string sgvalue = "";
        public int charId;
        public string ProductCode = "";
        public string Grade = "";
        public int ValueCountInSGData;
        public int tracecatcount;
        public static string TraceCategory = "";
        public string filename = "";
       // public 
        
        //public string TraceCat = "";
        SPCAccessModule.SPCAccessDB accessObj;
        int[] ArrSubgroupNo;
        ArrayList arrlstSubgroupNo = new ArrayList();
        ArrayList arrlstActualreading = new ArrayList();
        ArrayList arrlstSGDate = new ArrayList();
        ArrayList arrlstTraceValue = new ArrayList();
        public OleDbConnection CloseConnection()
        {
            if (connection.State == System.Data.ConnectionState.Open)
            {
                connection.Close();
            }
            return connection;
        }
        public SPCWBConsolidatedReport()
        {
            InitializeComponent();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            this.sfdgSPCWBConsolidatedReport.SearchController.Search(txtSearch.Text);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            #region step-1 Create OpenFileDialog to Fill CheckListBox With selected Files in OpenFileDialog
            this.sfdgSPCWBConsolidatedReport.AutoScrollOffset = new System.Drawing.Point(100, 100);
            if (System.IO.File.Exists(AppDomain.CurrentDomain.BaseDirectory + "//SPCWB.ini") == false)//If file SPCWB.ini not present then it creates that file and path,version and startup option are added to that file 
            {
                iniObj.IniWriteValue("VersionInfo", "Version", "6");
                iniObj.IniWriteValue("Settings", "DataPath", AppDomain.CurrentDomain.BaseDirectory + @"Data");//default path
                iniObj.IniWriteValue("StartUpOption", "ShowAtStartUp", "Yes");
            }

            string DataPath = iniObj.IniReadValue("Settings", "DataPath");

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.InitialDirectory = DataPath;
            openFileDialog.Filter = "SPC WorkBench Files (*.spcx) | *.spcx";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                foreach (var filename in openFileDialog.FileNames)
                {
                    FilePath = openFileDialog.FileName;
                    FileName = filename.Remove(0, FilePath.LastIndexOf('\\') + 1);


                    filepath = Path.GetDirectoryName(FilePath);

                    //checkcol = checkBoxColumn;//
                    fileNames.Add(FileName);
                    var lastModified = System.IO.File.GetLastWriteTime(FilePath);
                    System.Data.DataRow dr1 = dtFillDatagridWithFiles.NewRow();
                    dr1[colName] = FileName;
                    dr1[colName1] = filename;
                    dr1[colName2] = lastModified.ToString();
                    #region check if FileName already exists in Datatable or not
                    System.Data.DataRow[] foundRows = dtFillDatagridWithFiles.Select("[File Name] = '" + FileName + "' ");
                    
                    if (foundRows.Length != 0)
                    {
                        MessageBox.Show("File Name already exists");
                    }
                    else
                    {
                        dtFillDatagridWithFiles.Rows.Add(dr1.ItemArray);
                        sfdgSPCWBConsolidatedReport.DataSource = dtFillDatagridWithFiles;
                        sfdgSPCWBConsolidatedReport.Columns[1].Width = 250;
                        sfdgSPCWBConsolidatedReport.AutoScrollOffset = new System.Drawing.Point(300, 300);
                    }
                    #endregion check if FileName already exists in Datatable or not

                }



            }
        }

        #endregion Create OpenFileDialog to Fill CheckListBox With selected Files in OpenFileDialog

        private void SPCWBConsolidatedReport_Load(object sender, EventArgs e)
        {
            //sfdgSPCWBConsolidatedReport.VerticalOverScrollMode = VerticalOverScrollMode.None;
            dtFillDatagridWithFiles.Columns.Add(colName, typeof(string));
            dtFillDatagridWithFiles.Columns.Add(colName1, typeof(string));
            dtFillDatagridWithFiles.Columns.Add(colName2, typeof(string));


            this.sfdgSPCWBConsolidatedReport.Columns.Add(new GridCheckBoxSelectorColumn()
            {
                MappingName = "fldSelect",
                HeaderText = string.Empty,
                TrueValue = "True",
                FalseValue = "False",
                AllowCheckBoxOnHeader = true,
                Width = 34,
                CheckBoxSize = new Size(14, 14)
            });

        }

        private void sfSPCWBConsolidatedReport_Click(object sender, EventArgs e)
        {

        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            try
            {

                string fileName = "sample.xlsx";
                Process[] processes = Process.GetProcessesByName("EXCEL");
                foreach (Process process in processes)
                {
                    if (process.MainWindowTitle.Contains(fileName))
                    {
                        // File is open
                        MessageBox.Show("The Excel file is already open, please close it and try again.", "File Already Open", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //this.Cursor = Cursors.Default;
                        return;
                    }
                }
               
                List<clsConsolidatedReportData> lstclsConsolidatedReportData = new List<clsConsolidatedReportData>();

                DataTable dtCharacteristic = new DataTable();
                DataTable dtCharactersticData = new DataTable();
                DataTable dtSgStat = new DataTable();
                DataTable dtSgHeader = new DataTable();
                DataTable dtSgData = new DataTable();
                DataTable dtTrace = new DataTable();
                DataTable dtTraceCategory = new DataTable();
                //Step1 Define Datatable dtAll

                DataTable dtAll = new DataTable();
                arrFileName.Clear();
                if (sfdgSPCWBConsolidatedReport.SelectedItems.Count > 0)
                {
                    foreach (var item in sfdgSPCWBConsolidatedReport.SelectedItems)
                    {
                        var datarow = (item as DataRowView).Row;
                        arrFileName.Add(new KeyValuePair<string, string>(datarow["File Name"].ToString(), datarow["File Path"].ToString()));
                    }
                    foreach (KeyValuePair<string, string> ele in arrFileName)
                    {
                        FileName = ele.Key;
                        string filepath = ele.Value;
                        int filecharcount = filepath.Length;

                        clsConsolidatedReportData oclsConsolidatedReportData = new clsConsolidatedReportData();
                        oclsConsolidatedReportData.clsChars = new List<clsChars>();
                        dtAll.Columns.Clear();
                        dtSgData.Clear();
                        dtTrace.Clear();
                        dtCharacteristic.Clear();
                        //Step2 Read file
                        string conString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=";
                        con = new OleDbConnection(conString + filepath);
                        con.Open();
                        //step3 Split File Name into Product Code and Grade
                        var filenameWithoutExtension = Path.GetFileNameWithoutExtension(FileName);
                        filename = Path.GetFileName(filenameWithoutExtension);

                        Grade = filename.Substring(filename.Length - 5);
                        ProductCode = filename.Substring(0, filename.Length - 5);

                        string charquery = "Select * from Characterstic";
                        OleDbCommand cmd = new OleDbCommand(charquery, con);
                        reader = cmd.ExecuteReader();
                        dtCharacteristic.Load(reader);

                        string SGDataQuery1 = "Select * from SGDATA";
                        OleDbCommand sgdatacmd = new OleDbCommand(SGDataQuery1, con);
                        reader = sgdatacmd.ExecuteReader();
                        dtSgData.Load(reader);

                        columntext = "";
                        dtTraceCategory = GetTraceCatFromTraceCategory(filepath);
                        foreach (System.Data.DataRow dttracecatrow in dtTraceCategory.Rows)
                        {
                            TraceCat = dttracecatrow["TraceCat"].ToString().TrimEnd();
                            //  TraceCat = GetTraceCatFromTrace(FilePath);
                            if ((TraceCat == "MO NO") || (TraceCat == "MO.NO") || (TraceCat == "Mo.No") || (TraceCat == "mo.no") || (TraceCat == "mo no") || (TraceCat == "MO No.") || (TraceCat == "M.O.") || (TraceCat == "M.O. NO".ToUpper() || (TraceCat == "MO".ToUpper() || (TraceCat == "OrderNo".ToUpper() || (TraceCat == "Order.No") || (TraceCat == "Order NO") || (TraceCat == "Order No") || (TraceCat == "ORDER NO")))))
                            {
                                // dtAll.Columns.Add(TraceCat, typeof(string));))
                                columntext = TraceCat;
                                TraceCategory = columntext;
                            }
                        }
                        if (columntext == "")
                        {
                            dtTraceCategory.Columns.Clear();
                            dtTraceCategory.Clear();
                            dtTraceCategory = GetTraceCatFromTraceCategory(filepath);
                            string FileName = Path.GetFileName(filepath);
                            MessageBox.Show("The traceability order no could not be found in this file " + FileName + " with " + filepath + " Choose the Order No category from the list given below. ", "Traceability Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            TraceabilitySelection tracevalueselect = new TraceabilitySelection(dtTraceCategory);
                            if (dtTraceCategory.Rows.Count != 0)
                            {
                                DialogResult dr = tracevalueselect.ShowDialog();
                                if (dr == DialogResult.OK)
                                {
                                    TraceCat = tracevalueselect.GetValue();
                                    columntext = TraceCat;
                                    TraceCategory = columntext;

                                }
                                if (dr == DialogResult.Cancel)
                                {
                                    return;
                                }

                            }
                            else if (dtTraceCategory.Rows.Count == 0)
                            {
                                DialogResult dialogResult = MessageBox.Show("There are no traceabilities defined in this file " + FileName + " with " + filepath + ". This file and its data will not be exported to Excel.Do you want to continue export to excel report without this file?", "Traceability Selection", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    //this.Close();
                                }
                                else if (dialogResult == DialogResult.No)
                                {
                                    return;
                                }

                            }



                        }

                        //TraceCat = GetTraceCatFromTrace(filepath);
                        string TraceValueQuery = "Select Distinct sgtrace.SGNO,Trace.Tracevalue,Trace.TraceId,Trace.TraceCat from Trace left join sgtrace on sgtrace.TraceID=Trace.TraceID where TraceCat='" + columntext + "' order by SgTrace.SGNO asc ";
                        OleDbCommand tracecmd = new OleDbCommand(TraceValueQuery, con);
                        reader = tracecmd.ExecuteReader();
                        dtTrace.Load(reader);


                        foreach (System.Data.DataRow dttracerow in dtTrace.Rows)
                        {
                            if (dttracerow["Tracevalue"].ToString() != "")
                            {
                                int TraceId = Convert.ToInt32(dttracerow["TraceId"]);
                                TraceCategory = dttracerow["TraceCat"].ToString();
                                Tracevalue = dttracerow["Tracevalue"].ToString();
                                SGNO = (dttracerow["SGNO"]).ToString();
                                foreach (System.Data.DataRow row in dtCharacteristic.Rows)
                                {
                                    CharName = row["CharName"].ToString();
                                    charId = Convert.ToInt32(row["CharID"]);
                                    accessObj = new SPCAccessModule.SPCAccessDB(filepath);
                                    oclsConsolidatedReportData.FilePath = filepath.Trim();
                                    // clsChars oclsChars = new clsChars();
                                    if (SGNO != "")
                                    {
                                        
                                        reader = accessObj.ReadDataForConsolidtedReport(filepath, SGNO, Tracevalue, TraceId, charId);
                                        
                                        while (reader.Read())
                                        {
                                            clsChars oclsChars = new clsChars();
                                            oclsConsolidatedReportData.ProductCode = ProductCode.Trim();
                                            oclsConsolidatedReportData.Grade = Grade.Trim();
                                            oclsChars.SubgroupNumber = ((reader["SGNO"]).ToString()).Trim();
                                            oclsChars.TraceId = ((reader["TraceId"]).ToString()).Trim();
                                            oclsChars.TraceCategory = ((reader["TraceCat"].ToString().Trim()));
                                            oclsChars.OrderNumber = (reader["Tracevalue"]).ToString().Trim();
                                            oclsChars.SGDate = Convert.ToDateTime(reader["SGDate"]);
                                            oclsChars.ParameterName = CharName;
                                            oclsChars.TOLLower = reader["LSL"].ToString();
                                            oclsChars.Target = reader["Target"].ToString();
                                            oclsChars.TOLUpper = reader["USL"].ToString();
                                            oclsChars.ActReading = reader["Value"].ToString();
                                            oclsConsolidatedReportData.clsChars.Add(oclsChars);
                                        }
                                        accessObj.CloseConnection();
                                    }
                                }
                            }
                        }
                        lstclsConsolidatedReportData.Add(oclsConsolidatedReportData);
                        con.Close();
                    }

                    DataTable dtCharacterstic;
                    //Export a Data To Excel
                    using (ExcelEngine ExcelEngineObject = new ExcelEngine())
                    {

                        IApplication Application = ExcelEngineObject.Excel;
                        Application.DefaultVersion = ExcelVersion.Excel2013;
                        IWorkbook workbook = ExcelEngineObject.Excel.Workbooks.Open("Sample.xltx", ExcelOpenType.Automatic);
                        bool chkifFileisFirstFile = false;
                        IWorksheet Worksheet = workbook.Worksheets[0];
                        int Filecount = 0;

                        //if (lstclsConsolidatedReportData[0].clsChars.Count > 0)
                        //   {
                        foreach (var itemFile in lstclsConsolidatedReportData)
                        {
                            if (itemFile.FilePath != null)
                            {
                                this.Cursor = Cursors.WaitCursor;
                            if (Filecount <= lstclsConsolidatedReportData.Count - 1)
                            {

                                    FileName = Path.GetFileName(itemFile.FilePath);
                                    if (lstclsConsolidatedReportData[Filecount].clsChars.Count != 0)
                                    {
                                        // Filecount = 1;
                                        
                                        
                                        #region Check current file path with first file 
                                        if (itemFile.FilePath.ToString() == lstclsConsolidatedReportData[0].FilePath)
                                        {
                                            chkifFileisFirstFile = true;
                                        }
                                        else
                                        {
                                            chkifFileisFirstFile = false;
                                        }
                                        #endregion Check current file path with first file 
                                        dtAll.Columns.Clear();
                                        dtAll.Clear();
                                        dtCharacteristic.Clear();
                                        dtAll.Columns.Add("SrNo", typeof(int));//Sr No Column added in datatable
                                        dtAll.Columns["SrNo"].AutoIncrement = true;
                                        dtAll.Columns["SrNo"].AutoIncrementSeed = 1;
                                        dtAll.Columns["SrNo"].AutoIncrementStep = 1;

                                        dtAll.Columns.Add("SubgroupNo", typeof(int));

                                        //clsChars oclsChars = new clsChars();
                                        //TraceCategory = oclsChars.TraceCategory;
                                        string TraceCategory = lstclsConsolidatedReportData[Filecount].clsChars[0].TraceCategory;
                                        dtAll.Columns.Add(TraceCategory, typeof(string));
                                        //}
                                        //TraceCategory=lstclsConsolidatedReportData[0].clsChars

                                        dtAll.Columns.Add("ProductCode", typeof(string));
                                        dtAll.Columns.Add("Grade", typeof(string));
                                        dtAll.Columns.Add("Date", typeof(DateTime));

                                        con = new OleDbConnection(conString + itemFile.FilePath);
                                        con.Open();

                                        string charquery = "Select * from Characterstic";
                                        OleDbCommand charcmd = new OleDbCommand(charquery, con);
                                        reader = charcmd.ExecuteReader();
                                        dtCharacteristic.Load(reader);
                                        con.Close();

                                        #region Check Whether SGSize Same or Different in Characterstic
                                        var distinctValues = dtCharacteristic.AsEnumerable().Select(row => row.Field<double>("SGSize")).Distinct();
                                        if (distinctValues.Count() == 1)
                                        {
                                            // MessageBox.Show("All values in the column are the same.");
                                            dtCharactersticData = GetCharacteristicData(itemFile.FilePath, distinctValues);
                                        }
                                        else if (distinctValues.Count() > 1)
                                        {
                                            // MessageBox.Show("All values in the column are the different.");
                                            dtCharactersticData = GetCharacteristicData(itemFile.FilePath, distinctValues);
                                        }
                                        #endregion

                                        bool CheckIfRowExistInDatable = false;
                                        foreach (System.Data.DataRow row in dtCharactersticData.Rows)
                                        {
                                            charId = Convert.ToInt32(row["CharID"]);
                                            dtAll.Columns.Add(row["CharName"] + " " + "TOL Lower");
                                            dtAll.Columns.Add(row["CharName"] + " " + "Target");
                                            dtAll.Columns.Add(row["CharName"] + " " + "TOL Upper");
                                            dtAll.Columns.Add(row["CharName"] + " " + "Actual Reading");

                                            var dataofchar = itemFile.clsChars.AsEnumerable().Where(x => x.ParameterName == row["CharName"].ToString()).ToList();
                                            var sgSizeOfCurChar = dtCharactersticData.AsEnumerable().Where(x => x.Field<string>("CharName") == row["CharName"].ToString()).Select(x => new { size = x["SGSize"] }).FirstOrDefault();
                                            var maxsgSize = Convert.ToInt32(dtCharactersticData.AsEnumerable().Max(s => s["SGSize"]));
                                            int AddedRowCount = 0;
                                            int SGRowCount = 0;
                                            int RowCount = 0;
                                            bool rowExist = false;
                                            if (dataofchar.Count == 0)
                                            {
                                                System.Data.DataRow drAll = null;
                                                if (CheckIfRowExistInDatable == false)
                                                {
                                                    drAll = dtAll.NewRow();
                                                    drAll["ProductCode"] = itemFile.ProductCode;
                                                    drAll["Grade"] = itemFile.Grade;
                                                    dtAll.Rows.Add(drAll);
                                                }
                                            }
                                            if (AddedRowCount <= maxsgSize)
                                            {
                                                System.Data.DataRow drAll = null;
                                                AddedRowCount++;
                                                RowCount = 0;
                                                if ((AddedRowCount) <= Convert.ToInt32(sgSizeOfCurChar.size))
                                                {
                                                    int DataCharValue = 0;
                                                    while (DataCharValue < dataofchar.Count)
                                                    {
                                                        var CharAlldata = dataofchar[DataCharValue];

                                                        if (CheckIfRowExistInDatable == false)
                                                        {

                                                            if (row["CharName"].ToString() == CharAlldata.ParameterName.ToString())
                                                            {
                                                                drAll = dtAll.NewRow();
                                                                drAll["ProductCode"] = itemFile.ProductCode;
                                                                drAll["Grade"] = itemFile.Grade;
                                                                drAll["SubgroupNo"] = CharAlldata.SubgroupNumber;

                                                                if (CharAlldata.OrderNumber != "")
                                                                {
                                                                    drAll[TraceCategory] = CharAlldata.OrderNumber;
                                                                }

                                                                drAll["Date"] = CharAlldata.SGDate.Date.ToShortDateString();
                                                                drAll[row["CharName"].ToString() + ' ' + "TOL Lower"] = CharAlldata.TOLLower;
                                                                drAll[row["CharName"].ToString() + ' ' + "Target"] = CharAlldata.Target;
                                                                drAll[row["CharName"].ToString() + ' ' + "TOL Upper"] = CharAlldata.TOLUpper;
                                                                drAll[row["CharName"].ToString() + ' ' + "Actual Reading"] = CharAlldata.ActReading;

                                                                dtAll.Rows.Add(drAll);

                                                            }
                                                            DataCharValue++;
                                                        }
                                                        else
                                                        {

                                                            if (SGRowCount < maxsgSize)
                                                            {
                                                                if (SGRowCount < Convert.ToInt32(sgSizeOfCurChar.size))
                                                                {
                                                                    if (dtAll.Rows.Count > 0)
                                                                    {
                                                                        dtAll.Rows[RowCount].SetField(row["CharName"] + " " + "TOL Lower", CharAlldata.TOLLower);
                                                                        dtAll.Rows[RowCount].SetField(row["CharName"] + " " + "Target", CharAlldata.Target);
                                                                        dtAll.Rows[RowCount].SetField(row["CharName"] + " " + "TOL Upper", CharAlldata.TOLUpper);
                                                                        dtAll.Rows[RowCount].SetField(row["CharName"] + " " + "Actual Reading", CharAlldata.ActReading);
                                                                        DataCharValue++;

                                                                    }
                                                                }
                                                                SGRowCount++;
                                                                RowCount++;
                                                                if (RowCount == dtAll.Rows.Count)
                                                                {
                                                                    RowCount = 0;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                //count must be init
                                                                SGRowCount = 0;
                                                                // RowCount = 0;
                                                            }

                                                        }

                                                    }
                                                    CheckIfRowExistInDatable = true;
                                                }

                                            }
                                        }
                                    }
                                    else
                                    {
                                        File.Delete("sample.xlsx");
                                        MessageBox.Show("" + FileName + "  this file is empty and does not have any data. This file will be skipped in the Excel report.  ", "SPWBConsolidated Report", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        //Filecount++;
                                    }
                                }


                                if (Worksheet != null && Worksheet.UsedRange != null)
                                {
                                    LastRowIndex = Worksheet.UsedRange.LastRow;

                                    //Get last filled row of excel
                                    int lastRowIndex = Worksheet.Rows.Length - 1;
                                }
                                this.Cursor = Cursors.Default;
                                //Method1
                                // write dt to excel row by row with columns
                                this.Cursor = Cursors.WaitCursor;
                                if (chkifFileisFirstFile == true)
                                {
                                    Worksheet.Range["A1:ZZ1"].CellStyle.Color = Color.Teal;
                                    Worksheet.Range["A1:ZZ1"].CellStyle.Font.RGBColor = Color.White;
                                    Worksheet.UsedRange.AutofitColumns();
                                    dtAll.Columns[2].ColumnName = "Order No";
                                    Worksheet.ImportDataTable(dtAll, true, 1, 1);//ImportDatatable to worksheet in excel
                                    Worksheet.Range["A1"].Activate();
                                    File.Delete("sample.xlsx");
                                    workbook.SaveAs("sample.xlsx");
                                    dtAll.Clear();
                                    Filecount++;
                                    //columntext = "";
                                }
                                else
                                {
                                    if (Worksheet.Columns.Length != 0)
                                    {
                                        IRange lastRowRange = Worksheet.Range[LastRowIndex + 2, 1, LastRowIndex + 2, Worksheet.Columns.Length];
                                        lastRowRange.CellStyle.Color = Color.Teal;
                                        lastRowRange.CellStyle.Font.RGBColor = Color.White;
                                        Worksheet.UsedRange.AutofitColumns();
                                        dtAll.Columns[2].ColumnName = "Order No";

                                        //lastRowRange.Characters(0, 0).Select();
                                        Worksheet.ImportDataTable(dtAll, true, LastRowIndex + 2, 1);//Append one file over another
                                        File.Delete("sample.xlsx");
                                        Worksheet.Range["A1"].Activate();
                                        //Worksheet.Range["A1"].po
                                        workbook.SaveAs("sample.xlsx");
                                        dtAll.Clear();
                                        Filecount++;
                                    }

                                }

                            }

                        }
                        ExcelEngineObject.Dispose();
                    }
                    if (File.Exists("sample.xlsx"))
                    {
                        Process process = new Process();
                        Process.Start("sample.xlsx");
                    }
                    this.Cursor = Cursors.Default;
                }
                // this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), ("Title").TrimEnd(c), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            // this.Cursor = Cursors.Default;
        }

          
       public void GetTraceCategoryWithTraceCat()
        {

        }
        //private void btnExportToExcel_Click(object sender, EventArgs e)
        //{

        //    try
        //    {
        //        this.Cursor = Cursors.WaitCursor;
        //        string fileName = "sample.xlsx";
        //        Process[] processes = Process.GetProcessesByName("EXCEL");
        //        foreach (Process process in processes)
        //        {
        //            if (process.MainWindowTitle.Contains(fileName))
        //            {
        //                // File is open
        //                MessageBox.Show("The file is already open, please close it and try again.", "File Already Open", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                this.Cursor = Cursors.Default;
        //                return;
        //            }
        //        }

        //        string filePath = System.AppDomain.CurrentDomain.BaseDirectory + fileName; // replace with your file path

        //        if (File.Exists(filePath))
        //        {
        //            File.Delete(filePath);
        //        }

        //        List<clsConsolidatedReportData> lstclsConsolidatedReportData = new List<clsConsolidatedReportData>();

        //        DataTable dtCharacteristic = new DataTable();
        //        DataTable dtSgStat = new DataTable();
        //        DataTable dtSgHeader = new DataTable();
        //        DataTable dtSgData = new DataTable();
        //        DataTable dtTrace = new DataTable();
        //        //Step1 Define Datatable dtAll

        //        DataTable dtAll = new DataTable();
        //        arrFileName.Clear();
        //        if (sfdgSPCWBConsolidatedReport.SelectedItems.Count > 0)
        //        {
        //            foreach (var item in sfdgSPCWBConsolidatedReport.SelectedItems)
        //            {
        //                var datarow = (item as DataRowView).Row;
        //                arrFileName.Add(new KeyValuePair<string, string>(datarow["File Name"].ToString(), datarow["File Path"].ToString()));
        //            }
        //            foreach (KeyValuePair<string, string> ele in arrFileName)
        //            {
        //                FileName = ele.Key;
        //                string filepath = ele.Value;
        //                int filecharcount = filepath.Length;

        //                clsConsolidatedReportData oclsConsolidatedReportData = new clsConsolidatedReportData();
        //                oclsConsolidatedReportData.clsChars = new List<clsChars>();
        //                dtAll.Columns.Clear();
        //                dtSgData.Clear();
        //                dtCharacteristic.Clear();
        //                //Step2 Read file
        //                string conString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=";
        //                con = new OleDbConnection(conString + filepath);
        //                con.Open();
        //                //step3 Split File Name into Product Code and Grade
        //                var filenameWithoutExtension = Path.GetFileNameWithoutExtension(FileName);
        //                string filename = Path.GetFileName(filenameWithoutExtension);

        //                Grade = filename.Substring(filename.Length - 4);
        //                ProductCode = filename.Substring(0, filename.Length - 4);


        //                string charquery = "Select * from Characterstic";
        //                OleDbCommand cmd = new OleDbCommand(charquery, con);
        //                reader = cmd.ExecuteReader();
        //                dtCharacteristic.Load(reader);

        //                string SGDataQuery1 = "Select * from SGDATA";
        //                OleDbCommand sgdatacmd = new OleDbCommand(SGDataQuery1, con);
        //                reader = sgdatacmd.ExecuteReader();
        //                dtSgData.Load(reader);

        //                string TraceCat = GetTraceCatFromTrace(filepath);
        //                string TraceValueQuery = "Select sgtrace.SGNO,Trace.Tracevalue from Trace left join sgtrace on sgtrace.TraceID=Trace.TraceID where TraceCat='" + TraceCat + "' order by SgTrace.SGNO asc ";
        //                OleDbCommand tracecmd = new OleDbCommand(TraceValueQuery, con);
        //                reader = tracecmd.ExecuteReader();
        //                // con.Close();
        //                dtTrace.Load(reader);
        //                foreach (System.Data.DataRow row in dtCharacteristic.Rows)
        //                {
        //                    CharName = row["CharName"].ToString();
        //                    charId = Convert.ToInt32(row["CharID"]);

        //                    foreach (System.Data.DataRow row1 in dtTrace.Rows)
        //                    {
        //                        if (row1["Tracevalue"].ToString() != "")
        //                        {
        //                            // SGNO = row1.Field<Int32?>("SGNO") ?? 0; ;
        //                            Tracevalue = row1["Tracevalue"].ToString();
        //                            SGNO = (row1["SGNO"]).ToString();

        //                        }
        //                        if (SGNO != "")
        //                        {
        //                            accessObj = new SPCAccessModule.SPCAccessDB(filepath);
        //                            reader = accessObj.ReadDataForConsolidtedReport(filepath, SGNO, Tracevalue, charId);
        //                            oclsConsolidatedReportData.FilePath = filepath.Trim();
        //                            while (reader.Read())
        //                            {
        //                                clsChars oclsChars = new clsChars();

        //                                oclsConsolidatedReportData.ProductCode = ProductCode.Trim();
        //                                oclsConsolidatedReportData.Grade = Grade.Trim();
        //                                oclsChars.SubgroupNumber = ((reader["SGNO"]).ToString()).Trim();
        //                                oclsChars.OrderNumber = (reader["Tracevalue"]).ToString().Trim();
        //                                oclsChars.SGDate = Convert.ToDateTime(reader["SGDate"]);
        //                                oclsChars.ParameterName = CharName;
        //                                oclsChars.TOLLower = reader["LSL"].ToString();
        //                                oclsChars.Target = reader["Target"].ToString();
        //                                oclsChars.TOLUpper = reader["USL"].ToString();
        //                                oclsChars.ActReading = reader["Value"].ToString();
        //                                oclsConsolidatedReportData.clsChars.Add(oclsChars);
        //                            }
        //                            accessObj.CloseConnection();
        //                        }
        //                    }
        //                    lstclsConsolidatedReportData.Add(oclsConsolidatedReportData);
        //                    con.Close();
        //                }

        //            }
        //            DataTable dtCharacterstic;
        //            //Export a Data To Excel
        //            using (ExcelEngine ExcelEngineObject = new ExcelEngine())
        //            {
        //                IApplication Application = ExcelEngineObject.Excel;
        //                Application.DefaultVersion = ExcelVersion.Excel2013;
        //                IWorkbook workbook = ExcelEngineObject.Excel.Workbooks.Open("Sample.xltx", ExcelOpenType.Automatic);
        //                bool chkifFileisFirstFile = false;
        //                IWorksheet Worksheet = workbook.Worksheets[0];
        //                if (lstclsConsolidatedReportData.Count > 0)
        //                {
        //                    foreach (var itemFile in lstclsConsolidatedReportData)
        //                    {
        //                        if (itemFile.FilePath != null)
        //                        {
        //                            #region Check current file path with first file 
        //                            if (itemFile.FilePath.ToString() == lstclsConsolidatedReportData[0].FilePath)
        //                            {
        //                                chkifFileisFirstFile = true;
        //                            }
        //                            else
        //                            {
        //                                chkifFileisFirstFile = false;
        //                            }
        //                            #endregion Check current file path with first file 
        //                            dtAll.Columns.Clear();
        //                            dtAll.Clear();
        //                            dtAll.Columns.Add("SrNo", typeof(int));//Sr No Column added in datatable
        //                            dtAll.Columns["SrNo"].AutoIncrement = true;
        //                            dtAll.Columns["SrNo"].AutoIncrementSeed = 1;
        //                            dtAll.Columns["SrNo"].AutoIncrementStep = 1;

        //                            dtAll.Columns.Add("SubgroupNo", typeof(int));
        //                            dtAll.Columns.Add(TraceCat, typeof(string));
        //                            dtAll.Columns.Add("ProductCode", typeof(string));
        //                            dtAll.Columns.Add("Grade", typeof(string));
        //                            dtAll.Columns.Add("Date", typeof(DateTime));

        //                            dtCharacterstic = GetCharacteristicData(itemFile.FilePath);
        //                            bool checkifrowexistindatable = false;
        //                            int irowCount = 0;
        //                            foreach (System.Data.DataRow row in dtCharacterstic.Rows)
        //                            {
        //                                dtAll.Columns.Add(row["CharName"] + " " + "TOL Lower");
        //                                dtAll.Columns.Add(row["CharName"] + " " + "Target");
        //                                dtAll.Columns.Add(row["CharName"] + " " + "TOL Upper");
        //                                dtAll.Columns.Add(row["CharName"] + " " + "Actual Reading");
        //                                int i = 0;
        //                                var dataofchar = itemFile.clsChars.AsEnumerable().Where(x => x.ParameterName == row["CharName"].ToString()).ToList();
        //                                if (dataofchar.Count == 0)
        //                                {
        //                                    System.Data.DataRow drAll = null;
        //                                    if (checkifrowexistindatable == false)
        //                                    {
        //                                        drAll = dtAll.NewRow();
        //                                        drAll["ProductCode"] = itemFile.ProductCode;
        //                                        drAll["Grade"] = itemFile.Grade;
        //                                        dtAll.Rows.Add(drAll);
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    foreach (var item in dataofchar)
        //                                    {
        //                                        if (row["CharName"].ToString() == item.ParameterName.ToString())
        //                                        {
        //                                            System.Data.DataRow drAll = null;
        //                                            if (checkifrowexistindatable == false)
        //                                            {
        //                                                drAll = dtAll.NewRow();
        //                                                drAll["ProductCode"] = itemFile.ProductCode;
        //                                                drAll["Grade"] = itemFile.Grade;

        //                                                drAll["SubgroupNo"] = item.SubgroupNumber;
        //                                                if (item.OrderNumber != "")
        //                                                {
        //                                                    drAll[TraceCat] = item.OrderNumber;
        //                                                }
        //                                                drAll["Date"] = item.SGDate.Date.ToShortDateString();

        //                                                drAll[row["CharName"].ToString() + ' ' + "TOL Lower"] = item.TOLLower;
        //                                                drAll[row["CharName"].ToString() + ' ' + "Target"] = item.Target;
        //                                                drAll[row["CharName"].ToString() + ' ' + "TOL Upper"] = item.TOLUpper;
        //                                                drAll[row["CharName"].ToString() + ' ' + "Actual Reading"] = item.ActReading;
        //                                                dtAll.Rows.Add(drAll);
        //                                            }
        //                                            else
        //                                            {
        //                                                irowCount = dtAll.Rows.Count;

        //                                                while (i < dtAll.Rows.Count)
        //                                                {

        //                                                    dtAll.Rows[i].SetField(row["CharName"] + " " + "TOL Lower", item.TOLLower);
        //                                                    dtAll.Rows[i].SetField(row["CharName"] + " " + "Target", item.Target);
        //                                                    dtAll.Rows[i].SetField(row["CharName"] + " " + "TOL Upper", item.TOLUpper);
        //                                                    dtAll.Rows[i].SetField(row["CharName"] + " " + "Actual Reading", item.ActReading);
        //                                                    // i++;
        //                                                    break;
        //                                                }

        //                                            }
        //                                            i++;
        //                                        }

        //                                    }
        //                                }

        //                                checkifrowexistindatable = true;
        //                            }
        //                        }

        //                        LastRowIndex = Worksheet.UsedRange.LastRow;//Get last filled row of excel
        //                        int lastRowIndex = Worksheet.Rows.Length - 1;
        //                        this.Cursor = Cursors.Default;
        //                        //Method1
        //                        // write dt to excel row by row with columns

        //                        if (chkifFileisFirstFile == true)
        //                        {
        //                            Worksheet.Range["A1:Z1"].CellStyle.Color = Color.Teal;
        //                            Worksheet.Range["A1:Z1"].CellStyle.Font.RGBColor = Color.White;
        //                            Worksheet.UsedRange.AutofitColumns();
        //                            Worksheet.ImportDataTable(dtAll, true, 1, 1);//ImportDatatable to worksheet in excel
        //                            workbook.SaveAs("sample.xlsx");
        //                        }
        //                        else
        //                        {
        //                            IRange lastRowRange = Worksheet.Range[LastRowIndex + 2, 1, LastRowIndex + 2, Worksheet.Columns.Length];
        //                            lastRowRange.CellStyle.Color = Color.Teal;
        //                            lastRowRange.CellStyle.Font.RGBColor = Color.White;
        //                            Worksheet.UsedRange.AutofitColumns();
        //                            Worksheet.ImportDataTable(dtAll, true, LastRowIndex + 2, 1);//Append one file over another
        //                            workbook.SaveAs("sample.xlsx");
        //                        }

        //                        // }
        //                    }
        //                }

        //                if (File.Exists("sample.xlsx"))
        //                {
        //                    Process process = new Process();
        //                    Process.Start("sample.xlsx");
        //                    workbook.Close();
        //                }
        //                else
        //                {
        //                    MessageBox.Show("Data From this File Cannot be Exported", "Title".TrimEnd(), MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                    this.Cursor = Cursors.Default;
        //                    return;
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString(), ("Title").TrimEnd(c), MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    this.Cursor = Cursors.Default;
        //}

        //private void btnExportToExcel_Click(object sender, EventArgs e)
        //{

        //    try
        //    {
        //        this.Cursor = Cursors.WaitCursor;
        //        string fileName = "sample.xlsx";
        //        Process[] processes = Process.GetProcessesByName("EXCEL");
        //        foreach (Process process in processes)
        //        {
        //            if (process.MainWindowTitle.Contains(fileName))
        //            {
        //                // File is open
        //                MessageBox.Show("The file is already open, please close it and try again.", "File Already Open", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                this.Cursor = Cursors.Default;
        //                return;
        //            }
        //        }

        //        //string filePath = System.AppDomain.CurrentDomain.BaseDirectory + fileName; // replace with your file path

        //        //if (File.Exists(filePath))
        //        //{
        //        //    File.Delete(filePath);
        //        //}

        //        List<clsConsolidatedReportData> lstclsConsolidatedReportData = new List<clsConsolidatedReportData>();

        //        DataTable dtCharacteristic = new DataTable();
        //        DataTable dtCharactersticData = new DataTable();
        //        DataTable dtSgStat = new DataTable();
        //        DataTable dtSgHeader = new DataTable();
        //        DataTable dtSgData = new DataTable();
        //        DataTable dtTrace = new DataTable();
        //        DataTable dtTraceCategory = new DataTable();
        //        //Step1 Define Datatable dtAll

        //        DataTable dtAll = new DataTable();
        //        arrFileName.Clear();
        //        if (sfdgSPCWBConsolidatedReport.SelectedItems.Count > 0)
        //        {
        //            foreach (var item in sfdgSPCWBConsolidatedReport.SelectedItems)
        //            {
        //                var datarow = (item as DataRowView).Row;
        //                arrFileName.Add(new KeyValuePair<string, string>(datarow["File Name"].ToString(), datarow["File Path"].ToString()));
        //            }
        //            foreach (KeyValuePair<string, string> ele in arrFileName)
        //            {
        //                FileName = ele.Key;
        //                string filepath = ele.Value;
        //                int filecharcount = filepath.Length;

        //                clsConsolidatedReportData oclsConsolidatedReportData = new clsConsolidatedReportData();
        //                oclsConsolidatedReportData.clsChars = new List<clsChars>();
        //                dtAll.Columns.Clear();
        //                dtSgData.Clear();
        //                dtCharacteristic.Clear();
        //                //Step2 Read file
        //                string conString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=";
        //                con = new OleDbConnection(conString + filepath);
        //                con.Open();
        //                //step3 Split File Name into Product Code and Grade
        //                var filenameWithoutExtension = Path.GetFileNameWithoutExtension(FileName);
        //                string filename = Path.GetFileName(filenameWithoutExtension);

        //                Grade = filename.Substring(filename.Length - 4);
        //                ProductCode = filename.Substring(0, filename.Length - 4);


        //                string charquery = "Select * from Characterstic";
        //                OleDbCommand cmd = new OleDbCommand(charquery, con);
        //                reader = cmd.ExecuteReader();
        //                dtCharacteristic.Load(reader);

        //                string SGDataQuery1 = "Select * from SGDATA";
        //                OleDbCommand sgdatacmd = new OleDbCommand(SGDataQuery1, con);
        //                reader = sgdatacmd.ExecuteReader();
        //                dtSgData.Load(reader);

        //                dtTraceCategory = GetTraceCatFromTraceCategory(filepath);
        //                // TraceCat = GetTraceCatFromTrace(filepath);
        //                //foreach (System.Data.DataRow dttraceCatrow in dtTraceCategory.Rows)
        //                //{
        //                //    string TraceCat1 = (dttraceCatrow["TraceCat"]).ToString();
        //                //    string TraceValueQuery = "Select Distinct sgtrace.SGNO,Trace.Tracevalue,Trace.TraceCat from Trace left join sgtrace on sgtrace.TraceID=Trace.TraceID where TraceCat='" + TraceCat1 + "' order by SgTrace.SGNO asc ";
        //                //    OleDbCommand tracecmd = new OleDbCommand(TraceValueQuery, con);
        //                //    reader = tracecmd.ExecuteReader();
        //                //    dtTrace.Load(reader);
        //                //}


        //                foreach (System.Data.DataRow dttracerow in dtTrace.Rows)
        //                {
        //                    if (dttracerow["Tracevalue"].ToString() != "")
        //                    {
        //                        Tracevalue = dttracerow["Tracevalue"].ToString();
        //                        SGNO = (dttracerow["SGNO"]).ToString();
        //                        //if (!(dttracerow["SGNO"] is DBNull))
        //                        //{
        //                        //    int SGNO = Convert.ToInt32(dttracerow["SGNO"]);
        //                        //    //if (SGNO == SGNOForEmail)
        //                        //    //{
        //                        //        Tracevalue = dttracerow["Tracevalue"].ToString();
        //                        //    //}
        //                        //}


        //                        foreach (System.Data.DataRow row in dtCharacteristic.Rows)
        //                        {
        //                            CharName = row["CharName"].ToString();
        //                            charId = Convert.ToInt32(row["CharID"]);
        //                            if (SGNO != "")
        //                            {
        //                                accessObj = new SPCAccessModule.SPCAccessDB(filepath);
        //                                reader = accessObj.ReadDataForConsolidtedReport(filepath, SGNO, Tracevalue, charId);
        //                                oclsConsolidatedReportData.FilePath = filepath.Trim();
        //                                while (reader.Read())
        //                                {
        //                                    clsChars oclsChars = new clsChars();
        //                                    oclsConsolidatedReportData.ProductCode = ProductCode.Trim();
        //                                    oclsConsolidatedReportData.Grade = Grade.Trim();
        //                                    oclsChars.SubgroupNumber = ((reader["SGNO"]).ToString()).Trim();
        //                                    oclsChars.TraceCategory = ((reader["TraceCat"].ToString().Trim()));
        //                                    oclsChars.OrderNumber = (reader["Tracevalue"]).ToString().Trim();
        //                                    oclsChars.SGDate = Convert.ToDateTime(reader["SGDate"]);
        //                                    oclsChars.ParameterName = CharName;
        //                                    oclsChars.TOLLower = reader["LSL"].ToString();
        //                                    oclsChars.Target = reader["Target"].ToString();
        //                                    oclsChars.TOLUpper = reader["USL"].ToString();
        //                                    oclsChars.ActReading = reader["Value"].ToString();
        //                                    oclsConsolidatedReportData.clsChars.Add(oclsChars);
        //                                }
        //                                accessObj.CloseConnection();

        //                            }


        //                        }
        //                    }
        //                }
        //                lstclsConsolidatedReportData.Add(oclsConsolidatedReportData);
        //                con.Close();


        //            }
        //            DataTable dtCharacterstic;
        //            //Export a Data To Excel
        //            using (ExcelEngine ExcelEngineObject = new ExcelEngine())
        //            {
        //                IApplication Application = ExcelEngineObject.Excel;
        //                Application.DefaultVersion = ExcelVersion.Excel2013;
        //                IWorkbook workbook = ExcelEngineObject.Excel.Workbooks.Open("Sample.xltx", ExcelOpenType.Automatic);
        //                bool chkifFileisFirstFile = false;
        //                IWorksheet Worksheet = workbook.Worksheets[0];
        //                bool checkWhetherTraceValueisPresent = false;
        //                #region count number of tracecategory whose tracevalue is not equal to zero
        //                string TraceValueQuery = "select Count(TraceCat) from Trace where TraceValue<>'" + 0 + "' ";
        //                OleDbCommand tracecmd = new OleDbCommand(TraceValueQuery, con);
        //                con.Open();
        //                tracecatcount = Convert.ToInt32(tracecmd.ExecuteScalar());
        //                //int tracecatcount = Convert.ToInt32(reader[0]);
        //                con.Close();
        //                #endregion count number of tracecategory whose tracevalue is not equal to zero


        //                if (lstclsConsolidatedReportData.Count > 0)
        //                {
        //                    foreach (var itemFile in lstclsConsolidatedReportData)
        //                    {
        //                        if (itemFile.FilePath != null)
        //                        {
        //                            #region Check current file path with first file 
        //                            if (itemFile.FilePath.ToString() == lstclsConsolidatedReportData[0].FilePath)
        //                            {
        //                                chkifFileisFirstFile = true;
        //                            }
        //                            else
        //                            {
        //                                chkifFileisFirstFile = false;
        //                            }
        //                            #endregion Check current file path with first file 
        //                            dtAll.Columns.Clear();
        //                            dtAll.Clear();
        //                            dtCharacteristic.Clear();
        //                            dtAll.Columns.Add("SrNo", typeof(int));//Sr No Column added in datatable
        //                            dtAll.Columns["SrNo"].AutoIncrement = true;
        //                            dtAll.Columns["SrNo"].AutoIncrementSeed = 1;
        //                            dtAll.Columns["SrNo"].AutoIncrementStep = 1;

        //                            dtAll.Columns.Add("SubgroupNo", typeof(int));
        //                            foreach (System.Data.DataRow dtTracerow in dtTraceCategory.Rows)
        //                            {
        //                                TraceCat = (dtTracerow["TraceCat"]).ToString();
        //                                // dtDataView.Columns.Add(TraceCategory, typeof(string));
        //                                dtAll.Columns.Add(TraceCat, typeof(string));
        //                            }
        //                            dtAll.Columns.Add("ProductCode", typeof(string));
        //                            dtAll.Columns.Add("Grade", typeof(string));
        //                            dtAll.Columns.Add("Date", typeof(DateTime));

        //                            con = new OleDbConnection(conString + itemFile.FilePath);
        //                            con.Open();

        //                            string charquery = "Select * from Characterstic";
        //                            OleDbCommand cmd = new OleDbCommand(charquery, con);
        //                            reader = cmd.ExecuteReader();
        //                            dtCharacteristic.Load(reader);
        //                            con.Close();

        //                            #region Check Whether SGSize Same or Different in Characterstic
        //                            var distinctValues = dtCharacteristic.AsEnumerable().Select(row => row.Field<double>("SGSize")).Distinct();
        //                            if (distinctValues.Count() == 1)
        //                            {
        //                                // MessageBox.Show("All values in the column are the same.");
        //                                dtCharactersticData = GetCharacteristicData(itemFile.FilePath, distinctValues);

        //                            }
        //                            else if (distinctValues.Count() > 1)
        //                            {
        //                                // MessageBox.Show("All values in the column are the different.");
        //                                dtCharactersticData = GetCharacteristicData(itemFile.FilePath, distinctValues);
        //                            }
        //                            #endregion

        //                            bool CheckIfRowExistInDatable = false;
        //                            //int irowCount = 0;
        //                            foreach (System.Data.DataRow row in dtCharactersticData.Rows)
        //                            {
        //                                charId = Convert.ToInt32(row["CharID"]);
        //                                dtAll.Columns.Add(row["CharName"] + " " + "TOL Lower");
        //                                dtAll.Columns.Add(row["CharName"] + " " + "Target");
        //                                dtAll.Columns.Add(row["CharName"] + " " + "TOL Upper");
        //                                dtAll.Columns.Add(row["CharName"] + " " + "Actual Reading");
        //                                var dataofchar = itemFile.clsChars.AsEnumerable().Where(x => x.ParameterName == row["CharName"].ToString()).ToList();
        //                                var sgSizeOfCurChar = dtCharactersticData.AsEnumerable().Where(x => x.Field<string>("CharName") == row["CharName"].ToString()).Select(x => new { size = x["SGSize"] }).FirstOrDefault();
        //                                var maxsgSize = Convert.ToInt32(dtCharactersticData.AsEnumerable().Max(s => s["SGSize"]));
        //                                int AddedRowCount = 0;
        //                                int SGRowCount = 0;
        //                                int RowCount = 0;
        //                                bool rowExist = false;
        //                                if (dataofchar.Count == 0)
        //                                {
        //                                    System.Data.DataRow drAll = null;
        //                                    if (CheckIfRowExistInDatable == false)
        //                                    {
        //                                        drAll = dtAll.NewRow();
        //                                        drAll["ProductCode"] = itemFile.ProductCode;
        //                                        drAll["Grade"] = itemFile.Grade;
        //                                        dtAll.Rows.Add(drAll);
        //                                    }
        //                                }
        //                                if (AddedRowCount <= maxsgSize)
        //                                {
        //                                    System.Data.DataRow drAll = null;
        //                                    AddedRowCount++;
        //                                    RowCount = 0;
        //                                    if ((AddedRowCount) <= Convert.ToInt32(sgSizeOfCurChar.size))
        //                                    {
        //                                        int DataCharValue = 0;

        //                                        while (DataCharValue < dataofchar.Count)
        //                                        {
        //                                            var value = dataofchar[DataCharValue];
        //                                            //if (RowCount <= DataCharValue / 2)


        //                                            //    foreach (System.Data.DataRow dtTracerowcat in dtTraceCategory.Rows)
        //                                            //    {
        //                                            //        TraceCat = (dtTracerowcat["TraceCat"]).ToString();

        //                                            //        if (TraceCat == value.TraceCategory.ToString())
        //                                            //        {

        //                                            //            dtAll.Rows[RowCount].SetField(5, value.OrderNumber);
        //                                            //            RowCount++;
        //                                            //        }


        //                                            //    }


        //                                            if (CheckIfRowExistInDatable == false)
        //                                            {


        //                                                if (row["CharName"].ToString() == value.ParameterName.ToString())
        //                                                {
        //                                                    //if (DataCharValue < dataofchar.Count / tracecatcount)
        //                                                    //{

        //                                                    drAll = dtAll.NewRow();
        //                                                    drAll["ProductCode"] = itemFile.ProductCode;
        //                                                    drAll["Grade"] = itemFile.Grade;
        //                                                    drAll["SubgroupNo"] = value.SubgroupNumber;

        //                                                    if (value.OrderNumber != "")
        //                                                    {

        //                                                        foreach (System.Data.DataRow dtTracerowcat in dtTraceCategory.Rows)
        //                                                        {
        //                                                            TraceCat = (dtTracerowcat["TraceCat"]).ToString();

        //                                                            if (TraceCat == value.TraceCategory.ToString())
        //                                                            {

        //                                                                drAll[TraceCat] = value.OrderNumber;

        //                                                            }
        //                                                        }

        //                                                    }

        //                                                    drAll["Date"] = value.SGDate.Date.ToShortDateString();
        //                                                    drAll[row["CharName"].ToString() + ' ' + "TOL Lower"] = value.TOLLower;
        //                                                    drAll[row["CharName"].ToString() + ' ' + "Target"] = value.Target;
        //                                                    drAll[row["CharName"].ToString() + ' ' + "TOL Upper"] = value.TOLUpper;
        //                                                    drAll[row["CharName"].ToString() + ' ' + "Actual Reading"] = value.ActReading;
        //                                                    dtAll.Rows.Add(drAll);

        //                                                }
        //                                                //else if (DataCharValue > dataofchar.Count / tracecatcount)
        //                                                //{
        //                                                //    //var value1 = dataofchar[DataCharValue];
        //                                                //    if (RowCount < dataofchar.Count / tracecatcount)
        //                                                //        foreach (System.Data.DataRow dtTracerowcat in dtTraceCategory.Rows)
        //                                                //        {
        //                                                //            TraceCat = (dtTracerowcat["TraceCat"]).ToString();

        //                                                //            if (TraceCat == value.TraceCategory.ToString())
        //                                                //            {
        //                                                //                int columnIndex = dtAll.Columns.IndexOf(TraceCat);
        //                                                //                dtAll.Rows[RowCount].SetField(columnIndex, value.OrderNumber);
        //                                                //                RowCount++;
        //                                                //            }
        //                                                //            // DataCharValue++;

        //                                                //        }
        //                                            }

        //                                            DataCharValue++;
        //                                       // }



        //                                            else
        //                                            {
        //                                                if (SGRowCount < maxsgSize)
        //                                                {
        //                                                    if (SGRowCount < Convert.ToInt32(sgSizeOfCurChar.size))
        //                                                    {
        //                                                        if (dtAll.Rows.Count > 0)
        //                                                        {
        //                                                            //value = dataofchar[DataCharValue];
        //                                                            dtAll.Rows[RowCount].SetField(row["CharName"] + " " + "TOL Lower", value.TOLLower);
        //                                                            dtAll.Rows[RowCount].SetField(row["CharName"] + " " + "Target", value.Target);
        //                                                            dtAll.Rows[RowCount].SetField(row["CharName"] + " " + "TOL Upper", value.TOLUpper);
        //                                                            dtAll.Rows[RowCount].SetField(row["CharName"] + " " + "Actual Reading", value.ActReading);



        //                                                            // }

        //                                                            DataCharValue++;
        //                                                        }
        //                                                    }
        //                                                    SGRowCount++;
        //                                                    RowCount++;

        //                                                }
        //                                                else
        //                                                {
        //                                                    //count must be init
        //                                                    SGRowCount = 0;
        //                                                }

        //                                            }

        //                                        }
        //                                        checkWhetherTraceValueisPresent = true;
        //                                        CheckIfRowExistInDatable = true;
        //                                    }


        //                                }
        //                            }

        //                            LastRowIndex = Worksheet.UsedRange.LastRow;//Get last filled row of excel
        //                            int lastRowIndex = Worksheet.Rows.Length - 1;
        //                            this.Cursor = Cursors.Default;
        //                            //Method1
        //                            // write dt to excel row by row with columns

        //                            if (chkifFileisFirstFile == true)
        //                            {
        //                                Worksheet.Range["A1:ZZ1"].CellStyle.Color = Color.Teal;
        //                                Worksheet.Range["A1:ZZ1"].CellStyle.Font.RGBColor = Color.White;
        //                                Worksheet.UsedRange.AutofitColumns();
        //                                Worksheet.ImportDataTable(dtAll, true, 1, 1);//ImportDatatable to worksheet in excel
        //                                workbook.SaveAs("sample.xlsx");
        //                            }
        //                            else
        //                            {
        //                                IRange lastRowRange = Worksheet.Range[LastRowIndex + 2, 1, LastRowIndex + 2, Worksheet.Columns.Length];
        //                                lastRowRange.CellStyle.Color = Color.Teal;
        //                                lastRowRange.CellStyle.Font.RGBColor = Color.White;
        //                                Worksheet.UsedRange.AutofitColumns();
        //                                Worksheet.ImportDataTable(dtAll, true, LastRowIndex + 2, 1);//Append one file over another
        //                                workbook.SaveAs("sample.xlsx");
        //                            }

        //                        }
        //                    }

        //                    if (File.Exists("sample.xlsx"))
        //                    {
        //                        Process process = new Process();
        //                        Process.Start("sample.xlsx");
        //                        workbook.Close();
        //                    }
        //                    //else
        //                    //{
        //                    //    MessageBox.Show("Data From this File Cannot be Exported", "Title".TrimEnd(), MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                    //    this.Cursor = Cursors.Default;
        //                    //    return;
        //                    // }
        //                }
        //            }

        //        }
        //    }

        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString(), ("Title").TrimEnd(c), MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //       this.Cursor = Cursors.Default;
        //}

        public DataTable GetCharacteristicData(string filepath,IEnumerable<double>distinctvalues)
        {
            string charquery = "";
            DataTable dtCharacteristic = new DataTable();
            string conString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=";
            con = new OleDbConnection(conString + filepath);
            con.Open();
            if (distinctvalues.Count() == 1)
            {
               charquery = "Select * from Characterstic order by CharId asc";
            }
            else if(distinctvalues.Count()>1)
            {
               charquery = "Select * from Characterstic order by SGSize desc";
            }
            OleDbCommand cmd = new OleDbCommand(charquery, con);
            reader = cmd.ExecuteReader();
            dtCharacteristic.Load(reader);
            con.Close();
            return dtCharacteristic;

        }
        public OleDbDataReader GetValueCountFromSGData(string filepath,string SubgroupNo,int CharID)
        {
            int id = 0;
            try
            {
               
                string conString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=";
                connection = new OleDbConnection(conString + filepath);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand cmd = new OleDbCommand("select count(value) from SGData where CharID="+CharID+" and SGNO="+SubgroupNo, connection);
                    cmd.Transaction = transaction;
                    reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        reader.Read();
                        //id = reader.GetInt32(0);

                    }
                    transaction.Commit();
                    CloseConnection();
                    return reader;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        return reader;
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                        return reader;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
                 return reader;
            }
        }




       public DataTable GetTraceCatFromTraceCategory(string filepath)
        {
            DataTable dt = new DataTable();
            try
            {
               
                string conString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=";
                connection = new OleDbConnection(conString + filepath);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand cmd = new OleDbCommand("Select TraceCat from Trace group by TraceCat order by min(TraceID)  ", connection);
                    cmd.Transaction = transaction;
                    reader = cmd.ExecuteReader();
                    //if (reader.Read())
                    //{
                    //    TraceCat = reader["TraceCat"].ToString();
                       
                    //}
                    dt.Load(reader);
                    transaction.Commit();
                    CloseConnection();
                    
                    return dt;
                   // return TraceCat;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                       // return TraceCat;
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                       // return TraceCat;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
               // return TraceCat;
            }
            return dt;

        }


        public string GetTraceCatFromTrace(string filepath)
        {
            DataTable dt = new DataTable();
            try
            {

                string conString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=";
                connection = new OleDbConnection(conString + filepath);
                if (connection.State == System.Data.ConnectionState.Closed)
                    connection.Open();

                OleDbTransaction transaction;
                //start transaction
                transaction = connection.BeginTransaction();
                try
                {
                    OleDbCommand cmd = new OleDbCommand("Select TraceCat from Trace where TraceID=1 group by TraceCat order by min(TraceID) ", connection);
                    cmd.Transaction = transaction;
                    reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        TraceCat = reader["TraceCat"].ToString();

                    }
                    transaction.Commit();
                    CloseConnection();
                    //return dt;
                    return TraceCat;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception : " + ex.Message);
                    // Attempt to roll back the transaction.
                    try
                    {
                        // return TraceCat;
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show("Rollback Exception Type: " + ex2.GetType());
                        MessageBox.Show("Message:" + ex2.Message);
                        // return TraceCat;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception : " + e.ToString());
                // return TraceCat;
            }
            return TraceCat;

        }

        private void btnClearFiles_Click(object sender, EventArgs e)
        {
            dtFillDatagridWithFiles.Clear();
            sfdgSPCWBConsolidatedReport.Refresh();
        }
    }
}






    

    


