using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Collections;
using System.Runtime.InteropServices;


namespace AutoZT
{
    class TagListFile
    {

        //fields
        private Excel.Application m_excelObj = null;
        //datatable holds excel data
        private DataTable m_excelData = null;
        //use datatable analogIFixTable to write to CSV file for IFIX script
        private DataTable analogIfixTable = null;
        private string m_pathToExcelDocument = "";
        private Excel.Workbook TagListWorkbook = null;
        private Excel.Sheets sheets = null;
        private Excel.Worksheet worksheet = null;
        private string[] worksheetNames = null;


        /// <summary>
        /// Constructor takes the path of the excel file
        /// </summary>
        /// <param name="pathToExcelDocument">path of the Tag Excel file</param>
        public TagListFile(string pathToExcelDocument)
        {
            m_pathToExcelDocument = pathToExcelDocument;


            // Does the file exist?
            if (!System.IO.File.Exists(m_pathToExcelDocument))
                throw new System.IO.FileNotFoundException();


        }

        /// <summary>
        /// Starts the Excel application; otherwise show an error and exits application.
        /// </summary>
        private void StartExcel()
        {
            m_excelObj = new Excel.Application();
            // See if the Excel Application Object was successfully constructed
            if (m_excelObj == null)
            {
                MessageBox.Show("ERROR: EXCEL couldn't be started!");
                System.Windows.Forms.Application.Exit();
            }
            m_excelObj.Visible = false;

        }
        /*remove this as doing this in the ReadExcelfunction
        /// <summary>
        /// Stops Excel and does the necessary garbage collection.
        /// </summary>
        private void stopExcel()
        {
            //Marshal.ReleaseComObject();
            //Marshal.FinalReleaseComObject
            //close workbook without saving
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
            //TagListWorkbook.Close(false, Type.Missing, Type.Missing);
            //m_excelObj.Application.Quit();
            //m_excelObj.Quit();
            
            
            //don't want save as dialog box to appear - get exception from hresult: 0x800AC472 if I put this in
            //m_excelObj.Application.DisplayAlerts = false;
           // m_excelObj = null;
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
        }
         */
/// <summary>
///  Errors are ignored per Microsoft's suggestion for this type of function:
    // http://support.microsoft.com/default.aspx/kb/317109
/// </summary>
/// <param name="obj"></param>
        private void Release(object obj)
        {
            
            try
            {
                Marshal.FinalReleaseComObject(obj);
            }
            catch { }
        }


        /// <summary>
        /// Reads each from worksheet from C6 to D6 and stores it in a dataTable. Stop the Excel application after the data is read.
        /// </summary>
        public void ReadDataFromExcelToDataTable()
        {
             Excel.Range AllCells = null;
             Excel.Range lastCell = null;
             Excel.Range row = null;

            try
            {
                //start Excel
                StartExcel();                        

                //open workbook
                TagListWorkbook = m_excelObj.Workbooks.Open(m_pathToExcelDocument, 0, true, 5,
                        "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, false, false);

                //create datatable with two columns
                m_excelData = new DataTable("TableData");
                m_excelData.Columns.Add("Tag Name", Type.GetType("System.String"));
                m_excelData.Columns.Add("Address", Type.GetType("System.String"));


                //Set Tag Name column as primary key
                m_excelData.Columns["Tag Name"].Unique = true;
                m_excelData.PrimaryKey = new DataColumn[] { m_excelData.Columns["Tag Name"] };
                
                //Get the sheets 
                sheets = TagListWorkbook.Worksheets;
                //total number of sheets
                int TotalSheets = sheets.Count;
                //instantiate string of worksheet array
                worksheetNames = new string[TotalSheets];

                //Read columns c6 to d6 from all the sheets and add into Data table
                for (int i = 1; i <= TotalSheets; i++)
                {
                    //get each sheet
                    worksheet = (Excel.Worksheet)sheets.get_Item(i);
                    //store sheet names to be used later for database setup
                    worksheetNames[i - 1] = worksheet.Name;

                    AllCells = worksheet.get_Range("C6", "D6");
                    //lastCell = null;                    
                    //Excel.Range AllCells = worksheet.get_Range("C6", "D6");
                    //Excel.Range lastCell = null;
                    
                    //get the last cell for column c6, d6
                    lastCell = AllCells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Missing.Value);

                    //read data starting from column C and row 6 and to D6 to the last cell of these columns
                    for (int j = 6; j <= lastCell.Row; j++)
                    {
                        row = worksheet.get_Range("C" + j.ToString(), "D" + j.ToString());
                        Array strs = (System.Array)row.Cells.Value2;
                        //convert values to array of strings
                        string[] strsArray = ConvertToStringArray(strs);
                        //add the array of strings to the datatable m_excelData
                        m_excelData.Rows.Add(strsArray);
                        //delete 'Blank' row
                        DeleteBlankRowsInDataTable();
                    }
                }               
            }
            
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //Clean up any references
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //clean up Excel objects
                Release(AllCells);
                Release(lastCell);
                Release(row);
                Release(worksheet);
                Release(sheets);
                //Close the workbook
                TagListWorkbook.Close(false, Type.Missing, Type.Missing);
                Release(TagListWorkbook);
                
                //m_excelObj.Application.Quit(); -
                m_excelObj.Quit();
                Release(m_excelObj);
                m_excelObj = null;  

            }                     

        }
        /// <summary>
        /// Add the values from the excel file to an array of string
        /// </summary>
        /// <param name="values">The value to convert to string.</param>
        /// <returns></returns>
        private string[] ConvertToStringArray(System.Array values)
        {
            // create a new string array
            string[] theArray = new string[values.Length];
            string sheetName = worksheet.Name;

            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                {   //label empty cell as Blank
                    theArray[i - 1] = "Blank";

                }
                else
                {
                    theArray[i - 1] = (string)values.GetValue(1, i).ToString();
                    //first column (Tag Name)             
                    if (i == 1)
                    {
                        //capitalize first letter and the letters after the space
                        theArray[i - 1] = PCase(theArray[i - 1]);
                        //eliminate white spaces
                        theArray[i - 1] = theArray[i - 1].Replace(" ", "");
                        //Put sheet name in front of tag names 
                        theArray[i - 1] = sheetName + "." + theArray[i - 1];


                        //if (theArray[i - 1].Length > maxLengthOFTags)
                        //    MessageBox.Show("The tag, " + theArray[i - 1] + " is too long. Please delete a minimum of " + 
                        //        (theArray[i - 1].Length - 30) + " letter(s).");
                    }

                }

            }
            return theArray;
        }


        /// <summary>
        /// Capitalize the first letter and the letters after a space
        /// </summary>
        /// <param name="strParam">String to convert</param>
        /// <returns>String with capitals</returns>
        public static String PCase(String strParam)
        {
            String strProper = "";
            //Capitalize first letter
            strProper = strParam.Substring(0, 1).ToUpper();
            //Store the remainaining characters as they are
            strParam = strParam.Substring(1);
            String strPrev = "";

            for (int iIndex = 0; iIndex < strParam.Length; iIndex++)
            {
                if (iIndex > 1)
                {
                    strPrev = strParam.Substring(iIndex - 1, 1);
                }

                if (strPrev.Equals(" ") || strPrev.Equals("."))//|| strPrev.Equals("\t") || strPrev.Equals("\n") || )
                {
                    strProper += strParam.Substring(iIndex, 1).ToUpper();
                }

                else
                {
                    strProper += strParam.Substring(iIndex, 1);
                }
            }

            return strProper;
        }

        /// <summary>
        /// Create IGS data table.
        /// </summary>
        /// <returns>IGS data table</returns>
        private static DataTable IgsHeadingTable()
        {
            //create table
            DataTable IgsHeadings = new DataTable("IGSTable");
            IgsHeadings.Columns.Add("Data Type", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Respect Data Type", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Client Access", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Scan Rate", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Scaling", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Raw Low", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Raw High", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Scaled Low", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Scaled High", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Scaled Data Type", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Clamp Low", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Clamp High", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Eng Units", Type.GetType("System.String"));
            IgsHeadings.Columns.Add("Description", Type.GetType("System.String"));
            //Old version of IGS driver does not have this column so will remove it. Version 7.40 of IGS does.
           // IgsHeadings.Columns.Add("Negate Value", Type.GetType("System.String"));
           


            return IgsHeadings;
        }

        /// <summary>
        ///Creates the Analog tags Ifix data table. 
        /// </summary>
        /// <returns>returns the data table</returns>
        private static DataTable IfixAnalogHeadingTable()
        {
            //create table
            DataTable ifixAnalogHeadings = new DataTable("IfixTable");
            ifixAnalogHeadings.Columns.Add("!A_NAME", Type.GetType("System.String"));
            //ifixAnalogHeadings.Columns.Add("A_TAG", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_NEXT", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_DESC", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ISCAN", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_SCANT", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_SMOTH", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_IODV", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_IOHT", Type.GetType("System.String"));
            //ifixAnalogHeadings.Columns.Add("A_IOAD", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_IOSC", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ELO", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_EHI", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_EGUDESC", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_IAM", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_IENAB", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ADI", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_LOLO", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_LO", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_HI", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_HIHI", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ROC", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_DBAND", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_PRI", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_EOUT", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_SA1", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_SA2", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_SA3", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA1", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA2", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA3", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA4", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA5", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA6", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA7", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA8", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA9", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA10", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA11", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA12", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA13", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA14", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_AREA15", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ALMEXT1", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ALMEXT2", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ESIGTYPE", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ESIGCONT", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ESIGACK", Type.GetType("System.String"));
            ifixAnalogHeadings.Columns.Add("A_ESIGTRAP!", Type.GetType("System.String"));

            return ifixAnalogHeadings;

        }
        /// <summary>
        /// Create the digital Ifix data table.
        /// </summary>
        /// <returns>The data table.</returns>
        private static DataTable IfixDigitalHeadingTable()
        {
            //create table
            DataTable ifixDigitalHeadings = new DataTable("IfixDigitalTable");
            ifixDigitalHeadings.Columns.Add("!A_NAME", Type.GetType("System.String"));
            //ifixDigitalHeadings.Columns.Add("A_TAG", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_NEXT", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_DESC", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_IODV", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_IOHT", Type.GetType("System.String"));
            //ifixDigitalHeadings.Columns.Add("A_IOAD", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_IAM", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_ISCAN", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_SCANT", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_INV", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_OPENDESC", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_CLOSEDESC", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_IENAB", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_ADI", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_PRI", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_ALMCK", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_EVENT", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_SA1", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_SA2", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_SA3", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_EOUT", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA1", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA2", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA3", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA4", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA5", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA6", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA7", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA8", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA9", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA10", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA11", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA12", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA13", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA14", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_AREA15", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_ALMEXT1", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_ALMEXT2", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_ESIGTYPE", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_ESIGCONT", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_ESIGACK", Type.GetType("System.String"));
            ifixDigitalHeadings.Columns.Add("A_ESIGTRAP!", Type.GetType("System.String"));

            return ifixDigitalHeadings;

        }

        private static DataTable OpcHeadingTable()
        {
            //create table
            DataTable OpcHeadings = new DataTable("OPCTable");
            //OpcHeadings.Columns.Add("OPC Item", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Update Rate", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("OPC AccessPath", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Description", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Point Type", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Reset Value", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Eng. Units", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend Pen Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend Order", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend Data File", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend DB Field", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend High Range", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend Low Range", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend Dec. Places", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend Pen Color", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend Text Color", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Trend Pen Width", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Data Routing", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Data Route OPC Destination", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Data Route OPC AccessPath", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Digital Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Digital Value", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Digital Delay", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Digital Description", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Digital Enable Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Digital Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High High Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High High Value", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High High DeadBand", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High High Delay", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High High Description", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High High Enable Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High High Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High Value", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High DeadBand", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High Delay", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High Description", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High Enable Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog High Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Value", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low DeadBand", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Delay", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Description", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Enable Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Low Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Low Value", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Low DeadBand", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Low Delay", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Low Description", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Low Enable Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Analog Low Low Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Event Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Event Value", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Event Both", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Event Description", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Event Enable Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Event Text Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Sound Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Sound Group", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Sound Repeat Time", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Priority", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Page Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm Log Enable", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Alarm File Name", Type.GetType("System.String"));
            OpcHeadings.Columns.Add("Append Date to Alarm File Name", Type.GetType("System.String"));

            return OpcHeadings;

        }

        /// <summary>
        /// Appends the data table to the file specified.
        /// </summary>
        /// <param name="pathAndFileName">Create the file and appends data if it does not exist</param>
        /// <param name="theDataTable">The data table to write to file</param>
        /// <param name="heading">Flag to write heading of data table</param>
        private static void WriteDataTableToCsvFile(string pathAndFileName, DataTable theDataTable, bool heading, bool newFile)
        {
            //fields
            string separator = ",";
            //store maximum number of columns
            int iColCount = theDataTable.Columns.Count;
            //create or append new data to file
            StreamWriter sw = new StreamWriter(pathAndFileName, newFile);

            //write column names if heading is true
            if (heading)
            {
                for (int i = 0; i < iColCount; i++)
                {
                    sw.Write(theDataTable.Columns[i]);
                    if (i < iColCount - 1)
                    {
                        sw.Write(separator);
                    }
                }
                sw.Write(sw.NewLine);
            }


            //write the rest of the rows

            foreach (DataRow dr in theDataTable.Rows)
            {
                for (int i = 0; i < iColCount; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string data = dr[i].ToString();
                        sw.Write(data);
                    }
                    if (i < iColCount - 1)
                    {
                        sw.Write(separator);
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();

        }


        /// <summary>
        /// Deletes the Tag Name row that is blank.
        /// </summary>        
        private void DeleteBlankRowsInDataTable()
        {
            DataRow[] getBlankRows = null;

            //find the empty rows in Tag Name column
            getBlankRows = FindRowsInDataTable(m_excelData, "Blank");

            foreach (DataRow dr in getBlankRows)
            {
                //removed the empty rows
                m_excelData.Rows.Remove(dr);

            }
        }

        /// <summary>
        /// Creates a CSV file
        /// </summary>
        /// <param name="absolutePathAndFileName">The file to append the data to.</param>
        public void SaveIGSDataTableToCsvFile(string absolutePathAndFileName)
        {
            //variables

            DataTable igsTable = m_excelData.Copy();
            //merge IGS table into the Excel table 

            igsTable.Merge(IgsHeadingTable(), false, MissingSchemaAction.Add);

            //fill the rest of the columns        

            foreach (DataRow setRow in igsTable.Rows)
            {
                setRow["Tag Name"] = "\"" + setRow["Tag Name"] + "\"";
                //check the 4th character to see if it is a period to indicate a scope in plc
                if (setRow["Address"].ToString().Substring(3, 1) == ".") 
                {
                    //Set Address column with PROGRAM in front
                    setRow["Address"] = "\"" + "PROGRAM:" + setRow[1] + "\"";
                }

                //Set DataType column 
                setRow["Data Type"] = "Float";
                //Set Respect Data type
                setRow["Respect Data Type"] = 1;
                //Set Client Access
                setRow["Client Access"] = "R/W";
                //Set Scan Rate
                setRow["Scan Rate"] = 100;
                setRow["Scaling"] = "";
                setRow["Raw Low"] = "";
                setRow["Raw High"] = "";
                setRow["Scaled Low"] = "";
                setRow["Scaled High"] = "";
                setRow["Scaled Data Type"] = "";
                setRow["Clamp Low"] = "";
                setRow["Clamp High"] = "";
                setRow["Eng Units"] = "";
                setRow["Description"] = "\"\"";
               //old version of IGS driver does not have this column so will remove it.
                // setRow["Negate Value"] = "";
              

            }
            //set Data Ready rows Data Type to Boolean          
            SetRowValueOfDataTable(igsTable, "Tag Name", "ready", "Data Type", "Boolean");

            //write igstable to csv file
            WriteDataTableToCsvFile(absolutePathAndFileName, igsTable, true, false);

        }
        public void SaveOPCDataTableToCsvFile(string absolutePathAndFileName, string channelName, string plcName)
        {
            //variables
            int endPoint = 0;

            DataTable opcTable = m_excelData.Copy();
            //merge IGS table into the Excel table 

            //remove address column that we don't need
            opcTable.Columns.RemoveAt(1);
            
            //merge tables
            opcTable.Merge(OpcHeadingTable(), false, MissingSchemaAction.Add);

            //change first column name
            opcTable.Columns["Tag Name"].ColumnName = "OPC Item";

            //DataColumn excelColumn = opcTable.Columns["Tag Name"];
            //excelColumn.SetOrdinal(2);
            //excelColumn.ColumnName = "OPC Item";
          
            //fill the rest of the columns        

            foreach (DataRow setRow in opcTable.Rows)
            {
                //number not in quotes for the OPC csv file       
                setRow["Update Rate"] = 1000;
                setRow["OPC AccessPath"] = "";
                setRow["Description"] = "";
                setRow["Point Type"] = 0;
                setRow["Reset Value"] = "#FALSE#";
                setRow["Eng. Units"] = "";
                setRow["Trend Enable"] = "#TRUE#";
                setRow["Trend Pen Group"] = "";
                setRow["Trend Order"] = 0;
                if ((setRow["OPC Item"].ToString().Substring(0, 6) == "Common") || (setRow["OPC Item"].ToString().Substring(0, 6) == "common") || (setRow["OPC Item"].ToString().Substring(0, 3) == "CMN") || (setRow["OPC Item"].ToString().Substring(0, 3) == "cmn"))
                {
                    setRow["Trend Data File"] = "\"CMN\"";
                }
                else if ((setRow["OPC Item"].ToString().Substring(0, 5) == "Daily") || (setRow["OPC Item"].ToString().Substring(0, 5) == "daily"))
                {
                    setRow["Trend Data File"] = "\"Daily\"";
                }
                else if ((setRow["OPC Item"].ToString().Substring(0, 2) == "MC") || (setRow["OPC Item"].ToString().Substring(0, 2) == "mc"))
                {
                    setRow["Trend Data File"] = "\"MC\"";
                }
                else if ((setRow["OPC Item"].ToString().Substring(0, 2) == "RC") || (setRow["OPC Item"].ToString().Substring(0, 2) == "rc"))
                {
                    setRow["Trend Data File"] = "\"RC\"";
                }
                else if ((setRow["OPC Item"].ToString().Substring(0, 2) == "ZW") || (setRow["OPC Item"].ToString().Substring(0, 2) == "zw"))
                {
                    endPoint = setRow["OPC Item"].ToString().IndexOf(".");
                    setRow["Trend Data File"] = "\"" + setRow["OPC Item"].ToString().Substring(0, endPoint).ToUpper() + "\"";                
                }
                else if ((setRow["OPC Item"].ToString().Substring(0, 3) == "MIT") || (setRow["OPC Item"].ToString().Substring(0, 3) == "mit"))
                {
                    endPoint = setRow["OPC Item"].ToString().IndexOf(".");
                    setRow["Trend Data File"] = "\"" + setRow["OPC Item"].ToString().Substring(0, endPoint).ToUpper() + "\"";
                }
                setRow["Trend DB Field"] = "";
                setRow["Trend High Range"] = 100;
                setRow["Trend Low Range"] = 0;
                setRow["Trend Dec. Places"] = 3;
                setRow["Trend Pen Color"] = 0;
                setRow["Trend Text Color"] = 0;
                setRow["Trend Pen Width"] = 1;
                setRow["Data Routing"] = "#FALSE#";
                setRow["Data Route OPC Destination"] = "";
                setRow["Data Route OPC AccessPath"] = "";
                setRow["Alarm Digital Enable"] = "#FALSE#";
                setRow["Alarm Digital Value"] = "#TRUE#";
                setRow["Alarm Digital Delay"] = 0;
                setRow["Alarm Digital Description"] = "\"Digital Alarm\"";
                setRow["Alarm Digital Enable Text Group"] = "#FALSE#";		
                setRow["Alarm Digital Text Group"] = "";
                setRow["Alarm Analog High High Enable"] = "#FALSE#";
                setRow["Alarm Analog High High Value"] = 100;
                setRow["Alarm Analog High High DeadBand"] = 0.1;
                setRow["Alarm Analog High High Delay"] = 0;          						
                setRow["Alarm Analog High High Description"] = "\"High High Alarm\"";
                setRow["Alarm Analog High High Enable Text Group"] = "#FALSE#";
                setRow["Alarm Analog High High Text Group"] = "\"Disabled\"";
                setRow["Alarm Analog High Enable"] = "#FALSE#";
                setRow["Alarm Analog High Value"] = 100;
                setRow["Alarm Analog High DeadBand"] = 0.1;
                setRow["Alarm Analog High Delay"] = 0;
                setRow["Alarm Analog High Description"] = "\"High Alarm\"";
                setRow["Alarm Analog High Enable Text Group"] = "#FALSE#";
                setRow["Alarm Analog High Text Group"] = "\"Disabled\"";
                setRow["Alarm Analog Low Enable"] = "#FALSE#";
                setRow["Alarm Analog Low Value"] = 0;
                setRow["Alarm Analog Low DeadBand"] = 0.1;
                setRow["Alarm Analog Low Delay"] = 0;
                setRow["Alarm Analog Low Description"] = "\"Low Alarm\"";
                setRow["Alarm Analog Low Enable Text Group"] = "#FALSE#";
                setRow["Alarm Analog Low Text Group"] = "\"Disabled\"";
                setRow["Alarm Analog Low Low Enable"] = "#FALSE#";
                setRow["Alarm Analog Low Low Value"] = 0;
                setRow["Alarm Analog Low Low DeadBand"] = 0.1;
                setRow["Alarm Analog Low Low Delay"] = 0;
                setRow["Alarm Analog Low Low Description"] = "\"Low Low Alarm\"";
                setRow["Alarm Analog Low Low Enable Text Group"] = "#FALSE#";
                setRow["Alarm Analog Low Low Text Group"] = "\"Disabled\"";
                setRow["Event Enable"] = "#FALSE#";
                setRow["Event Value"] = "#TRUE#";
                setRow["Event Both"] = "#FALSE#";
                setRow["Event Description"] = "\"Event\"";
                setRow["Event Enable Text Group"] = "#FALSE#";
                setRow["Event Text Group"] = "\"Disabled\"";
                setRow["Alarm Sound Enable"] = "#FALSE#";
                setRow["Alarm Sound Group"] = "";
                setRow["Alarm Sound Repeat Time"] = 0;
                setRow["Alarm Priority"] = 0;
                setRow["Alarm Page Enable"] = "#FALSE#";
                setRow["Alarm Log Enable"] = "#FALSE#";
                setRow["Alarm File Name"] = "\"Alarm\"";
                setRow["Append Date to Alarm File Name"] = "#TRUE#";

            }

            foreach (DataRow setRow in opcTable.Rows)
            {
                //add gateway server to the tag name of IGS tag
                setRow["OPC Item"] = "\"" + "Intellution.IntellutionGatewayOPCServer\\" + channelName + "." + plcName + "." + setRow["OPC Item"] + "\"";
            }

            //set data ready tags to digital
            SetRowValueOfDataTable(opcTable, "OPC Item", "ready", "Point Type", "1");

            //set trend enable to false for data ready tags
            SetRowValueOfDataTable(opcTable,"OPC Item", "ready", "Trend Enable", "#FALSE#");
            
            //set Trend data file rows to blank as we don't want to log data ready bits
            SetRowValueOfDataTable(opcTable, "OPC Item", "ready", "Trend Data File", "");
                     
            //write igstable to csv file
            WriteDataTableToCsvFile(absolutePathAndFileName, opcTable, true, false);

        }


        /// <summary>
        /// Sets the row value in the table to one specified
        /// </summary>
        /// <param name="searchTable">The table to search</param>
        /// <param name="setColumnName">/The column to update</param>
        /// <param name="findRowValue">The string to search for</param>
        /// <param name="setRowValue">The value to set to</param>
        private void SetRowValueOfDataTable(DataTable searchTable, string findColumnName, string findRowValue, string setColumnName, string setRowValue) //, string expression)
        {
            DataRow[] foundRows;
            string str = "[" + findColumnName + "]" + " " + "LIKE '*" + findRowValue + "*'";
            foundRows = searchTable.Select(str);

            foreach (DataRow row in foundRows)
            {

                row[setColumnName] = setRowValue;

            }

        }
        /// <summary>
        /// Find the value in the Tag Name column of the Data Table
        /// </summary>
        /// <param name="searchTable">The data table to search</param>
        /// <param name="findRowValue">The value to search for</param>
        /// <returns></returns>
        private DataRow[] FindRowsInDataTable(DataTable searchTable, string findRowValue) //, string expression)
        {
            DataRow[] findRows;
            string findString = "[" + "Tag Name" + "]" + " " + "LIKE '*" + findRowValue + "*'";
            findRows = searchTable.Select(findString);

            return findRows;
        }

        /// <summary>
        /// Create four data tables to hold the parts of the IFIX file. Call write method to print data tables to CSV file.
        /// </summary>
        /// <param name="absolutePathAndFileName">The file to write the data tables to</param>
        /// <param name="ifixDatabaseName">The name of the IFIX database file</param>
        /// <param name="driverName">The name of the driver used</param>
        /// <param name="channelName">The name of the IGS channel</param>
        /// <param name="plcName">The name of the IGS plc</param>
        public void SaveIfixDataTableToCsvFile(string absolutePathAndFileName, string ifixDatabaseName, string driverName, string channelName, string plcName)
        {
            //variables

            string date = DateTime.Now.ToShortDateString();
            string time = DateTime.Now.ToShortTimeString();

            //copy data from the excel datatable
            analogIfixTable = m_excelData.Copy();

            //merge the ifix heading table into the analogIfixTable
            analogIfixTable.Merge(IfixAnalogHeadingTable(), false, MissingSchemaAction.Add);

            //delete rows that have digital tags
            DataRow[] getDigitalRows = null;
            //DataRow delDigitalRow = null;

            //find the digital rows that have "ready"
            getDigitalRows = FindRowsInDataTable(analogIfixTable, "ready");

            foreach (DataRow dr in getDigitalRows)
            {
                //removed the data ready rows since these are digital tags
                analogIfixTable.Rows.Remove(dr);

            }


            //change order of columns to match proper ifix order
            DataColumn excelColumn = analogIfixTable.Columns["Tag Name"];
            excelColumn.SetOrdinal(2);
            excelColumn.ColumnName = "A_TAG";
            excelColumn = analogIfixTable.Columns["Address"];
            excelColumn.SetOrdinal(9);
            excelColumn.ColumnName = "A_IOAD";


            //fill analog analogIfixTable 
            foreach (DataRow setRow in analogIfixTable.Rows)
            {
                if (driverName == "IGS")
                {
                    //Set A_IOAD column to A_Tag as this is name ifix needs from IGS
                    setRow["A_IOAD"] = "\"" + channelName + "." + plcName + "." + setRow["A_TAG"] + "\"";
                }
                else if (driverName == "GE9")
                {           
                    //Set A_IOAD column to A_Tag as this is name ifix needs from IGS
                    setRow["A_IOAD"] = "\"" + plcName + ":" + setRow["A_IOAD"] + "\"";
               
                }
                //add Modicon code here

                //Need to make length of tags smaller as IFIX has max length of 30 characters
                setRow["A_TAG"] = setRow["A_TAG"].ToString().Replace("Common", "CMN");
                //need the '.' so doesn't rename tags that have daily in it to DA
                setRow["A_TAG"] = setRow["A_TAG"].ToString().Replace("Daily.", "DA.");
                //Set A_TAG periods to underscores as period is not allowed in IFIX
                setRow["A_TAG"] = setRow["A_TAG"].ToString().Replace(".", "_");
              
                if (setRow["A_TAG"].ToString().Length > 30)
                    MessageBox.Show("The tag, " + setRow["A_TAG"].ToString() + " is too long. Please delete a minimum of " +
                                (setRow["A_TAG"].ToString().Length - 30) + " letter(s). Otherwise the IFIX file will not import properly.");

                //Change to UPPERCASE
                setRow["A_TAG"] = "\"" + setRow["A_TAG"].ToString().ToUpper() + "\"";


                
                //Fill in the remaining columns 
                setRow["!A_NAME"] = "\"AI\"";
                setRow["A_NEXT"] = "\"\"";
                setRow["A_DESC"] = "\"\"";
                setRow["A_ISCAN"] = "\"ON\"";
                setRow["A_SCANT"] = "\"1\"";
                setRow["A_SMOTH"] = "\"0\"";
                setRow["A_IODV"] = "\"" + driverName + "\"";
                setRow["A_IOHT"] = "\"\"";
                setRow["A_IOSC"] = "\"None\"";
                setRow["A_ELO"] = "\"0.00000E+00\"";
                setRow["A_EHI"] = "\"1.00000E+15\"";
                setRow["A_EGUDESC"] = "\"\"";
                setRow["A_IAM"] = "\"AUTO\"";
                setRow["A_IENAB"] = "\"DISABLE\"";
                setRow["A_ADI"] = "\"NONE\"";
                setRow["A_LOLO"] = "\"0.00000E+00\"";
                setRow["A_LO"] = "\"0.00000E+00\""; 
                setRow["A_HI"] = "\"1.00000E+15\"";
                setRow["A_HIHI"] = "\"1.00000E+15\"";
                setRow["A_ROC"] = "\"0.00000E+00\"";
                setRow["A_DBAND"] = "\"0.00000E+00\"";
                setRow["A_PRI"] = "\"LOW\"";
                setRow["A_EOUT"] = "\"NO\"";
                setRow["A_SA1"] = "\"NONE\"";
                setRow["A_SA2"] = "\"NONE\"";
                setRow["A_SA3"] = "\"NONE\"";
                setRow["A_AREA1"] = "\"ALL\"";
                setRow["A_AREA2"] = "\"\"";
                setRow["A_AREA3"] = "\"\"";
                setRow["A_AREA4"] = "\"\"";
                setRow["A_AREA5"] = "\"\"";
                setRow["A_AREA6"] = "\"\"";
                setRow["A_AREA7"] = "\"\"";
                setRow["A_AREA8"] = "\"\"";
                setRow["A_AREA9"] = "\"\"";
                setRow["A_AREA10"] = "\"\"";
                setRow["A_AREA11"] = "\"\"";
                setRow["A_AREA12"] = "\"\"";
                setRow["A_AREA13"] = "\"\"";
                setRow["A_AREA14"] = "\"\"";
                setRow["A_AREA15"] = "\"\"";
                setRow["A_ALMEXT1"] = "\"\"";
                setRow["A_ALMEXT2"] = "\"\"";
                setRow["A_ESIGTYPE"] = "\"NONE\"";
                setRow["A_ESIGCONT"] = "\"YES\"";
                setRow["A_ESIGACK"] = "\"NO\"";
                setRow["A_ESIGTRAP!"] = "\"REJECT\"";


            }

            //Set TMP tags low range to -1.00000E+5
            SetRowValueOfDataTable(analogIfixTable, "A_TAG", "tmp", "A_ELO", "\"-1.00000E+5\"");
            SetRowValueOfDataTable(analogIfixTable, "A_TAG", "tmp", "A_LOLO", "\"-1.00000E+5\"");
            SetRowValueOfDataTable(analogIfixTable, "A_TAG", "tmp", "A_LO", "\"-1.00000E+5\"");

            //Set pressures to negative
            SetRowValueOfDataTable(analogIfixTable, "A_TAG", "pressure", "A_ELO", "\"-1.00000E+5\"");
            SetRowValueOfDataTable(analogIfixTable, "A_TAG", "pressure", "A_LOLO", "\"-1.00000E+5\"");
            SetRowValueOfDataTable(analogIfixTable, "A_TAG", "pressure", "A_LO", "\"-1.00000E+5\"");


            DataTable digitalTable = m_excelData.Clone();
            digitalTable.Columns["Tag Name"].AllowDBNull = true;
            DataRow[] getRows = null;
            DataRow addRow = null;

            //find the digital rows that have "ready"
            getRows = FindRowsInDataTable(m_excelData, "ready");

            //max number of columns 
            int numColumns = digitalTable.Columns.Count;

            foreach (DataRow dr in getRows)
            {
                //add new table row before assigning values
                addRow = digitalTable.NewRow();
                //copy every column value from getRows to addrow
                for (int index = 0; index < numColumns; index++)
                    addRow[index] = dr[index];
                //add rows to digital table
                digitalTable.Rows.Add(addRow);

            }
            //merge datatable to add columns from ifixdigitalheading table
            digitalTable.Merge(IfixDigitalHeadingTable(), false, MissingSchemaAction.Add);

            //change order of columns to match proper ifix order
            excelColumn = null;
            excelColumn = digitalTable.Columns["Tag Name"];
            excelColumn.SetOrdinal(2);
            excelColumn.ColumnName = "A_TAG";
            excelColumn = digitalTable.Columns["Address"];
            excelColumn.SetOrdinal(9);
            excelColumn.ColumnName = "A_IOAD";

            //add digital tag for GE9 driver

            //fill digital IfixTable 
            foreach (DataRow setRow in digitalTable.Rows)
            {
                if (driverName == "IGS")
                {
                    //Set A_IOAD column to A_Tag as this is name ifix needs from IGS
                    setRow["A_IOAD"] = "\"" + channelName + "." + plcName + "." + setRow["A_TAG"] + "\"";
                }
                else if (driverName == "GE9")
                {
                    //Set A_IOAD column to A_Tag as this is name ifix needs from IGS
                    setRow["A_IOAD"] = "\"" + plcName + ":" + setRow["A_IOAD"] + "\"";
                    
                }
                //Set A_TAG periods to underscores as period is not allowed in IFIX
                setRow["A_TAG"] = setRow["A_TAG"].ToString().Replace(".", "_");
              
                if (setRow["A_TAG"].ToString().Length > 30)
                    MessageBox.Show("The tag, " + setRow["A_TAG"].ToString() + " is too long. Please delete a minimum of " +
                                (setRow["A_TAG"].ToString().Length - 30) + " letter(s). Otherwise the IFIX file will not import properly.");

                //Change to UPPERCASE
                setRow["A_TAG"] = "\"" + setRow["A_TAG"].ToString().ToUpper() + "\"";

                //Fill in the remaining columns 
                setRow["!A_NAME"] = "\"DI\"";
                setRow["A_NEXT"] = "\"\"";
                setRow["A_DESC"] = "\"\"";
                setRow["A_IODV"] = "\"" + driverName + "\"";
                setRow["A_IOHT"] = "\"\"";
                setRow["A_IAM"] = "\"AUTO\"";
                setRow["A_ISCAN"] = "\"ON\"";
                setRow["A_SCANT"] = "\"1\"";
                setRow["A_INV"] = "\"NO\"";
                setRow["A_OPENDESC"] = "\"OPEN\"";
                setRow["A_CLOSEDESC"] = "\"CLOSE\"";
                setRow["A_IENAB"] = "\"DISABLE\"";
                setRow["A_ADI"] = "\"NONE\"";
                setRow["A_PRI"] = "\"LOW\"";
                setRow["A_ALMCK"] = "\"COS\"";
                setRow["A_EVENT"] = "\"DISABLE\"";
                setRow["A_SA1"] = "\"NONE\"";
                setRow["A_SA2"] = "\"NONE\"";
                setRow["A_SA3"] = "\"NONE\"";
                setRow["A_EOUT"] = "\"NO\"";
                setRow["A_AREA1"] = "\"ALL\"";
                setRow["A_AREA2"] = "\"\"";
                setRow["A_AREA3"] = "\"\"";
                setRow["A_AREA4"] = "\"\"";
                setRow["A_AREA5"] = "\"\"";
                setRow["A_AREA6"] = "\"\"";
                setRow["A_AREA7"] = "\"\"";
                setRow["A_AREA8"] = "\"\"";
                setRow["A_AREA9"] = "\"\"";
                setRow["A_AREA10"] = "\"\"";
                setRow["A_AREA11"] = "\"\"";
                setRow["A_AREA12"] = "\"\"";
                setRow["A_AREA13"] = "\"\"";
                setRow["A_AREA14"] = "\"\"";
                setRow["A_AREA15"] = "\"\"";
                setRow["A_ALMEXT1"] = "\"\"";
                setRow["A_ALMEXT2"] = "\"\"";
                setRow["A_ESIGTYPE"] = "\"NONE\"";
                setRow["A_ESIGCONT"] = "\"YES\"";
                setRow["A_ESIGACK"] = "\"NO\"";
                setRow["A_ESIGTRAP!"] = "\"REJECT\"";


            }

            //clone digital table

            DataTable bottomTable = digitalTable.Clone();
            //when the clone was done this tag became a primary key need to allow nulls since the will be a blank line.
            bottomTable.Columns["A_TAG"].AllowDBNull = true;
            //the bottom table will be strings 
            foreach (DataColumn dc in bottomTable.Columns)
            {
                dc.DataType = Type.GetType("System.String");

            }

            //create new row
            DataRow bottomRows = null;

            //add blank row
            bottomTable.Rows.Add(bottomTable.NewRow());

            //add next row
            bottomRows = bottomTable.NewRow();
            //set each row with the appropiate column value
            bottomRows[0] = "[BLOCK TYPE";
            bottomRows[1] = "TAG";
            bottomRows[2] = "NEXT BLOCK";
            bottomRows[3] = "DESCRIPTION";
            bottomRows[4] = "I/O DEVICE";
            bottomRows[5] = "H/W OPTIONS";
            bottomRows[6] = "I/O ADDRESS";
            bottomRows[7] = "INITIAL A/M STATUS";
            bottomRows[8] = "INITIAL SCAN";
            bottomRows[9] = "SCAN TIME";
            bottomRows[10] = "INVERT OUTPUT";
            bottomRows[11] = "OPEN TAG"; ;
            bottomRows[12] = "CLOSE TAG";
            bottomRows[13] = "ALARM ENABLE";
            bottomRows[14] = "ALARM AREA(S)";
            bottomRows[15] = "ALARM PRIORITY";
            bottomRows[16] = "ALARM TYPE";
            bottomRows[17] = "EVENT MESSAGES";
            bottomRows[18] = "SECURITY AREA 1";
            bottomRows[19] = "SECURITY AREA 2";
            bottomRows[20] = "SECURITY AREA 3";
            bottomRows[21] = "ENABLE OUTPUT";
            bottomRows[22] = "ALARM AREA 1";
            bottomRows[23] = "ALARM AREA 2";
            bottomRows[24] = "ALARM AREA 3";
            bottomRows[25] = "ALARM AREA 4";
            bottomRows[26] = "ALARM AREA 5";
            bottomRows[27] = "ALARM AREA 6";
            bottomRows[28] = "ALARM AREA 7";
            bottomRows[29] = "ALARM AREA 8";
            bottomRows[30] = "ALARM AREA 9";
            bottomRows[31] = "ALARM AREA 10";
            bottomRows[32] = "ALARM AREA 11";
            bottomRows[33] = "ALARM AREA 12";
            bottomRows[34] = "ALARM AREA 13";
            bottomRows[35] = "ALARM AREA 14";
            bottomRows[36] = "ALARM AREA 15";
            bottomRows[37] = "USER FIELD 1";
            bottomRows[38] = "USER FIELD 2";
            bottomRows[39] = "ESIG TYPE";
            bottomRows[40] = "ESIG ALLOW CONT USE";
            bottomRows[41] = "ESIG XMPT ALARM ACK";
            bottomRows[42] = "ESIG UNSIGNED WRITES]";

            //add row to table
            bottomTable.Rows.Add(bottomRows);


            //clone IFix Heading table
            // DataTable topTable = IfixAnalogHeadingTable();
            DataTable topTable = analogIfixTable.Clone();
            //when the clone was done this tag became a primary key need to allow nulls since the will be a blank line.
            topTable.Columns["A_TAG"].AllowDBNull = true;
            //the top table will be strings 
            foreach (DataColumn dc in topTable.Columns)
            {
                dc.DataType = Type.GetType("System.String");

            }

            //create new row
            DataRow topRows = null;
            topRows = topTable.NewRow();

            //Set first row
            topRows[0] = "[NodeName : FIX,Database : " + ifixDatabaseName;
            topRows[1] = "File Name : C:\\Program Files\\GE Fanuc\\Proficy iFIX\\PDB\\" + ifixDatabaseName + ".csv";
            topRows[2] = "Date : " + date;
            topRows[3] = "Time : " + time + "]";
            //add row to table
            topTable.Rows.Add(topRows);
            //add blank row
            topTable.Rows.Add(topTable.NewRow());

            //add next row
            topRows = topTable.NewRow();


            topRows[0] = "[BLOCK TYPE";
            topRows[1] = "TAG";
            topRows[2] = "NEXT BLK";
            topRows[3] = "DESCRIPTION";
            topRows[4] = "INITIAL SCAN";
            topRows[5] = "SCAN TIME";
            topRows[6] = "SMOOTHING";
            topRows[7] = "I/O DEVICE";
            topRows[8] = "H/W OPTIONS";
            topRows[9] = "I/O ADDRESS";
            topRows[10] = "SIGNAL CONDITIONING";
            topRows[11] = "LOW EGU LIMIT";
            topRows[12] = "HIGH EGU LIMIT"; ;
            topRows[13] = "EGU TAG";
            topRows[14] = "INITIAL A/M STATUS";
            topRows[15] = "ALARM ENABLE";
            topRows[16] = "ALARM AREA(S)";
            topRows[17] = "LO LO ALARM LIMIT";
            topRows[18] = "LO ALARM LIMIT";
            topRows[19] = "HI ALARM LIMIT";
            topRows[20] = "HI HI ALARM LIMIT";
            topRows[21] = "ROC ALARM LIMIT";
            topRows[22] = "DEAD BAND";
            topRows[23] = "ALARM PRIORITY";
            topRows[24] = "ENABLE OUTPUT";
            topRows[25] = "SECURITY AREA 1";
            topRows[26] = "SECURITY AREA 2";
            topRows[27] = "SECURITY AREA 3";
            topRows[28] = "ALARM AREA 1";
            topRows[29] = "ALARM AREA 2";
            topRows[30] = "ALARM AREA 3";
            topRows[31] = "ALARM AREA 4";
            topRows[32] = "ALARM AREA 5";
            topRows[33] = "ALARM AREA 6";
            topRows[34] = "ALARM AREA 7";
            topRows[35] = "ALARM AREA 8";
            topRows[36] = "ALARM AREA 9";
            topRows[37] = "ALARM AREA 10";
            topRows[38] = "ALARM AREA 11";
            topRows[39] = "ALARM AREA 12";
            topRows[40] = "ALARM AREA 13";
            topRows[41] = "ALARM AREA 14";
            topRows[42] = "ALARM AREA 15";
            topRows[43] = "USER FIELD 1";
            topRows[44] = "USER FIELD 2";
            topRows[45] = "ESIG TYPE";
            topRows[46] = "ESIG ALLOW CONT USE";
            topRows[47] = "ESIG XMPT ALARM ACK";
            topRows[48] = "ESIG UNSIGNED WRITES]";

            //add row to table
            topTable.Rows.Add(topRows);

            //write toptable to csv
            WriteDataTableToCsvFile(absolutePathAndFileName, topTable, false, false);
            WriteDataTableToCsvFile(absolutePathAndFileName, analogIfixTable, true, true);
            WriteDataTableToCsvFile(absolutePathAndFileName, bottomTable, false, true);
            WriteDataTableToCsvFile(absolutePathAndFileName, digitalTable, true, true);

        }
        /// <summary>
        /// Writes tags for 
        /// </summary>
        /// <param name="absolutePathAndFileName"></param>
        public void WriteIfixScriptToTextFile(string absolutePathAndFileName)
        {
                 
            DataTable ifixScriptDataTable = analogIfixTable.Copy();
            string beginString = "            & .";
            string endString = ".F_CV & \",\" _";
            string lastTagEndString = ".F_CV";
            StreamWriter sw = File.CreateText(absolutePathAndFileName);

            BitArray temperatureExists = new BitArray(4, false);
            BitArray commonTags = new BitArray(3, false);
            string firstZWWorksheetName = "";

            int startPoint = 0;
            int totalCommonRows = 0;
            int totalDailyRows = 0;
            int totalTrainRows = 0;
            int totalMit1Rows = 0;
            int totalMitRows = 0;
            int totalTrain1Rows = 0;
            int startListCommon = 0;
            int endListCommon = 0;
            int startListDaily = 0;
            int endListDaily = 0;
            int startListTrain = 0;
            int endListTrain = 0;
            int startListMit = 0;
            int endListMit = 0;
            int startListMit1 = 0;
            int endListMit1 = 0;
            int startListTrain1 = 0;
            int endListTrain1 = 0;
            int startListMC = 0;
            int endListMC = 0;
            int totalMCRows = 0;
            int startListMC1 = 0;
            int endListMC1 = 0;
            int totalMC1Rows = 0;
            int startListRC = 0;
            int endListRC = 0;
            int totalRCRows = 0;
            int startListRC1 = 0;
            int endListRC1 = 0;
            int totalRC1Rows = 0;
            
            

            //rename tag name column
            ifixScriptDataTable.Columns["A_TAG"].ColumnName = "Tag Name";
            //remove the "" that were used in ifix table from tag name column
            foreach (DataRow setRow in ifixScriptDataTable.Rows)
            {
                setRow["Tag Name"] = setRow["Tag Name"].ToString().Replace("\"", "");

            }

            //get start and end indexes for the tags
            //GetStartEndIndexesForTagsInDataTable(ref ifixScriptDataTable, ref startListCommon, ref endListCommon, ref totalCommonRows, ref startListDaily, ref endListDaily, ref totalDailyRows, ref startListTrain, ref endListTrain, ref totalTrainRows, ref startListTrain1, ref endListTrain1, ref totalTrain1Rows, ref startListMit, ref endListMit, ref totalMitRows, ref startListMit1, ref endListMit1, ref totalMit1Rows, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,1);
           GetStartEndIndexesForTagsInDataTable(ref ifixScriptDataTable, ref startListCommon, ref endListCommon, ref totalCommonRows, ref startListDaily, ref endListDaily, ref totalDailyRows, ref startListTrain, ref endListTrain, ref totalTrainRows, ref startListTrain1, ref endListTrain1, ref totalTrain1Rows, ref startListMit, ref endListMit, ref totalMitRows, ref startListMit1, ref endListMit1, ref totalMit1Rows, ref startListMC, ref endListMC, ref totalMCRows, ref startListMC1, ref endListMC1, ref totalMC1Rows, ref startListRC, ref endListRC, ref totalRCRows, ref startListRC1, ref endListRC1, ref totalRC1Rows);

            //get first 'ZW' worksheet name from excel file
            firstZWWorksheetName = SearchFirstNameOfWorkSheet("ZW");

            /*
            //get first 'ZW' worksheet name from excel file
            for (int j = 0; j < worksheetNames.Length; j++)
            {
                if (worksheetNames[j].ToString().Contains("ZW"))
                {
                    firstZWWorksheetName = worksheetNames[j].ToString();
                    break;
                }
            }
            */

            //check if there is permeate or feed temperature in common and put in train
            for (int i = startListCommon; i <= endListCommon; i++)
            {
                if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("PERMEATETEMPERATURE"))
                {
                    temperatureExists[0] = true;
                }
                else if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("FEEDTEMPERATURE"))
                {
                    temperatureExists[1] = true;
                }
                else if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("PERMEATEFLOW" + firstZWWorksheetName) || (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains(firstZWWorksheetName + "PERMEATEFLOW")))
                {
                    commonTags[0] = true;
                }
                else if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("PERMEATETURBIDITY" + firstZWWorksheetName) || (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains(firstZWWorksheetName + "PERMEATETURBIDITY")))
                {
                    commonTags[1] = true;
                }
                else if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("MEMBRANETANKLEVEL" + firstZWWorksheetName) || (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains(firstZWWorksheetName + "MEMBRANETANKLEVEL")))
                {
                    commonTags[2] = true;
                }
            }


            //check if temperature exists in trains already
            for (int i = startListTrain1; i <= endListTrain1; i++)
            {
                if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("PERMEATETEMPERATURE"))
                {
                    temperatureExists[2] = true;
                }
                else if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("FEEDTEMPERATURE"))
                {
                    temperatureExists[3] = true;
                }


            }

            //Write script if there is a common tab in the Excel file
            if (totalCommonRows > 0)
            {
                sw.WriteLine("Private Sub Common_OnTimeOut(ByVal lTimerId As Long)");
                sw.WriteLine("'Declarations");
                sw.WriteLine("Dim FileName As String");
                sw.WriteLine("Dim RecordString As String");
                sw.WriteLine("Dim FileNoOut As Integer");
                sw.Write(sw.NewLine);
                sw.WriteLine("'variables for the directory structure");
                sw.WriteLine("Dim rootPath, dirPath");
                sw.Write(sw.NewLine);
                sw.WriteLine("'If there is an error go to the next line");
                sw.WriteLine("On Error Resume Next");
                sw.Write(sw.NewLine);
                sw.WriteLine("'root path");
                sw.WriteLine("rootPath = \"C:\\ZenoTrac\\\"");
                sw.Write(sw.NewLine);
                sw.WriteLine("'set directory path");
                sw.WriteLine("dirPath = \"C:\\ZenoTrac\\Common\\\"");
                sw.Write(sw.NewLine);
                sw.WriteLine("Set objFile = CreateObject(\"Scripting.FileSystemObject\")");
                sw.Write(sw.NewLine);
                sw.WriteLine("'Check if the root ZenoTrac folder exists, if not create it");
                sw.WriteLine("If objFile.FolderExists(rootPath) = False Then");
                sw.WriteLine("  objFile.CreateFolder (rootPath)");
                sw.WriteLine("End If");
                sw.Write(sw.NewLine);
                sw.WriteLine("'Check if the folder exists, if not create it");
                sw.WriteLine("If objFile.FolderExists(dirPath) = False Then");
                sw.WriteLine("  objFile.CreateFolder (dirPath)");
                sw.WriteLine("End If");
                sw.Write(sw.NewLine);

                sw.WriteLine("'Format today's file name");
                sw.WriteLine("FileName = \"C:\\ZenoTrac\\Common\\\" & Format$(Date, \"YYYY MM DD\") & \" 0000 Common.CSV\"");
                sw.WriteLine("'Format the current production record.");
                sw.WriteLine("With Fix32.Fix 'Modify node name to match the actual node");
                sw.WriteLine("RecordString = Date$ & \" \" & Time$ & \",\" _");

                //add membrane tank level and permeate turbidity, permeate Flow to common
                for (int i = startListTrain; i <= endListTrain; i++)
                {
                    //no permeate flow in common
                    if (commonTags[0] == false)
                    {
                        if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("PERMEATEFLOW"))
                        {

                            sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + endString);

                        }
                    }
                    if (commonTags[1] == false)
                    {
                        if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("PERMEATETURBIDITY"))
                        {
                            sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + endString);

                        }
                    }
                    if (commonTags[2] == false)
                    {
                        if (ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Contains("MEMBRANETANKLEVEL"))
                        {
                            sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + endString);

                        }
                    }

                }

                //write the common tags 
                for (int i = startListCommon; i <= endListCommon; i++)
                {
                    //Last tag will have no comma at the end
                    if (i == endListCommon)
                    {
                        sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + lastTagEndString);

                    }
                    else
                    {
                        sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + endString);

                    }

                }

                sw.WriteLine("End With");
                sw.WriteLine("'Find a free file ID to write to...");
                sw.WriteLine("FileNoOut = FreeFile()");
                sw.WriteLine("Open FileName For Append As #FileNoOut");
                sw.WriteLine("'Write data to output file");
                sw.WriteLine("Print #FileNoOut, RecordString");
                sw.WriteLine("Close #FileNoOut");
                sw.WriteLine("End Sub");

                sw.Write(sw.NewLine);
            }
            //Daily Tables
            if (totalDailyRows > 0)
            {
                sw.WriteLine("Private Sub Daily_OnTimeOut(ByVal lTimerId As Long)");
                sw.WriteLine("'Declarations");
                sw.WriteLine("Dim FileName As String");
                sw.WriteLine("Dim RecordString As String");
                sw.WriteLine("Dim FileNoOut As Integer");
                sw.Write(sw.NewLine);
                sw.WriteLine("'variables for the directory structure");
                sw.WriteLine("Dim rootPath, dirPath");
                sw.Write(sw.NewLine);
                sw.WriteLine("'If there is an error go to the next line");
                sw.WriteLine("On Error Resume Next");
                sw.Write(sw.NewLine);
                sw.WriteLine("'root path");
                sw.WriteLine("rootPath = \"C:\\ZenoTrac\\\"");
                sw.Write(sw.NewLine);
                sw.WriteLine("'set directory path");
                sw.WriteLine("dirPath = \"C:\\ZenoTrac\\Daily\\\"");
                sw.Write(sw.NewLine);
                sw.WriteLine("Set objFile = CreateObject(\"Scripting.FileSystemObject\")");
                sw.Write(sw.NewLine);
                sw.WriteLine("'Check if the root ZenoTrac folder exists, if not create it");
                sw.WriteLine("If objFile.FolderExists(rootPath) = False Then");
                sw.WriteLine("  objFile.CreateFolder (rootPath)");
                sw.WriteLine("End If");
                sw.Write(sw.NewLine);
                sw.WriteLine("'Check if the folder exists, if not create it");
                sw.WriteLine("If objFile.FolderExists(dirPath) = False Then");
                sw.WriteLine("  objFile.CreateFolder (dirPath)");
                sw.WriteLine("End If");
                sw.Write(sw.NewLine);

                sw.WriteLine("'Format today's file name");
                sw.WriteLine("FileName = \"C:\\ZenoTrac\\Daily\\\" & Format$(Date, \"YYYY MM DD\") & \" 0000 Daily.CSV\"");
                sw.WriteLine("'Format the current production record.");
                sw.WriteLine("With Fix32.Fix 'Modify node name to match the actual node");
                sw.WriteLine("     RecordString = Date$ & \" \" & Time$ & \",\" _");

                for (int i = startListDaily; i <= endListDaily; i++)
                {
                    //the last tag will not have the comma at the end
                    if (i == endListDaily)
                    {
                        sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + lastTagEndString);

                    }
                    else
                    {
                        sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + endString);
                    }
                }
                sw.WriteLine("End With");
                sw.WriteLine("'Find a free file ID to write to...");
                sw.WriteLine("FileNoOut = FreeFile()");
                sw.WriteLine("Open FileName For Append As #FileNoOut");
                sw.WriteLine("'Write data to output file");
                sw.WriteLine("Print #FileNoOut, RecordString");
                sw.WriteLine("Close #FileNoOut");
                sw.WriteLine("End Sub");

            }
            //Train tables
            if (totalTrainRows > 0)
            {
                //number of trains
                int totalNumberOfTrains = NumberofWorksheetsWithName("ZW");
                int[] totalTrainTags = new int[totalNumberOfTrains];
                totalTrainTags = GetNumberofTagsForWorksheet("ZW");

                for (int k = 0; k < totalTrainTags.Length; k++)
                {
                    //get the indexes for the last tag in each ZW worksheet
                    totalTrainTags[k] = totalTrainTags[k] + (startListTrain - 1);
                }
                                
                sw.Write(sw.NewLine);
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("ZW"))
                    {

                        sw.WriteLine("Private Sub " + worksheetNames[j] + "_OnTrue()");
                        sw.WriteLine("'Declarations");
                        sw.WriteLine("Dim FileName As String");
                        sw.WriteLine("Dim RecordString As String");
                        sw.WriteLine("Dim FileNoOut As Integer");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'variables for the directory structure");
                        sw.WriteLine("Dim rootPath, dirPath");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'If there is an error go to the next line");
                        sw.WriteLine("On Error Resume Next");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'root path");
                        sw.WriteLine("rootPath = \"C:\\ZenoTrac\\\"");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'set directory path");
                        sw.WriteLine("dirPath = \"C:\\ZenoTrac\\" + worksheetNames[j] + "\\\"");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("Set objFile = CreateObject(\"Scripting.FileSystemObject\")");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'Check if the root ZenoTrac folder exists, if not create it");
                        sw.WriteLine("If objFile.FolderExists(rootPath) = False Then");
                        sw.WriteLine("  objFile.CreateFolder (rootPath)");
                        sw.WriteLine("End If");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'Check if the folder exists, if not create it");
                        sw.WriteLine("If objFile.FolderExists(dirPath) = False Then");
                        sw.WriteLine("  objFile.CreateFolder (dirPath)");
                        sw.WriteLine("End If");
                        sw.Write(sw.NewLine);

                        sw.WriteLine("'Format today's file name");
                        sw.WriteLine("FileName = \"C:\\ZenoTrac\\" + worksheetNames[j] + "\\\" & Format$(Date, \"YYYY MM DD\") & \" 0000 " + worksheetNames[j] + ".CSV\"");
                        sw.WriteLine("'Format the current production record.");
                        sw.WriteLine("With Fix32.Fix 'Modify node name to match the actual node");
                        sw.WriteLine("     RecordString = Date$ & \" \" & Time$ & \",\" _");
                       
                        for (int i = startListTrain; i <= endListTrain; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = ifixScriptDataTable.Rows[i]["Tag Name"].ToString().IndexOf("_");
                            //add the first ZW train first
                            if (worksheetNames[j].ToString() == ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the train then check if permeate temperature exists
                                if (i == totalTrainTags[totalNumberOfTrains - 1] && totalNumberOfTrains >= 1)
                                {                                    
                                    //add no comma if there is no temperature in the common
                                    if (temperatureExists[0] == false && temperatureExists[1] == false)
                                    {
                                        sw.WriteLine(beginString + worksheetNames[j] + ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Substring(startPoint) + lastTagEndString);
                                    }
                                    //if temperature exists in the train already then no comma
                                    else if (temperatureExists[2] == true || temperatureExists[3] == true)
                                    {
                                        sw.WriteLine(beginString + worksheetNames[j] + ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Substring(startPoint) + lastTagEndString);

                                    }
                                    else
                                    {
                                        sw.WriteLine(beginString + worksheetNames[j] + ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Substring(startPoint) + endString);

                                    }
                                }
                                else
                                {
                                    sw.WriteLine(beginString + worksheetNames[j] + ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Substring(startPoint) + endString);
                                }
                            }                           
                            
                        }

                        //check if there is permeate or feed temperature in common and put in train
                        if (temperatureExists[0] == true && temperatureExists[2] == false)
                        {
                            sw.WriteLine(beginString + "CMN_PERMEATETEMPERATURE");
                        }
                        else if (temperatureExists[0] == false && temperatureExists[1] == true && temperatureExists[3] == false)
                        {
                            sw.WriteLine(beginString + "CMN_FEEDTEMPERATURE");
                        }


                        sw.WriteLine("End With");
                        sw.WriteLine("'Find a free file ID to write to...");
                        sw.WriteLine("FileNoOut = FreeFile()");
                        sw.WriteLine("Open FileName For Append As #FileNoOut");
                        sw.WriteLine("'Write data to output file");
                        sw.WriteLine("Print #FileNoOut, RecordString");
                        sw.WriteLine("Close #FileNoOut");
                        sw.WriteLine("End Sub");
                        sw.Write(sw.NewLine);

                        //decrement counter to traverse through the last tag indexes of each sheet
                        totalNumberOfTrains = totalNumberOfTrains - 1;
                    }
                }
            }

            //MIT        
            if (totalMitRows > 0)
            {
                int totalMitSheets = NumberofWorksheetsWithName("MIT");
                int[] totalMitTags = new int[totalMitSheets];
                totalMitTags = GetNumberofTagsForWorksheet("MIT");

                for (int k = 0; k < totalMitTags.Length; k++)
                {
                    //get the indexes for the last tag in each MIT worksheet
                    totalMitTags[k] = totalMitTags[k] + (startListMit - 1);
                }
                
               
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("MIT"))
                    {
                        sw.WriteLine("Private Sub " + worksheetNames[j] + "_OnTrue()");
                        sw.WriteLine("'Declarations");
                        sw.WriteLine("Dim FileName As String");
                        sw.WriteLine("Dim RecordString As String");
                        sw.WriteLine("Dim FileNoOut As Integer");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'variables for the directory structure");
                        sw.WriteLine("Dim rootPath, dirPath");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'If there is an error go to the next line");
                        sw.WriteLine("On Error Resume Next");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'root path");
                        sw.WriteLine("rootPath = \"C:\\ZenoTrac\\\"");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'set directory path");
                        sw.WriteLine("dirPath = \"C:\\ZenoTrac\\" + worksheetNames[j] + "\\\"");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("Set objFile = CreateObject(\"Scripting.FileSystemObject\")");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'Check if the root ZenoTrac folder exists, if not create it");
                        sw.WriteLine("If objFile.FolderExists(rootPath) = False Then");
                        sw.WriteLine("  objFile.CreateFolder (rootPath)");
                        sw.WriteLine("End If");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'Check if the folder exists, if not create it");
                        sw.WriteLine("If objFile.FolderExists(dirPath) = False Then");
                        sw.WriteLine("  objFile.CreateFolder (dirPath)");
                        sw.WriteLine("End If");
                        sw.Write(sw.NewLine);                      

                        sw.WriteLine("'Format today's file name");
                        sw.WriteLine("FileName = \"C:\\ZenoTrac\\" + worksheetNames[j] + "\\\" & Format$(Date, \"YYYY MM DD\") & \" 0000 " + worksheetNames[j] + ".CSV\"");
                        sw.WriteLine("'Format the current production record.");
                        sw.WriteLine("With Fix32.Fix 'Modify node name to match the actual node");
                        sw.WriteLine("     RecordString = Date$ & \" \" & Time$ & \",\" _");

                        for (int i = startListMit; i <= endListMit; i++)
                        {
                            startPoint = ifixScriptDataTable.Rows[i]["Tag Name"].ToString().IndexOf("_");                            
                            if (worksheetNames[j].ToString() == ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the MIT then put no comma
                                if (i == totalMitTags[totalMitSheets - 1] && totalMitSheets >= 1)
                                {
                                    sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + lastTagEndString);
                                }
                                else
                                {
                                    sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + endString);

                                }
                            }
                                

                        }
                        sw.WriteLine("End With");
                        sw.WriteLine("'Find a free file ID to write to...");
                        sw.WriteLine("FileNoOut = FreeFile()");
                        sw.WriteLine("Open FileName For Append As #FileNoOut");
                        sw.WriteLine("'Write data to output file");
                        sw.WriteLine("Print #FileNoOut, RecordString");
                        sw.WriteLine("Close #FileNoOut");
                        sw.WriteLine("End Sub");
                        sw.Write(sw.NewLine);
                        //decrement counter to traverse through the last tag indexes of each sheet                        
                        totalMitSheets = totalMitSheets - 1;
                    }
                }
            }


            //MC        
            if (totalMCRows > 0)
            {
                int totalMCSheets = NumberofWorksheetsWithName("MC");
                int[] totalMCTags = new int[totalMCSheets];
                totalMCTags = GetNumberofTagsForWorksheet("MC");

                for (int k = 0; k < totalMCTags.Length; k++)
                {
                    //get the indexes for the last tag in each MC worksheet
                    totalMCTags[k] = totalMCTags[k] + (startListMC - 1);
                }


                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("MC"))
                    {
                        sw.WriteLine("Private Sub " + worksheetNames[j] + "_OnTrue()");
                        sw.WriteLine("'Declarations");
                        sw.WriteLine("Dim FileName As String");
                        sw.WriteLine("Dim RecordString As String");
                        sw.WriteLine("Dim FileNoOut As Integer");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'variables for the directory structure");
                        sw.WriteLine("Dim rootPath, dirPath");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'If there is an error go to the next line");
                        sw.WriteLine("On Error Resume Next");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'root path");
                        sw.WriteLine("rootPath = \"C:\\ZenoTrac\\\"");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'set directory path");
                        sw.WriteLine("dirPath = \"C:\\ZenoTrac\\" + worksheetNames[j] + "\\\"");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("Set objFile = CreateObject(\"Scripting.FileSystemObject\")");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'Check if the root ZenoTrac folder exists, if not create it");
                        sw.WriteLine("If objFile.FolderExists(rootPath) = False Then");
                        sw.WriteLine("  objFile.CreateFolder (rootPath)");
                        sw.WriteLine("End If");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'Check if the folder exists, if not create it");
                        sw.WriteLine("If objFile.FolderExists(dirPath) = False Then");
                        sw.WriteLine("  objFile.CreateFolder (dirPath)");
                        sw.WriteLine("End If");
                        sw.Write(sw.NewLine);


                        sw.WriteLine("'Format today's file name");
                        sw.WriteLine("FileName = \"C:\\ZenoTrac\\" + worksheetNames[j] + "\\\" & Format$(Date, \"YYYY MM DD\") & \" 0000 " + worksheetNames[j] + ".CSV\"");
                        sw.WriteLine("'Format the current production record.");
                        sw.WriteLine("With Fix32.Fix 'Modify node name to match the actual node");
                        sw.WriteLine("     RecordString = Date$ & \" \" & Time$ & \",\" _");

                        for (int i = startListMC; i <= endListMC; i++)
                        {
                            startPoint = ifixScriptDataTable.Rows[i]["Tag Name"].ToString().IndexOf("_");
                            if (worksheetNames[j].ToString() == ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the MC then put no comma
                                if (i == totalMCTags[totalMCSheets - 1] && totalMCSheets >= 1)
                                {
                                    sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + lastTagEndString);
                                }
                                else
                                {
                                    sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + endString);

                                }
                            }


                        }
                        sw.WriteLine("End With");
                        sw.WriteLine("'Find a free file ID to write to...");
                        sw.WriteLine("FileNoOut = FreeFile()");
                        sw.WriteLine("Open FileName For Append As #FileNoOut");
                        sw.WriteLine("'Write data to output file");
                        sw.WriteLine("Print #FileNoOut, RecordString");
                        sw.WriteLine("Close #FileNoOut");
                        sw.WriteLine("End Sub");
                        sw.Write(sw.NewLine);
                        //decrement counter to traverse through the last tag indexes of each sheet                        
                        totalMCSheets = totalMCSheets - 1;
                    }
                }
            }


            //RC        
            if (totalRCRows > 0)
            {
                int totalRCSheets = NumberofWorksheetsWithName("RC");
                int[] totalRCTags = new int[totalRCSheets];
                totalRCTags = GetNumberofTagsForWorksheet("RC");

                for (int k = 0; k < totalRCTags.Length; k++)
                {
                    //get the indexes for the last tag in each RC worksheet
                    totalRCTags[k] = totalRCTags[k] + (startListRC - 1);
                }


                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("RC"))
                    {
                        sw.WriteLine("Private Sub " + worksheetNames[j] + "_OnTrue()");
                        sw.WriteLine("'Declarations");
                        sw.WriteLine("Dim FileName As String");
                        sw.WriteLine("Dim RecordString As String");
                        sw.WriteLine("Dim FileNoOut As Integer");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'variables for the directory structure");
                        sw.WriteLine("Dim rootPath, dirPath");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'If there is an error go to the next line");
                        sw.WriteLine("On Error Resume Next");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'root path");
                        sw.WriteLine("rootPath = \"C:\\ZenoTrac\\\"");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'set directory path");
                        sw.WriteLine("dirPath = \"C:\\ZenoTrac\\" + worksheetNames[j] + "\\\"");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("Set objFile = CreateObject(\"Scripting.FileSystemObject\")");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'Check if the root ZenoTrac folder exists, if not create it");
                        sw.WriteLine("If objFile.FolderExists(rootPath) = False Then");
                        sw.WriteLine("  objFile.CreateFolder (rootPath)");
                        sw.WriteLine("End If");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("'Check if the folder exists, if not create it");
                        sw.WriteLine("If objFile.FolderExists(dirPath) = False Then");
                        sw.WriteLine("  objFile.CreateFolder (dirPath)");
                        sw.WriteLine("End If");
                        sw.Write(sw.NewLine);

                        sw.WriteLine("'Format today's file name");
                        sw.WriteLine("FileName = \"C:\\ZenoTrac\\" + worksheetNames[j] + "\\\" & Format$(Date, \"YYYY MM DD\") & \" 0000 " + worksheetNames[j] + ".CSV\"");
                        sw.WriteLine("'Format the current production record.");
                        sw.WriteLine("With Fix32.Fix 'Modify node name to match the actual node");
                        sw.WriteLine("     RecordString = Date$ & \" \" & Time$ & \",\" _");

                        for (int i = startListRC; i <= endListRC; i++)
                        {
                            startPoint = ifixScriptDataTable.Rows[i]["Tag Name"].ToString().IndexOf("_");
                            if (worksheetNames[j].ToString() == ifixScriptDataTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the RC then put no comma
                                if (i == totalRCTags[totalRCSheets - 1] && totalRCSheets >= 1)
                                {
                                    sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + lastTagEndString);
                                }
                                else
                                {
                                    sw.WriteLine(beginString + ifixScriptDataTable.Rows[i]["Tag Name"].ToString() + endString);

                                }
                            }


                        }
                        sw.WriteLine("End With");
                        sw.WriteLine("'Find a free file ID to write to...");
                        sw.WriteLine("FileNoOut = FreeFile()");
                        sw.WriteLine("Open FileName For Append As #FileNoOut");
                        sw.WriteLine("'Write data to output file");
                        sw.WriteLine("Print #FileNoOut, RecordString");
                        sw.WriteLine("Close #FileNoOut");
                        sw.WriteLine("End Sub");
                        sw.Write(sw.NewLine);
                        //decrement counter to traverse through the last tag indexes of each sheet                        
                        totalRCSheets = totalRCSheets - 1;
                    }
                }
            }
            sw.Close();
        }


        public void WriteSQLDatabaseScriptToTextFile(string absolutePathAndFileName, string softwareType, string aoNumber, string siteName, string temperature, string flowRate, int cassettesPerTrain, int modulesPerCassette, float areaPerModule, bool areaSquareFeetChecked, string siteAssigned)
        {

            //declare variables
            DataTable databaseTagsTable = m_excelData.Copy();

            //remove address column that we don't need
            //databaseTagsTable.Columns.RemoveAt(1);
            StreamWriter sw = File.CreateText(absolutePathAndFileName);

            // 0 = Total Daily Feed Flow, 1 = Total Daily Reject Flow, 2 = Total Daily Waste Flow, 3 = Total Daily Plant Permeate Flow, 4 = Total Plant Daily Permeate Flow
            //take out totalDailyFlowExists as we don't need to do daily recovery
            //BitArray totalDailyFlowExists = new BitArray(5, false);
            BitArray temperatureExists = new BitArray(4, false);
            BitArray flowRateExists = new BitArray(3, false);
            BitArray tmpExists = new BitArray(3, false);
            BitArray commonTags = new BitArray(3, false);
            bool pressureDifferenceExists = false;
            bool aoNumberExists = false;
            string firstZWWorkSheetName = "";
            string tabSpace1 = "    ";
            string tabSpace2 = "        ";
            int startPoint = 0;
            int totalDataTableRows = 0;
            int totalCommonRows = 0;
            int totalDailyRows = 0;
            int totalTrainRows = 0;
            int totalMit1Rows = 0;
            int totalMitRows = 0;
            int totalTrain1Rows = 0;
            int startListCommon = 0;
            int endListCommon = 0;
            int startListDaily = 0;
            int endListDaily = 0;
            int startListTrain = 0;
            int endListTrain = 0;
            int startListMit1 = 0;
            int endListMit1 = 0;
            int startListMit = 0;
            int endListMit = 0;
            int startListTrain1 = 0;
            int endListTrain1 = 0;
            int startListMC = 0;
            int endListMC = 0;
            int totalMCRows = 0;
            int startListMC1 = 0;
            int endListMC1 = 0;
            int totalMC1Rows = 0;
            int startListRC = 0;
            int endListRC = 0;
            int totalRCRows = 0;
            int startListRC1 = 0;
            int endListRC1 = 0;
            int totalRC1Rows = 0;



            //number of trains
            int totalNumberOfTrains = NumberofWorksheetsWithName("ZW");
            //array of indices for train tabs
            int[] totalTrainTags = new int[totalNumberOfTrains];
            totalTrainTags = GetNumberofTagsForWorksheet("ZW");
            //mit sheets in Excel file
            int totalMitSheets = NumberofWorksheetsWithName("MIT");
            int[] totalMitTags = new int[totalMitSheets];
            totalMitTags = GetNumberofTagsForWorksheet("MIT");

            //MC sheets in Excel file
            int totalMCSheets = NumberofWorksheetsWithName("MC");
            int[] totalMCTags = new int[totalMCSheets];
            totalMCTags = GetNumberofTagsForWorksheet("MC");

            //RC sheets in Excel file
            int totalRCSheets = NumberofWorksheetsWithName("RC");
            int[] totalRCTags = new int[totalRCSheets];
            totalRCTags = GetNumberofTagsForWorksheet("RC");
           
            //check if this site has an ao number
            if (aoNumber != "")
            {
                aoNumberExists = true;
            }

            //Create Database

            sw.WriteLine("DECLARE");
            sw.WriteLine("@DatabaseName NVARCHAR (100),");
            sw.WriteLine("@DataFileName VARCHAR (200),");
            sw.WriteLine("@DataFileLocation VARCHAR (200),");
            sw.WriteLine("@TransactLogName VARCHAR (200),");
            sw.WriteLine("@TransactLogLocation VARCHAR (200),");
            sw.WriteLine("@DatabaseStatus BIT,");
            sw.WriteLine("@DatabaseSQL VARCHAR (2000),");
            sw.WriteLine("@ErrorSave INT");

            sw.Write(sw.NewLine);
            sw.WriteLine("--default 0 means database doesn't exist");
            sw.WriteLine("SET @DatabaseStatus = 0");
            sw.WriteLine("SET @ErrorSave = 0");
            if (aoNumberExists == true)
            {
                sw.WriteLine("SET @DatabaseName = '" + aoNumber + siteName + "'");
            }
            else
            {
                sw.WriteLine("SET @DatabaseName = '" + siteName + "'");
            }
            sw.WriteLine("SET @DataFileName = @DatabaseName + '_Data'");
            sw.WriteLine("SET @TransactLogName = @DatabaseName + '_Log'");
            sw.WriteLine("SET @TransactLogLocation = 'D:\\SQL Server Logs\\ZenoTrac Logs\\' + @TransactLogName + '.LDF'");
            sw.WriteLine("SET @DataFileLocation = 'E:\\Microsoft SQL Server\\MSSQL\\Data\\ZenoTrac Data\\' + @DataFileName + '.MDF'");
            sw.Write(sw.NewLine);
            sw.WriteLine("SET @DatabaseSQL = ('CREATE DATABASE [' + @DatabaseName + ']");
            sw.WriteLine("ON");
            sw.WriteLine("(NAME = ''' + @DataFileName + ''',");
            sw.WriteLine("FILENAME = ''' + @DataFileLocation + ''' ,");
            sw.WriteLine("SIZE = 2,");
            sw.WriteLine("FILEGROWTH = 10%)");
            sw.WriteLine("LOG ON");
            sw.WriteLine("(NAME = ''' + @TransactLogName + ''',");
            sw.WriteLine("FILENAME = ''' + @TransactLogLocation + ''' ,");
            sw.WriteLine("SIZE = 1,");
            sw.WriteLine("FILEGROWTH = 10%)");
            sw.WriteLine("COLLATE SQL_Latin1_General_CP1_CI_AS'");
            sw.WriteLine(");");
            sw.Write(sw.NewLine);
            sw.WriteLine("--check if database exists and set status bit to 1 if it does");
            sw.WriteLine("IF EXISTS (SELECT [name] FROM [master].[dbo].[sysdatabases]");
            sw.WriteLine("WHERE [name] = @DatabaseName)");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "SET @DatabaseStatus = 1");
            sw.WriteLine("END");
            sw.Write(sw.NewLine);
            sw.WriteLine("--Create Database if it doesn't exist");
            sw.WriteLine("IF @DatabaseStatus = 1");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "PRINT 'Cannot create the database, ' + @DatabaseName + ' because it already exists.'");
            sw.WriteLine("END");

            sw.WriteLine("ELSE");
            sw.WriteLine("--Create the database");
            sw.WriteLine(tabSpace1 + "IF @DatabaseStatus  = 0");
            sw.WriteLine(tabSpace1 + "BEGIN");
            sw.WriteLine(tabSpace2 + "EXECUTE (@DatabaseSQL)");
            sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
            sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
            sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
            sw.WriteLine(tabSpace2 + "ELSE");
            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'The database, ' + @DatabaseName + ' created successfully.'");
            sw.Write(sw.NewLine);

            if (aoNumberExists == true)
            {
                sw.WriteLine("--set the appropiate options");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'autoclose', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'trunc. log', N'true'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'torn page detection', N'true'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'read only', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'dbo use', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'single', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'autoshrink', N'true'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'ANSI null default', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'recursive triggers', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'ANSI nulls', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'concat null yields null', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'cursor close on commit', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'default to local cursor', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'quoted identifier', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'ANSI warnings', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'auto create statistics', N'true'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + aoNumber + siteName + "', N'auto update statistics', N'true'");
                sw.WriteLine(tabSpace1 + "END");
            }
            else
            {
                sw.WriteLine("--set the appropiate options");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'autoclose', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'trunc. log', N'true'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'torn page detection', N'true'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'read only', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'dbo use', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'single', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'autoshrink', N'true'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'ANSI null default', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'recursive triggers', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'ANSI nulls', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'concat null yields null', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'cursor close on commit', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'default to local cursor', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'quoted identifier', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'ANSI warnings', N'false'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'auto create statistics', N'true'");
                sw.WriteLine(tabSpace2 + "EXEC sp_dboption N'" + siteName + "', N'auto update statistics', N'true'");
                sw.WriteLine(tabSpace1 + "END");           
            
            }

            sw.WriteLine("GO");

            sw.Write(sw.NewLine);
            //add database to SQL maintenance plan
            sw.WriteLine("USE [msdb]");
            sw.WriteLine("GO");
            sw.Write(sw.NewLine);
            sw.WriteLine("DECLARE");
            sw.WriteLine("@DatabaseName VARCHAR (100)");
            sw.Write(sw.NewLine);
            if (aoNumberExists == true)
            {
                sw.WriteLine("SET @DatabaseName = '" + aoNumber + siteName + "'");
            }
            else
            {
                sw.WriteLine("SET @DatabaseName = '" + siteName + "'");
            }
            sw.Write(sw.NewLine);
            sw.WriteLine("--check if database exists before adding it to maintenance plan");
            sw.WriteLine("IF EXISTS (SELECT [name] FROM [master].[dbo].[sysdatabases]");
            sw.WriteLine("WHERE [name] = @DatabaseName)");
            sw.WriteLine("BEGIN");
            sw.WriteLine("--check if database exists in the maintenance plan already");
            sw.WriteLine(tabSpace1 + "IF EXISTS (SELECT [database_name] FROM [msdb].[dbo].[sysdbmaintplan_databases]");
            sw.WriteLine(tabSpace1 + "WHERE [database_name] = @DatabaseName)");
            sw.WriteLine(tabSpace1 + "BEGIN");
            sw.WriteLine(tabSpace2 + "PRINT 'Cannot insert the database ' + @DatabaseName + ' in the maintenance plan because it already exists.'");
            sw.WriteLine(tabSpace1 + "END");
            sw.WriteLine(tabSpace1 + "ELSE");
            sw.WriteLine(tabSpace1 + "BEGIN");
            sw.WriteLine(tabSpace2 + "Execute sp_add_maintenance_plan_db N'3F938632-174E-4C5B-AE64-1CD99ADB27CC',@DatabaseName");
            sw.WriteLine(tabSpace2 + "PRINT 'Successfully added the database, ' + @DatabaseName + ' to the maintenance plan.'");            
            sw.WriteLine(tabSpace1 + "END");
            sw.WriteLine("END");
            sw.WriteLine("ELSE");
            sw.WriteLine(tabSpace1 + "PRINT 'Cannot insert the database ' + @DatabaseName + ' in the maintenance plan because it does not exist.'");
            sw.WriteLine("GO");

            sw.Write(sw.NewLine);

            //Create Tables
            if (aoNumberExists == true)
            {
                sw.WriteLine("USE [" + aoNumber + siteName + "]");
            }
            else
            {
                sw.WriteLine("USE [" + siteName + "]");
            }
            sw.WriteLine("GO");
            sw.Write(sw.NewLine);
            sw.WriteLine("DECLARE");
            sw.WriteLine("@DatabaseName VARCHAR (100),");
            sw.WriteLine("@AONumber VARCHAR (100),");
            sw.WriteLine("@SiteName VARCHAR (100),");
            sw.WriteLine("@NoOfTrains INT,");
            sw.WriteLine("@DatabaseStatus BIT,");
            sw.WriteLine("@ErrorSave INT,");
            sw.WriteLine("@nSQL VARCHAR(8000)");
            sw.Write(sw.NewLine);
            if (aoNumberExists == true)
            {
                sw.WriteLine("SET @DatabaseName = '" + aoNumber + siteName + "'");
            }
            else
            {
                sw.WriteLine("SET @DatabaseName = '" + siteName + "'");
            }
            sw.WriteLine("SET @SiteName = '" + siteName + "'");
            sw.WriteLine("SET @AONumber = '" + aoNumber + "'");
            sw.WriteLine("SET @DatabaseStatus = 0");
            sw.WriteLine("SET @ErrorSave = 0");
            sw.Write(sw.NewLine);
            sw.WriteLine("--check if database exists and set status bit to 1 if it does");
            sw.WriteLine("IF EXISTS (SELECT [name] FROM [master].[dbo].[sysdatabases]");
            sw.WriteLine("WHERE [name] = @DatabaseName)");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "SET @DatabaseStatus = 1");
            sw.WriteLine("END");
            sw.Write(sw.NewLine);
            sw.WriteLine("--Create table and stored procedures if the database exists ");
            sw.WriteLine("IF @DatabaseStatus = 1");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[ZTQAQC]') IS NOT NULL");
            sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ZTQAQC because it already exists.'");
            sw.WriteLine(tabSpace1 + "ELSE");
            sw.WriteLine(tabSpace1 + "BEGIN");
            sw.WriteLine("-- Create ZTQA/QC Table");
            sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[ZTQAQC] (");
            sw.WriteLine(tabSpace2 + "[IssueNo] [int] IDENTITY (1, 1) NOT NULL ,");
            sw.WriteLine(tabSpace2 + "[EntryDate] [datetime] NULL ,");
            sw.WriteLine(tabSpace2 + "[EnteredBy] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");
            sw.WriteLine(tabSpace2 + "[ZenoTracFolder] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");
            sw.WriteLine(tabSpace2 + "[GraphName] [nvarchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");
            sw.WriteLine(tabSpace2 + "[Description] [nvarchar] (1500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");
            sw.WriteLine(tabSpace2 + "[ResolutionDate] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");
            sw.WriteLine(tabSpace2 + "[ResolvedBy] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");
            sw.WriteLine(tabSpace2 + "[ResolutionStatus] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");
            sw.WriteLine(tabSpace2 + "[Notes] [nvarchar] (1500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL");
            sw.WriteLine(tabSpace2 + ") ON [PRIMARY] '");
            sw.Write(sw.NewLine);

            sw.WriteLine(tabSpace2 + "--PRINT (@nSQL)");
            sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
            sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
            sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
            sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
            sw.WriteLine(tabSpace2 + "ELSE");
            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ZTQAQC.'");
            sw.WriteLine(tabSpace1 + "END");
            sw.Write(sw.NewLine);

            //table for the PLC tag addresses
            if (aoNumberExists == true)
            {
                sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PLCTagMapping]') IS NOT NULL");
            }
            else
            {
                sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_PLCTagMapping]') IS NOT NULL");
            }

            if (aoNumberExists == true)
            {
                sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_' + @AONumber + '_PLCTagMapping because it already exists.'");
            }
            else
            {
                sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_PLCTagMapping because it already exists.'");
            }
            sw.WriteLine(tabSpace1 + "ELSE");
            sw.WriteLine(tabSpace1 + "BEGIN");
            sw.WriteLine("--  Create PLC Tag Addresses Table");
            sw.Write(sw.NewLine);
            if (aoNumberExists == true)
            {
                sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PLCTagMapping] (");
            }
            else
            {
                sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[' + @SiteName + '_PLCTagMapping] (");
            }
            sw.WriteLine(tabSpace2 + "[ZenoTracTagName] [nvarchar] (1500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");
            sw.WriteLine(tabSpace2 + "[PLCAddress] [nvarchar] (1500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ");
            sw.WriteLine(tabSpace2 + ") ON [PRIMARY] '");
            sw.Write(sw.NewLine);
            sw.WriteLine(tabSpace2 + "--PRINT (@nSQL)");
            sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
            sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
            sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
            sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
            sw.WriteLine(tabSpace2 + "ELSE");
            if (aoNumberExists == true)
            {
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, " + siteName + "_" + aoNumber + "_PLCTagMapping.'");
            }
            else
            {
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, " + siteName + "_PLCTagMapping.'");
            }
            
            //total number of rows          
            totalDataTableRows = databaseTagsTable.Rows.Count;
       
            sw.Write(sw.NewLine);
           // sw.WriteLine(tabSpace2 + "SET @nSQL = '");
            if (aoNumberExists == true)
            {
                sw.WriteLine(tabSpace2 + "INSERT INTO [" + aoNumber + siteName + "].[dbo].[" + siteName + "_" + aoNumber + "_PLCTagMapping] ([ZenoTracTagName], [PLCAddress])");
            }
            else
            {
                sw.WriteLine(tabSpace2 + "INSERT INTO [" + siteName + "].[dbo].[" + siteName + "_PLCTagMapping] ([ZenoTracTagName], [PLCAddress])");
            }
            for (int i = 0; i <= (totalDataTableRows - 1); i++)
            {   //last row
                if (i == (totalDataTableRows - 1))
                {
                    sw.WriteLine(tabSpace2 + "SELECT '" + databaseTagsTable.Rows[i]["Tag Name"] + "', '" + databaseTagsTable.Rows[i]["Address"] + "'");
                    
                }
                else
                {
                    sw.WriteLine(tabSpace2 + "SELECT '" + databaseTagsTable.Rows[i]["Tag Name"] + "', '" + databaseTagsTable.Rows[i]["Address"] + "'");
                    sw.WriteLine(tabSpace2 + "UNION ALL");
                }
            }
            sw.Write(sw.NewLine);
            sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
            sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
            sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
            sw.WriteLine(tabSpace2 + "ELSE");
            if (aoNumberExists == true)
            {
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully inserted tags to the table, " + siteName + "_" + aoNumber + "_PLCTagMapping.'");
            }
            else
            {
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully inserted tags to the table, " + siteName + "_PLCTagMapping.'");
            }
            
            sw.WriteLine(tabSpace1 + "END");

            sw.Write(sw.NewLine);
            //get start and end indexes for the tags
            GetStartEndIndexesForTagsInDataTable(ref databaseTagsTable, ref startListCommon, ref endListCommon, ref totalCommonRows, ref startListDaily, ref endListDaily, ref totalDailyRows, ref startListTrain, ref endListTrain, ref totalTrainRows, ref startListTrain1, ref endListTrain1, ref totalTrain1Rows, ref startListMit, ref endListMit, ref totalMitRows, ref startListMit1, ref endListMit1, ref totalMit1Rows, ref startListMC, ref endListMC, ref totalMCRows, ref startListMC1, ref endListMC1, ref totalMC1Rows, ref startListRC, ref endListRC, ref totalRCRows, ref startListRC1, ref endListRC1, ref totalRC1Rows);
 
            //get first worksheet name in the excel file
            firstZWWorkSheetName = SearchFirstNameOfWorkSheet("ZW");

            /*
            //get first worksheet name in the excel file
            for (int j = 0; j < worksheetNames.Length; j++)
            {
                if (worksheetNames[j].ToString().Contains("ZW"))
                {
                    firstZWWorkSheetName = worksheetNames[j].ToString();
                    break;
                }

            }
             * 
             */

            //check if there is permeate or feed temperature in common and put in train
            for (int i = startListCommon; i <= endListCommon; i++)
            {
                if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateTemperature"))
                {
                    temperatureExists[0] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("FeedTemperature"))
                {
                    temperatureExists[1] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateFlow" + firstZWWorkSheetName) || databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains(firstZWWorkSheetName + "PermeateFlow"))
                {
                    commonTags[0] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateTurbidity" + firstZWWorkSheetName) || databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains(firstZWWorkSheetName + "PermeateTurbidity"))
                {
                    commonTags[1] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("MembraneTankLevel" + firstZWWorkSheetName) || databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains(firstZWWorkSheetName + "MembraneTankLevel"))
                {
                    commonTags[2] = true;
                }
            }

            //Check if flowrates and tmps exist
            for (int i = startListTrain1; i <= endListTrain1; i++)
            {
                if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("BeforeBPFlowRate"))
                {
                    flowRateExists[0] = true;
                    
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("DuringBPFlowRate"))
                {
                    flowRateExists[1] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("AfterBPFlowRate"))
                {
                    flowRateExists[2] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("BeforeBPTMP"))
                {
                    tmpExists[0] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("DuringBPTMP"))
                {
                    tmpExists[1] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("AfterBPTMP"))
                {
                    tmpExists[2] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateTemperature"))
                {
                    temperatureExists[2] = true;
                }
                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("FeedTemperature"))
                {
                    temperatureExists[3] = true;
                }
            }

            
            //Can take this out - don't need to do calculation for Daily
            /*
            //check if mit exists to see if it is drinking water plant.  Then set appropiate tags to true.
            if (totalMit1Rows > 0)
            {
                for (int i = startListDaily; i <= endListDaily; i++)
                {
                    if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("FeedFlow"))
                    {
                        totalDailyFlowExists[0] = true;
                    }
                    else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("RejectFlow"))
                    {
                        totalDailyFlowExists[1] = true;
                    }
                    else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("WasteFlow"))
                    {
                        totalDailyFlowExists[2] = true;
                    }
                    else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("DailyPlantPermeateFlow"))
                    {
                        totalDailyFlowExists[3] = true;
                    }
                    else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PlantDailyPermeateFlow"))
                    {
                        totalDailyFlowExists[4] = true;
                    }
                }
            }
            */

            //Write script if there is a common tab in the Excel file
            if (totalCommonRows > 0)
            {
                sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[TempDataCommon]') IS NOT NULL");
                sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, TempDataCommon because it already exists.'");
                sw.WriteLine(tabSpace1 + "ELSE");
                sw.WriteLine(tabSpace1 + "BEGIN");
                sw.WriteLine("--Temp Common Table");
                sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[TempDataCommon] (");

                if (softwareType == "IFIX")
                    sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NULL ,");
                else
                    //OPC Trend
                    sw.WriteLine(tabSpace2 + "[DateandTime] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");

                //add membrane tank level and permeate turbidity, permeate Flow to common
                for (int i = startListTrain; i <= endListTrain; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    /*
                    if (commonTags[0] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateFlow"))
                        {

                            sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + "] [float] NULL ,");

                        }
                    }
                    */ 
                    if (commonTags[1] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateTurbidity"))
                        {
                            sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + "] [float] NULL ,");

                        }
                    }
                    if (commonTags[2] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("MembraneTankLevel"))
                        {
                            sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + "] [float] NULL ,");

                        }
                    }


                }

                //write the common tags 
                for (int i = startListCommon; i <= endListCommon; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    //Last tag will have no comma at the end
                    if (i == endListCommon)
                    {
                        sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ");

                    }
                    else
                    {
                        sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");

                    }

                }
                sw.WriteLine(tabSpace2 + ") ON [PRIMARY] '");
                sw.Write(sw.NewLine);
                sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                sw.WriteLine(tabSpace2 + "ELSE");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, TempDataCommon.'");
                sw.WriteLine(tabSpace1 + "END");
                sw.Write(sw.NewLine);
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PDCommon]') IS NOT NULL");
                }
                else
                {
                    sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_PDCommon]') IS NOT NULL");
                }
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_' + @AONumber + '_PDCommon because it already exists.'");
                }
                else
                {
                    sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_PDCommon because it already exists.'");
                }
                sw.WriteLine(tabSpace1 + "ELSE");
                sw.WriteLine(tabSpace1 + "BEGIN");
                sw.WriteLine("--Common Production Table");
                sw.Write(sw.NewLine);
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PDCommon] (");
                }
                else
                {
                    sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[' + @SiteName + '_PDCommon] (");
                }
                sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NOT NULL ,");

                //add membrane tank level and permeate turbidity, to common
                for (int i = startListTrain; i <= endListTrain; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    /*
                    if (commonTags[0] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateFlow"))
                        {

                            sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + "] [float] NULL ,");

                        }
                    }*/
                    if (commonTags[1] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateTurbidity"))
                        {
                            sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + "] [float] NULL ,");

                        }
                    }
                    if (commonTags[2] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("MembraneTankLevel"))
                        {
                            sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + "] [float] NULL ,");

                        }
                    }

                }

                //write the common tags 
                for (int i = startListCommon; i <= endListCommon; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    //Last tag will have no comma at the end
                    sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");
                }
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_' + @AONumber + '_PDCommon] PRIMARY KEY  CLUSTERED");
                }
                else
                {
                    sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_PDCommon] PRIMARY KEY  CLUSTERED");
                }
                sw.WriteLine(tabSpace2 + "  (");
                sw.WriteLine(tabSpace2 + tabSpace2 + "[DateandTime]");
                sw.WriteLine(tabSpace2 + "  )  ON [PRIMARY]");
                sw.WriteLine(tabSpace2 + ") ON [PRIMARY]'");
                sw.Write(sw.NewLine);
                sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                sw.WriteLine(tabSpace2 + "ELSE");
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_' + @AONumber + '_PDCommon.'");
                }
                else
                {
                    sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_PDCommon.'");
                }
                sw.WriteLine(tabSpace1 + "END");
                sw.Write(sw.NewLine);
            }

            //Daily Tables
            if (totalDailyRows > 0)
            {
                sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[TempDataDaily]') IS NOT NULL");
                sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, TempDataDaily because it already exists.'");
                sw.WriteLine(tabSpace1 + "ELSE");
                sw.WriteLine(tabSpace1 + "BEGIN");
                sw.WriteLine("--Daily Temp Table");
                sw.Write(sw.NewLine);
                sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[TempDataDaily] (");
                if (softwareType == "IFIX")
                    sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NULL ,");
                else
                    //OPC Trend
                    sw.WriteLine(tabSpace2 + "[DateandTime] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");

                for (int i = startListDaily; i <= endListDaily; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    //the last tag will not have the comma at the end
                    if (i == endListDaily)
                    {
                        sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ");

                    }
                    else
                    {
                        sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");

                    }
                }
                sw.WriteLine(tabSpace2 + ") ON [PRIMARY] '");
                sw.Write(sw.NewLine);
                sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                sw.WriteLine(tabSpace2 + "ELSE");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, TempDataDaily.'");
                sw.WriteLine(tabSpace1 + "END");
                sw.Write(sw.NewLine);
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PDDaily]') IS NOT NULL");
                }
                else
                {
                    sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_PDDaily]') IS NOT NULL");
                }
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_' + @AONumber + '_PDDaily because it already exists.'");
                }
                else
                {
                    sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_PDDaily because it already exists.'");
                }
                sw.WriteLine(tabSpace1 + "ELSE");
                sw.WriteLine(tabSpace1 + "BEGIN");
                sw.WriteLine("--Daily Production Table");
                sw.Write(sw.NewLine);
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PDDaily] (");
                }
                else
                {
                    sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[' + @SiteName + '_PDDaily] (");
                }
                sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NOT NULL ,");

                for (int i = startListDaily; i <= endListDaily; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    if (i == endListDaily)
                    {                      
                            sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");
                        
                    }
                    else
                    {
                        sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");

                    }
                }
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_' + @AONumber + '_PDDaily] PRIMARY KEY  CLUSTERED");
                }
                else
                {
                    sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_PDDaily] PRIMARY KEY  CLUSTERED");
                }
                sw.WriteLine(tabSpace2 + "  (");
                sw.WriteLine(tabSpace2 + tabSpace2 + "[DateandTime]");
                sw.WriteLine(tabSpace2 + "  )  ON [PRIMARY]");
                sw.WriteLine(tabSpace2 + ") ON [PRIMARY]'");
                sw.Write(sw.NewLine);
                sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                sw.WriteLine(tabSpace2 + "ELSE");
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_' + @AONumber + '_PDDaily.'");
                }
                else
                {
                    sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_PDDaily.'");
                }
                sw.WriteLine(tabSpace1 + "END");
                sw.Write(sw.NewLine);

            }

            //Train tables

            if (totalTrainRows > 0)
            {
                //number of trains
                totalNumberOfTrains = NumberofWorksheetsWithName("ZW");
            

                for (int k = 0; k < totalTrainTags.Length; k++)
                {
                    //get the indexes for the last tag in each MIT worksheet
                    totalTrainTags[k] = totalTrainTags[k] + (startListTrain - 1);
                }
                sw.Write(sw.NewLine);
                sw.WriteLine("--Temp Train Tables");

                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("ZW"))
                    {
                        sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[TempData" + worksheetNames[j] + "]') IS NOT NULL");
                        sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, TempData" + worksheetNames[j] + " because it already exists.'");
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");

                        sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].[TempData" + worksheetNames[j] + "] (");

                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NULL ,");
                        }
                        else
                        {
                            //OPC Trend
                            sw.WriteLine(tabSpace2 + "[DateandTime] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");
                        }
                                               
                        for (int i = startListTrain; i <= endListTrain; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            //find the train tags
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //last tag
                                if (i == totalTrainTags[totalNumberOfTrains - 1] && totalNumberOfTrains >= 1)
                                {                                    
                                    //no temperature then no comma
                                    if (temperatureExists[0] == false && temperatureExists[1] == false)
                                    {
                                        sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ");
                                    }
                                    //if temperature exists in the train already then no comma
                                    else if (temperatureExists[2] == true || temperatureExists[3] == true)
                                    {
                                        sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ");
                                    }
                                    else
                                    {
                                        sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");
                                    }
                                }
                                else
                                {
                                    sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");
                                }
                            }

                        }

                        if (temperatureExists[0] == true && temperatureExists[2] == false)
                        {
                            sw.WriteLine(tabSpace2 + "[PermeateTemperature] [float] NULL ");
                        }
                        //no permeate temp and feedtemp exists in the common but not in the trains then add it.
                        else if (temperatureExists[0] == false && temperatureExists[1] == true && temperatureExists[3] == false)
                        {
                            sw.WriteLine(tabSpace2 + "[FeedTemperature] [float] NULL ");
                        }

                        sw.WriteLine(tabSpace2 + ") ON [PRIMARY] '");
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, TempData" + worksheetNames[j] + ".'");
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);

                        totalNumberOfTrains = totalNumberOfTrains - 1;
                    }
                    
                }
                //Train Production tables

                sw.WriteLine("--Production Train Tables");
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("ZW"))
                    {
                        
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "]') IS NOT NULL");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_PD" + worksheetNames[j] + "]') IS NOT NULL");
                        }
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " because it already exists.'");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_PD" + worksheetNames[j] + " because it already exists.'");
                        }
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "] (");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[' + @SiteName + '_PD" + worksheetNames[j] + "] (");
                        }
                        sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NOT NULL ,");
                        for (int i = startListTrain; i <= endListTrain; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                             //find the train tags
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {                               
                                sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");
                            }
                        }
                        //check if BeforeBPFlowrate exists
                        if (flowRateExists[0] == true)
                        {
                            sw.WriteLine(tabSpace2 + "[BeforeBPFlux] [float] NULL ,");
                        }
                        //check if DuringBPFlowrate exists
                        if (flowRateExists[1] == true)
                        {
                            sw.WriteLine(tabSpace2 + "[DuringBPFlux] [float] NULL ,");
                        }
                        //check if AfterBPFlowrate exists
                        if (flowRateExists[2] == true)
                        {
                            sw.WriteLine(tabSpace2 + "[AfterBPFlux] [float] NULL ,");

                        }
                        if (flowRateExists[0] == true && tmpExists[0] == true)
                        {
                            sw.WriteLine(tabSpace2 + "[BeforeBPPermeability] [float] NULL ,");
                        }
                        if (flowRateExists[1] == true && tmpExists[1] == true)
                        {
                            sw.WriteLine(tabSpace2 + "[DuringBPPermeability] [float] NULL ,");
                        }
                        if (flowRateExists[2] == true && tmpExists[2] == true)
                        {
                            sw.WriteLine(tabSpace2 + "[AfterBPPermeability] [float] NULL ,");
                        }
                        if (temperatureExists[0] == true || temperatureExists[1] == true)
                        {
                            if (flowRateExists[0] == true)
                            {
                                sw.WriteLine(tabSpace2 + "[BeforeBPTempCorrFlux] [float] NULL ,");
                            }
                            if (flowRateExists[1] == true)
                            {
                                sw.WriteLine(tabSpace2 + "[DuringBPTempCorrFlux] [float] NULL ,");
                            }

                            if (flowRateExists[2] == true)
                            {
                                sw.WriteLine(tabSpace2 + "[AfterBPTempCorrFlux] [float] NULL ,");
                            }

                            if (flowRateExists[0] == true && tmpExists[0] == true)
                            {
                                sw.WriteLine(tabSpace2 + "[BeforeBPTempCorrPermeability] [float] NULL ,");
                            }
                            if (flowRateExists[1] == true && tmpExists[1] == true)
                            {
                                sw.WriteLine(tabSpace2 + "[DuringBPTempCorrPermeability] [float] NULL ,");
                            }
                            if (flowRateExists[2] == true && tmpExists[2] == true)
                            {
                                sw.WriteLine(tabSpace2 + "[AfterBPTempCorrPermeability] [float] NULL ,");
                            }
                        }

                        if (temperatureExists[0] == true && temperatureExists[2] == false)
                        {
                            sw.WriteLine(tabSpace2 + "[PermeateTemperature] [float] NULL ,");
                        }
                        //no permeate temp and feedtemp exists in the common but not in the trains then add it.
                        else if (temperatureExists[0] == false && temperatureExists[1] == true && temperatureExists[3] == false)
                        {
                            sw.WriteLine(tabSpace2 + "[FeedTemperature] [float] NULL ,");
                        }
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "] PRIMARY KEY  CLUSTERED");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_PD" + worksheetNames[j] + "] PRIMARY KEY  CLUSTERED");
                        }
                        sw.WriteLine(tabSpace2 + "  (");
                        sw.WriteLine(tabSpace2 + tabSpace2 + "[DateandTime]");
                        sw.WriteLine(tabSpace2 + "  )  ON [PRIMARY]");
                        sw.WriteLine(tabSpace2 + ") ON [PRIMARY]'");
                        sw.Write(sw.NewLine);
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ".'");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_PD" + worksheetNames[j] + ".'");
                        }
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                    }
                }          
            }

            //MIT 

            if (totalMitRows > 0)
            {
                totalMitSheets = NumberofWorksheetsWithName("MIT");
             
                for (int k = 0; k < totalMitTags.Length; k++)
                {
                    //get the indexes for the last tag in each MIT worksheet
                    totalMitTags[k] = totalMitTags[k] + (startListMit - 1);
                  
                }

                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("MIT"))
                    {
                        sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[TempData" + worksheetNames[j] + "]') IS NOT NULL");
                        sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, TempData" + worksheetNames[j] + " because it already exists.'");
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");

                        sw.WriteLine("--Temp MIT Tables");
                        sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[TempData" + worksheetNames[j] + "] (");
                        if (softwareType == "IFIX")
                            sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NULL ,");
                        else
                            //OPC Trend
                            sw.WriteLine(tabSpace2 + "[DateandTime] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");

                        for (int i = startListMit; i <= endListMit; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {                               
                                //when you get to the last tag for the MIT then put no comma
                                if (i == totalMitTags[totalMitSheets - 1] && totalMitSheets >= 1)
                                {
                                    
                                    sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ");
                                }
                                else
                                {
                                    sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");

                                }
                            }
                        }

                        sw.WriteLine(tabSpace2 + ") ON [PRIMARY] '");

                        sw.Write(sw.NewLine);
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, TempData" + worksheetNames[j] + ".'");
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                        //decrement counter to traverse through the last tag indexes of each sheet
                        totalMitSheets = totalMitSheets - 1;
                    }
                }

              
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("MIT"))
                    {
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "]') IS NOT NULL");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_PD" + worksheetNames[j] + "]') IS NOT NULL");
                        }
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " because it already exists.'");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_PD" + worksheetNames[j] + " because it already exists.'");
                        }
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");
                        sw.WriteLine("--Production MIT Tables");
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "] (");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[' + @SiteName + '_PD" + worksheetNames[j] + "] (");
                        }
                        sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NOT NULL ,");

                        //check if Pressure difference exists already
                        for (int i = startListMit; i <= endListMit; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PressureDifference"))
                                {
                                    pressureDifferenceExists = true;
                                }
                            }
                        }

                        //only put in Pressure Difference if there is start pressure and end pressure
                        if (pressureDifferenceExists == false)
                        {
                            for (int i = startListMit; i <= endListMit; i++)
                            {
                                //remove sheetname from the tagname when printing the tag
                                startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                                if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                                {
                                    if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("StartPressure"))
                                    {
                                        for (int k = startListMit; k <= endListMit; k++)
                                        {
                                            //remove sheetname from the tagname when printing the tag
                                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                                            //only want to go through each worksheet (ie MIT1 - should only go through MIT1 tags)
                                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[k]["Tag Name"].ToString().Substring(0, startPoint))
                                            {
                                                if (databaseTagsTable.Rows[k]["Tag Name"].ToString().Contains("EndPressure") || databaseTagsTable.Rows[k]["Tag Name"].ToString().Contains("FinishPressure"))
                                                {
                                                    sw.WriteLine(tabSpace2 + "[PressureDifference] [float] NULL ,");
                                                    // break;
                                                }
                                            }

                                        }
                                        // break;
                                    }

                                }
                            }
                        }
                        for (int i = startListMit; i <= endListMit; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {                               
                                sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");
                            }

                        }
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "] PRIMARY KEY  CLUSTERED");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_PD" + worksheetNames[j] + "] PRIMARY KEY  CLUSTERED");
                        }
                        sw.WriteLine(tabSpace2 + "  (");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "[DateandTime]");
                        sw.WriteLine(tabSpace2 + "  )  ON [PRIMARY]");
                        sw.WriteLine(tabSpace2 + ") ON [PRIMARY]'");
                        sw.Write(sw.NewLine);
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ".'");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_PD" + worksheetNames[j] + ".'");
                        }
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                    }
                }
            }

            //MC - Maintenance Clean 

            if (totalMCRows > 0)
            {
                totalMCSheets = NumberofWorksheetsWithName("MC");

                for (int k = 0; k < totalMCTags.Length; k++)
                {
                    //get the indexes for the last tag in each MC worksheet
                    totalMCTags[k] = totalMCTags[k] + (startListMC - 1);

                }

                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("MC"))
                    {
                        sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[TempData" + worksheetNames[j] + "]') IS NOT NULL");
                        sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, TempData" + worksheetNames[j] + " because it already exists.'");
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");

                        sw.WriteLine("--Temp MC Tables");
                        sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[TempData" + worksheetNames[j] + "] (");
                        if (softwareType == "IFIX")
                            sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NULL ,");
                        else
                            //OPC Trend
                            sw.WriteLine(tabSpace2 + "[DateandTime] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");

                        for (int i = startListMC; i <= endListMC; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the MC then put no comma
                                if (i == totalMCTags[totalMCSheets - 1] && totalMCSheets >= 1)
                                {

                                    sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ");
                                }
                                else
                                {
                                    sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");

                                }
                            }
                        }

                        sw.WriteLine(tabSpace2 + ") ON [PRIMARY] '");

                        sw.Write(sw.NewLine);
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, TempData" + worksheetNames[j] + ".'");
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                        //decrement counter to traverse through the last tag indexes of each sheet
                        totalMCSheets = totalMCSheets - 1;
                    }
                }

                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("MC"))
                    {
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "]') IS NOT NULL");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_PD" + worksheetNames[j] + "]') IS NOT NULL");
                        }
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " because it already exists.'");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_PD" + worksheetNames[j] + " because it already exists.'");
                        }
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");
                        sw.WriteLine("--Production MC Tables");
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "] (");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[' + @SiteName + '_PD" + worksheetNames[j] + "] (");
                        }
                        sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NOT NULL ,");

                        for (int i = startListMC; i <= endListMC; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");
                            }

                        }
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "] PRIMARY KEY  CLUSTERED");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_PD" + worksheetNames[j] + "] PRIMARY KEY  CLUSTERED");
                        }
                        sw.WriteLine(tabSpace2 + "  (");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "[DateandTime]");
                        sw.WriteLine(tabSpace2 + "  )  ON [PRIMARY]");
                        sw.WriteLine(tabSpace2 + ") ON [PRIMARY]'");
                        sw.Write(sw.NewLine);
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ".'");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_PD" + worksheetNames[j] + ".'");
                        }
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                    }
                }
            }

            //RC - Recovery Clean 

            if (totalRCRows > 0)
            {
                totalRCSheets = NumberofWorksheetsWithName("RC");

                for (int k = 0; k < totalRCTags.Length; k++)
                {
                    //get the indexes for the last tag in each RC worksheet
                    totalRCTags[k] = totalRCTags[k] + (startListRC - 1);

                }

                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("RC"))
                    {
                        sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[TempData" + worksheetNames[j] + "]') IS NOT NULL");
                        sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, TempData" + worksheetNames[j] + " because it already exists.'");
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");

                        sw.WriteLine("--Temp RC Tables");
                        sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[TempData" + worksheetNames[j] + "] (");
                        if (softwareType == "IFIX")
                            sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NULL ,");
                        else
                            //OPC Trend
                            sw.WriteLine(tabSpace2 + "[DateandTime] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,");

                        for (int i = startListRC; i <= endListRC; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the RC then put no comma
                                if (i == totalRCTags[totalRCSheets - 1] && totalRCSheets >= 1)
                                {

                                    sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ");
                                }
                                else
                                {
                                    sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");

                                }
                            }
                        }

                        sw.WriteLine(tabSpace2 + ") ON [PRIMARY] '");

                        sw.Write(sw.NewLine);
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, TempData" + worksheetNames[j] + ".'");
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                        //decrement counter to traverse through the last tag indexes of each sheet
                        totalRCSheets = totalRCSheets - 1;
                    }
                }

                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    if (worksheetNames[j].ToString().Contains("RC"))
                    {
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "]') IS NOT NULL");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace1 + "IF OBJECT_ID('[' + @DatabaseName + '].[dbo].[' + @SiteName + '_PD" + worksheetNames[j] + "]') IS NOT NULL");
                        }
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " because it already exists.'");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "PRINT 'Cannot create the table, ' + @SiteName + '_PD" + worksheetNames[j] + " because it already exists.'");
                        }
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");
                        sw.WriteLine("--Production RC Tables");
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "] (");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "SET @nSQL = 'Create Table [' + @DatabaseName + '].[dbo].' + '[' + @SiteName + '_PD" + worksheetNames[j] + "] (");
                        }
                        sw.WriteLine(tabSpace2 + "[DateandTime] [datetime] NOT NULL ,");

                        for (int i = startListRC; i <= endListRC; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                sw.WriteLine(tabSpace2 + "[" + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + "] [float] NULL ,");
                            }

                        }
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + "] PRIMARY KEY  CLUSTERED");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + "CONSTRAINT [PK_' + @SiteName + '_PD" + worksheetNames[j] + "] PRIMARY KEY  CLUSTERED");
                        }
                        sw.WriteLine(tabSpace2 + "  (");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "[DateandTime]");
                        sw.WriteLine(tabSpace2 + "  )  ON [PRIMARY]");
                        sw.WriteLine(tabSpace2 + ") ON [PRIMARY]'");
                        sw.Write(sw.NewLine);
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ".'");
                        }
                        else
                        {
                            sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the table, ' + @SiteName + '_PD" + worksheetNames[j] + ".'");
                        }
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                    }
                }
            }


            //Stored Procedures
           // totalNumberOfTrains = NumberofWorksheetsWithName("ZW");
            if (totalTrainRows > 0)
            {
                totalNumberOfTrains = NumberofWorksheetsWithName("ZW");
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with ZW in it
                    if (worksheetNames[j].ToString().Contains("ZW"))
                    {
                        sw.WriteLine(tabSpace1 + "IF EXISTS (SELECT * FROM dbo.sysobjects");
                        sw.WriteLine(tabSpace2 + "WHERE id = object_id(N'[dbo].[up_fromTmptoProductionData" + worksheetNames[j] + "]')");
                        sw.WriteLine(tabSpace2 + "AND OBJECTPROPERTY(id, N'IsProcedure') = 1)");
                        sw.WriteLine(tabSpace1 + "PRINT 'Cannot create the stored procedure, up_fromTmptoProductionData" + worksheetNames[j] + " because it already exists.'");
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");
                        sw.Write(sw.NewLine);
                        sw.WriteLine(tabSpace2 + "SET @nSQL = '");
                        if (siteAssigned == "Sandeep")
                        {
                            sw.WriteLine("--SJ");
                        }
                        else if (siteAssigned == "Saima")
                        {
                            sw.WriteLine("--SB");
                        }
                        else if (siteAssigned == "Edison")
                        {
                            sw.WriteLine("--EC");
                        }
                        else
                        {
                            sw.WriteLine("--DM");
                        }

                        sw.WriteLine("--' + CAST(DATENAME(MONTH, GETDATE()) AS VARCHAR) + ' ' + CAST(DATEPART(DAY, GETDATE())AS VARCHAR) + ', ' +  CAST(DATEPART(YEAR, GETDATE())AS VARCHAR) + '");
                        if (areaSquareFeetChecked == true)
                        {
                            sw.WriteLine("--Units in imperial");
                        }
                        else
                        {
                            sw.WriteLine("--Units in metric");
                        }

                        sw.WriteLine("CREATE PROCEDURE up_fromTmptoProductionData" + worksheetNames[j]);
                        sw.WriteLine("AS");
                        if (flowRateExists[0] == true || flowRateExists[1] == true || flowRateExists[2] == true)
                        {
                            sw.WriteLine("DECLARE");
                            sw.Write(sw.NewLine);
                            sw.WriteLine("--Constants");
                            sw.WriteLine("@MembraneArea float  -- Membrane area for flux calculations");
                            sw.Write(sw.NewLine);
                            sw.WriteLine("SET NOCOUNT ON -- Capture status meassages");
                            sw.Write(sw.NewLine);
                            if (areaSquareFeetChecked == true)
                            {
                                sw.WriteLine("SET @MembraneArea = " + cassettesPerTrain + "* " + modulesPerCassette + "* " + areaPerModule + " -- cassettes/train *  modules/cassette *  sqft/mod = sqft/train");
                            }
                            else
                            {
                                sw.WriteLine("SET @MembraneArea = " + cassettesPerTrain + "* " + modulesPerCassette + "* " + areaPerModule + " -- cassettes/train *  modules/cassette *  sqmetres/mod = sqmetres/train");
                            }
                        }

                        sw.Write(sw.NewLine);
                        sw.WriteLine("--delete rows that have no date and time value");
                        sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                        sw.WriteLine("WHERE DateandTime IS NULL");
                        sw.Write(sw.NewLine);
                        //Don't want to delete these rows as it was causing data gaps
                        /*
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that have duplicate date and time stamps");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM TempData" + worksheetNames[j] + " GROUP BY DateandTime having Count(*)>1)");
                        }
                        else
                        {
                            sw.WriteLine("--delete rows that have duplicate date and time stamps");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempData" + worksheetNames[j] + " GROUP BY CAST (LEFT(DateandTime,17) AS DATETIME) having Count(*)>1)");
                        }
                        sw.Write(sw.NewLine);
                         */
                        //Adding new statement
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ")");
                            }
                            else
                            {
                                sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_PD" + worksheetNames[j] + ")");
                            }
          
                        }
                        else
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ")");
                            }
                            else
                            {
                                sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_PD" + worksheetNames[j] + ")");
                            }
                           
                        }

                        sw.Write(sw.NewLine);
                        /*
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " WHERE DateandTime IN (SELECT DateandTime FROM TempData" + worksheetNames[j] + "))");

                        }
                        else
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " WHERE DateandTime IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempData" + worksheetNames[j] + "))");
                        }
                        sw.Write(sw.NewLine);
                         */
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine("INSERT INTO ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " (");
                        }
                        else
                        {
                            sw.WriteLine("INSERT INTO ' + @SiteName + '_PD" + worksheetNames[j] + " (");
                        }
                        sw.WriteLine("DateandTime ,");
                        for (int i = startListTrain; i <= endListTrain; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                             //add the first ZW train first
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //if it's the last tag to write and there is no BeforeBPflowrates then add no comma
                                if (i == totalTrainTags[totalNumberOfTrains - 1] && totalNumberOfTrains >= 1)
                                {                                    
                                    //no flowrates and no temperature then no comma
                                    if (flowRateExists[0] == false && flowRateExists[1] == false && flowRateExists[2] == false && temperatureExists[0] == false && temperatureExists[1] == false)
                                    {
                                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));

                                    }
                                    //if temperature in the trains exist then no comma
                                    else if ((flowRateExists[0] == false && flowRateExists[1] == false && flowRateExists[2] == false) && (temperatureExists[2] == true || temperatureExists[3] == true))
                                    {
                                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                                    }
                                    else
                                    {
                                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                    }
                                }
                                else
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                }
                            }

                        }
                        //check if BeforeBPFlowrate exists
                        if (flowRateExists[0] == true)
                        {
                            sw.WriteLine("BeforeBPFlux ,");
                        }
                        //check if DuringBPFlowrate exists
                        if (flowRateExists[1] == true)
                        {
                            sw.WriteLine("DuringBPFlux ,");
                        }

                        //check if AfterBPFlowrate exists and if comma is needed
                        //no comma when there is no tmp's and no temperature

                        if (flowRateExists[2] == true)
                        {
                            if (temperatureExists[0] == true || temperatureExists[1] == true)
                            {
                                sw.WriteLine("AfterBPFlux ,");

                            }
                            else
                            {
                                //no comma when no temperature and no tmps
                                if (tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false)
                                {
                                    sw.WriteLine("AfterBPFlux");
                                }
                                else
                                {
                                    sw.WriteLine("AfterBPFlux ,");
                                }

                            }
                        }
                        if (flowRateExists[0] == true && tmpExists[0] == true)
                        {
                            sw.WriteLine("BeforeBPPermeability ,");
                        }
                        if (flowRateExists[1] == true && tmpExists[1] == true)
                        {
                            sw.WriteLine("DuringBPPermeability ,");
                        }
                        if (flowRateExists[2] == true && tmpExists[2] == true)
                        {
                            //no temperature then no comma
                            if (temperatureExists[0] == false && temperatureExists[1] == false)
                            {
                                sw.WriteLine("AfterBPPermeability ");
                            }
                            else
                            {
                                sw.WriteLine("AfterBPPermeability ,");
                            }
                        }

                        if (temperatureExists[0] == true || temperatureExists[1] == true)
                        {
                            if (flowRateExists[0] == true)
                            {
                                sw.WriteLine("BeforeBPTempCorrFlux ,");
                            }
                            //check if DuringBPFlowrate exists
                            if (flowRateExists[1] == true)
                            {
                                sw.WriteLine("DuringBPTempCorrFlux ,");
                            }
                            //check if AfterBPFlowrate exists
                            if (flowRateExists[2] == true)
                            {
                                if ((tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false) && (temperatureExists[2] == true || temperatureExists[3] == true))
                                {
                                    sw.WriteLine("AfterBPTempCorrFlux ");
                                }
                                else
                                {
                                    sw.WriteLine("AfterBPTempCorrFlux ,");
                                }
                            }
                            if (flowRateExists[0] == true && tmpExists[0] == true)
                            {
                                sw.WriteLine("BeforeBPTempCorrPermeability ,");
                            }
                            if (flowRateExists[1] == true && tmpExists[1] == true)
                            {
                                sw.WriteLine("DuringBPTempCorrPermeability ,");
                            }
                            if (flowRateExists[2] == true && tmpExists[2] == true)
                            {
                                //check if temerature is in the trains then no comma
                                if (temperatureExists[2] == true || temperatureExists[3] == true)
                                {
                                    sw.WriteLine("AfterBPTempCorrPermeability");
                                }
                                else
                                {
                                    sw.WriteLine("AfterBPTempCorrPermeability ,");
                                }
                            }
                        }

                        //check if there is permeate or feed temperature in common and put in train
                        if (temperatureExists[0] == true && temperatureExists[2] == false)
                        {
                            sw.WriteLine("PermeateTemperature");
                        }
                        //no permeate temp and feedtemp exists in the common but not in the trains then add it.
                        else if (temperatureExists[0] == false && temperatureExists[1] == true && temperatureExists[3] == false)
                        {
                            sw.WriteLine("FeedTemperature");
                        }

                        sw.WriteLine(")");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("SELECT DISTINCT");
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("DateandTime , ");
                        }
                        else
                        {
                            //OPC Trend
                            sw.WriteLine("CAST (LEFT(DateandTime,17) AS DATETIME) ,");
                        }
                        
                        for (int i = startListTrain; i <= endListTrain; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                
                                if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("BeforeBPTMP"))
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " * -1.0 ,");

                                }

                                else if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("AfterBPTMP"))
                                {
                                    if (i == totalTrainTags[totalNumberOfTrains - 1] && totalNumberOfTrains >= 1)
                                    {                                        
                                        //last tag to write and there are no Backpulse Flowrates then write no comma after the BeforeBPTMP

                                        if ((flowRateExists[0] == false && flowRateExists[1] == false && flowRateExists[2] == false) && databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("AfterBPTMP"))
                                        {
                                            sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " * -1.0");
                                        }
                                        //i == endListTrain1 &&
                                        else if ( (flowRateExists[0] == true || flowRateExists[1] == true || flowRateExists[2] == true) && databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("AfterBPTMP"))
                                        {
                                            sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " * -1.0 ,");

                                        }
                                    }
                                    else
                                    {
                                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " * -1.0 ,");
                                    }
                                }
                                else if (i == totalTrainTags[totalNumberOfTrains - 1] && totalNumberOfTrains >= 1 && (flowRateExists[0] == false && flowRateExists[1] == false && flowRateExists[2] == false) && (temperatureExists[0] == false && temperatureExists[1] == false))
                                {                                    
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                                }
                                //if temperature in the trains exist then no comma
                                else if (i == totalTrainTags[totalNumberOfTrains - 1] && totalNumberOfTrains >= 1 && (flowRateExists[0] == false && flowRateExists[1] == false && flowRateExists[2] == false) && (temperatureExists[2] == true || temperatureExists[3] == true))
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                                }
                                else
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                }
                            }
                        }

                        if (flowRateExists[0] == true)
                        {
                            sw.Write(sw.NewLine);
                            sw.WriteLine("--The Following values are computed");
                            sw.Write(sw.NewLine);
                            sw.WriteLine("--Flux");
                            if (flowRate == "L/s")
                            {
                                sw.WriteLine("((BeforeBPFlowRate * 3600.00) / @MembraneArea ) AS BeforeBPFlux ,  --lmh");
                            }
                            if (flowRate == "m3/h")
                            {
                                sw.WriteLine("((BeforeBPFlowRate * 1000.00) / @MembraneArea ) AS BeforeBPFlux ,  --lmh");
                            }
                            if (flowRate == "gpm")
                            {
                                sw.WriteLine("((BeforeBPFlowRate * 1440.00) / @MembraneArea ) AS BeforeBPFlux ,  --gfd");

                            }
                        }
                        //check if DuringBPFlowrate exists
                        if (flowRateExists[1] == true)
                        {
                            if (flowRate == "L/s")
                            {
                                sw.WriteLine("((DuringBPFlowRate * 3600.00) / @MembraneArea ) AS DuringBPFlux ,  --lmh");
                            }
                            if (flowRate == "m3/h")
                            {
                                sw.WriteLine("((DuringBPFlowRate * 1000.00) / @MembraneArea ) AS DuringBPFlux ,  --lmh");
                            }
                            if (flowRate == "gpm")
                            {
                                sw.WriteLine("((DuringBPFlowRate * 1440.00) / @MembraneArea ) AS DuringBPFlux ,  --gfd");
                            }
                        }

                        //check if AfterBPFlowrate exists
                        if (flowRateExists[2] == true)
                        {
                            if (flowRate == "L/s")
                            {
                                if (temperatureExists[0] == true || temperatureExists[1] == true)
                                {
                                    sw.WriteLine("((AfterBPFlowRate * 3600.00) / @MembraneArea ) AS AfterBPFlux ,  --lmh");
                                }
                                else
                                {
                                    //no comma when no temperature and no tmps
                                    if (tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false)
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 3600.00) / @MembraneArea ) AS AfterBPFlux  --lmh");
                                    }
                                    else
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 3600.00) / @MembraneArea ) AS AfterBPFlux ,  --lmh");
                                    }
                                }
                            }


                            if (flowRate == "m3/h")
                            {
                                if (temperatureExists[0] == true || temperatureExists[1] == true)
                                {
                                    sw.WriteLine("((AfterBPFlowRate * 1000.00) / @MembraneArea ) AS AfterBPFlux ,  --lmh");
                                }
                                else
                                {
                                    //no comma when no temperature and no tmps
                                    if (tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false)
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1000.00) / @MembraneArea ) AS AfterBPFlux  --lmh");
                                    }
                                    else
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1000.00) / @MembraneArea ) AS AfterBPFlux ,  --lmh");
                                    }
                                }
                            }

                            if (flowRate == "gpm")
                            {
                                if (temperatureExists[0] == true || temperatureExists[1] == true)
                                {
                                    sw.WriteLine("((AfterBPFlowRate * 1440.00) / @MembraneArea ) AS AfterBPFlux ,  --gfd");
                                }
                                else
                                {
                                    //no comma when no temperature and no tmps
                                    if (tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false)
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1440.00) / @MembraneArea ) AS AfterBPFlux --gfd");
                                    }
                                    else
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1440.00) / @MembraneArea ) AS AfterBPFlux ,  --gfd");
                                    }
                                }
                            }
                        }

                        //check if BeforeBPFlowrate exists
                        if (flowRateExists[0] == true && tmpExists[0] == true)
                        {
                            if (flowRate == "L/s")
                            {
                                sw.Write(sw.NewLine);
                                sw.WriteLine("--PERMEABILITY");
                                sw.WriteLine("--BEFORE");
                                sw.WriteLine("CASE BeforeBPTMP");
                                sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "(((BeforeBPFlowRate * 3600.00) / @MembraneArea ) / (-1.0 * BeforeBPTMP * 0.01))   --lmh/bar");
                                sw.WriteLine(tabSpace2 + "END AS BeforeBPPermeability , ");

                            }
                            if (flowRate == "m3/h")
                            {
                                sw.Write(sw.NewLine);
                                sw.WriteLine("--PERMEABILITY");
                                sw.WriteLine("--BEFORE");
                                sw.WriteLine("CASE BeforeBPTMP");
                                sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "(((BeforeBPFlowRate * 1000.00) / @MembraneArea ) / (-1.0 * BeforeBPTMP * 0.01))   --lmh/bar");
                                sw.WriteLine(tabSpace2 + "END AS BeforeBPPermeability , ");
                            }

                            if (flowRate == "gpm")
                            {
                                sw.Write(sw.NewLine);
                                sw.WriteLine("--PERMEABILITY");
                                sw.WriteLine("--BEFORE");
                                sw.WriteLine("CASE BeforeBPTMP");
                                sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "(((BeforeBPFlowRate * 1440.00) / @MembraneArea ) / (-1.0 * BeforeBPTMP ))   --gfd/psi");
                                sw.WriteLine(tabSpace2 + "END AS BeforeBPPermeability , ");
                            }
                        }

                        //check if DuringBPFlowrate exists
                        if (flowRateExists[1] == true && tmpExists[1] == true)
                        {
                            if (flowRate == "L/s")
                            {
                                sw.Write(sw.NewLine);
                                sw.WriteLine("--DURING");
                                sw.WriteLine("CASE DuringBPTMP");
                                sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "(((DuringBPFlowRate * 3600.00) / @MembraneArea ) / (DuringBPTMP * 0.01))   --lmh/bar");
                                sw.WriteLine(tabSpace2 + "END AS DuringBPPermeability , ");
                            }
                            if (flowRate == "m3/h")
                            {
                                sw.Write(sw.NewLine);
                                sw.WriteLine("--DURING");
                                sw.WriteLine("CASE DuringBPTMP");
                                sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "(((DuringBPFlowRate * 1000.00) / @MembraneArea ) / (DuringBPTMP * 0.01))   --lmh/bar");
                                sw.WriteLine(tabSpace2 + "END AS DuringBPPermeability , ");
                            }
                            if (flowRate == "gpm")
                            {
                                sw.Write(sw.NewLine);
                                sw.WriteLine("--DURING");
                                sw.WriteLine("CASE DuringBPTMP");
                                sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "(((DuringBPFlowRate * 1440.00) / @MembraneArea ) / (DuringBPTMP))   --gfd/psi");
                                sw.WriteLine(tabSpace2 + "END AS DuringBPPermeability , ");
                            }
                        }

                        //check if AfterBPFlowrate exists
                        if (flowRateExists[2] == true && tmpExists[2] == true)
                        {
                            if (flowRate == "L/s")
                            {
                                if (temperatureExists[0] == true || temperatureExists[1] == true)
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "(((AfterBPFlowRate * 3600.00) / @MembraneArea ) / (-1.0 * AfterBPTMP * 0.01))  --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS AfterBPPermeability , ");
                                }
                                else
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "(((AfterBPFlowRate * 3600.00) / @MembraneArea ) / (-1.0 * AfterBPTMP * 0.01))  --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS AfterBPPermeability");
                                }
                            }
                            if (flowRate == "m3/h")
                            {
                                if (temperatureExists[0] == true || temperatureExists[1] == true)
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "(((AfterBPFlowRate * 1000.00) / @MembraneArea ) / (-1.0 * AfterBPTMP * 0.01))   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS AfterBPPermeability , ");
                                }
                                else
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "(((AfterBPFlowRate * 1000.00) / @MembraneArea ) / (-1.0 * AfterBPTMP * 0.01))   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS AfterBPPermeability");
                                }
                            }
                            if (flowRate == "gpm")
                            {
                                if (temperatureExists[0] == true || temperatureExists[1] == true)
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "(((AfterBPFlowRate * 1440.00) / @MembraneArea ) / (-1.0 * AfterBPTMP))   --gfd/psi");
                                    sw.WriteLine(tabSpace2 + "END AS AfterBPPermeability , ");
                                }
                                else
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) THEN NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "(((AfterBPFlowRate * 1440.00) / @MembraneArea ) / (-1.0 * AfterBPTMP))   --gfd/psi");
                                    sw.WriteLine(tabSpace2 + "END AS AfterBPPermeability");
                                }
                            }
                        }

                        //temp flux
                        if (temperatureExists[0] == true || temperatureExists[1] == true)
                        {
                            //check if BeforeBPFlowrate exists
                            if (flowRateExists[0] == true)
                            {
                                sw.Write(sw.NewLine);
                                sw.WriteLine("--Temp Corrected FLUX");
                                sw.WriteLine("--Temperature in " + temperature);
                                sw.Write(sw.NewLine);
                                if (flowRate == "L/s" && temperature == "Degree C")
                                {
                                    sw.WriteLine("((BeforeBPFlowRate * 3600.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS BeforeBPTempCorrFlux  , --lmh");
                                }
                                if (flowRate == "L/s" && temperature == "Degree F")
                                {
                                    sw.WriteLine("((BeforeBPFlowRate * 3600.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS BeforeBPTempCorrFlux ,  --lmh");
                                }
                                if (flowRate == "m3/h" && temperature == "Degree C")
                                {
                                    sw.WriteLine("((BeforeBPFlowRate * 1000.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS BeforeBPTempCorrFlux  , --lmh");
                                }
                                if (flowRate == "m3/h" && temperature == "Degree F")
                                {
                                    sw.WriteLine("((BeforeBPFlowRate * 1000.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS BeforeBPTempCorrFlux ,  --lmh");
                                }
                                if (flowRate == "gpm" && temperature == "Degree F")
                                {
                                    sw.WriteLine("((BeforeBPFlowRate * 1440.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS BeforeBPTempCorrFlux ,  --gfd");
                                }
                                if (flowRate == "gpm" && temperature == "Degree C")
                                {
                                    sw.WriteLine("((BeforeBPFlowRate * 1440.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS BeforeBPTempCorrFlux  , --gfd");
                                }
                            }

                            //check if DuringBPFlowrate exists
                            if (flowRateExists[1] == true)
                            {

                                if (flowRate == "L/s" && temperature == "Degree C")
                                {
                                    sw.WriteLine("((DuringBPFlowRate * 3600.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS DuringBPTempCorrFlux  , --lmh");
                                }
                                if (flowRate == "L/s" && temperature == "Degree F")
                                {
                                    sw.WriteLine("((DuringBPFlowRate * 3600.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS DuringBPTempCorrFlux ,  --lmh");
                                }
                                if (flowRate == "m3/h" && temperature == "Degree C")
                                {
                                    sw.WriteLine("((DuringBPFlowRate * 1000.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS DuringBPTempCorrFlux  , --lmh");
                                }
                                if (flowRate == "m3/h" && temperature == "Degree F")
                                {
                                    sw.WriteLine("((DuringBPFlowRate * 1000.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS DuringBPTempCorrFlux ,  --lmh");
                                }
                                if (flowRate == "gpm" && temperature == "Degree F")
                                {
                                    sw.WriteLine("((DuringBPFlowRate * 1440.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS DuringBPTempCorrFlux ,  --gfd");
                                }
                                if (flowRate == "gpm" && temperature == "Degree C")
                                {
                                    sw.WriteLine("((DuringBPFlowRate * 1440.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS DuringBPTempCorrFlux  , --gfd");
                                }
                            }

                            //check if AfterBPFlowrate exists
                            if (flowRateExists[2] == true)
                            {

                                if (flowRate == "L/s" && temperature == "Degree C")
                                {
                                    //no comma if no tmps and temperature exists in the trains already
                                    if ((tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false) && (temperatureExists[2] == true || temperatureExists[3] == true))
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 3600.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux  --lmh");
                                    }
                                    else
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 3600.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux  , --lmh");
                                    }
                                }
                                if (flowRate == "L/s" && temperature == "Degree F")
                                {
                                    if ((tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false) && (temperatureExists[2] == true || temperatureExists[3] == true))
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 3600.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux  --lmh");
                                    }
                                    else
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 3600.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux ,  --lmh");
                                    }
                                }
                                if (flowRate == "m3/h" && temperature == "Degree C")
                                {
                                    if ((tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false) && (temperatureExists[2] == true || temperatureExists[3] == true))
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1000.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux  --lmh");
                                    }
                                    else
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1000.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux  , --lmh");
                                    }
                                }
                                if (flowRate == "m3/h" && temperature == "Degree F")
                                {
                                    if ((tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false) && (temperatureExists[2] == true || temperatureExists[3] == true))
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1000.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux  --lmh");
                                    }
                                    else
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1000.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux ,  --lmh");
                                    }
                                }
                                if (flowRate == "gpm" && temperature == "Degree F")
                                {
                                    if ((tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false) && (temperatureExists[2] == true || temperatureExists[3] == true))
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1440.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux  --gfd");
                                    }
                                    else
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1440.00) / @MembraneArea ) *  ((3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux ,  --gfd");
                                    }
                                }
                                if (flowRate == "gpm" && temperature == "Degree C")
                                {
                                    if ((tmpExists[0] == false && tmpExists[1] == false && tmpExists[2] == false) && (temperatureExists[2] == true || temperatureExists[3] == true))
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1440.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux  --gfd");
                                    }
                                    else
                                    {
                                        sw.WriteLine("((AfterBPFlowRate * 1440.00) / @MembraneArea ) *  ((1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / 0.994) AS AfterBPTempCorrFlux  , --gfd");
                                    }
                                }
                            }
                        }
                        //Temperature corrected Permeability
                        if (temperatureExists[0] == true || temperatureExists[1] == true)
                        {
                            //check if BeforeBPFlowrate exists
                            if (flowRateExists[0] == true && tmpExists[0] == true)
                            {
                                if (flowRate == "L/s" && temperature == "Degree C")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--TEMPERATURE CORRECTED PERMEABILITY");
                                    sw.WriteLine("--BEFORE");
                                    sw.WriteLine("CASE BeforeBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((BeforeBPFlowRate * 3600.00) / @MembraneArea ) * (1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / (-1.0 * BeforeBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS BeforeBPTempCorrPermeability ,");
                                }
                                if (flowRate == "L/s" && temperature == "Degree F")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--TEMPERATURE CORRECTED PERMEABILITY");
                                    sw.WriteLine("--BEFORE");
                                    sw.WriteLine("CASE BeforeBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((BeforeBPFlowRate * 3600.00) / @MembraneArea ) * (3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / (-1.0 * BeforeBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS BeforeBPTempCorrPermeability ,");
                                }
                                if (flowRate == "m3/h" && temperature == "Degree C")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--TEMPERATURE CORRECTED PERMEABILITY");
                                    sw.WriteLine("--BEFORE");
                                    sw.WriteLine("CASE BeforeBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((BeforeBPFlowRate * 1000.00) / @MembraneArea ) * (1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / (-1.0 * BeforeBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS BeforeBPTempCorrPermeability ,");
                                }
                                if (flowRate == "m3/h" && temperature == "Degree F")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--TEMPERATURE CORRECTED PERMEABILITY");
                                    sw.WriteLine("--BEFORE");
                                    sw.WriteLine("CASE BeforeBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((BeforeBPFlowRate * 1000.00) / @MembraneArea ) * (3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / (-1.0 * BeforeBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS BeforeBPTempCorrPermeability ,");
                                }
                                if (flowRate == "gpm" && temperature == "Degree F")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--TEMPERATURE CORRECTED PERMEABILITY");
                                    sw.WriteLine("--BEFORE");
                                    sw.WriteLine("CASE BeforeBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((BeforeBPFlowRate * 1440.00) / @MembraneArea ) * (3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / (-1.0 * BeforeBPTMP * 0.994)   --gfd/psi");
                                    sw.WriteLine(tabSpace2 + "END AS BeforeBPTempCorrPermeability ,");
                                }
                                if (flowRate == "gpm" && temperature == "Degree C")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--TEMPERATURE CORRECTED PERMEABILITY");
                                    sw.WriteLine("--BEFORE");
                                    sw.WriteLine("CASE BeforeBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((BeforeBPFlowRate * 1440.00) / @MembraneArea ) * (1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / (-1.0 * BeforeBPTMP * 0.994)   --gfd/psi");
                                    sw.WriteLine(tabSpace2 + "END AS BeforeBPTempCorrPermeability ,");
                                }

                            }

                            //check if DuringBPFlowrate exists
                            if (flowRateExists[1] == true && tmpExists[1] == true)
                            {
                                if (flowRate == "L/s" && temperature == "Degree C")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--DURING");
                                    sw.WriteLine("CASE DuringBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((DuringBPFlowRate * 3600.00) / @MembraneArea ) * (1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / (DuringBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS DuringBPTempCorrPermeability ,");
                                }
                                if (flowRate == "L/s" && temperature == "Degree F")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--DURING");
                                    sw.WriteLine("CASE DuringBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((DuringBPFlowRate * 3600.00) / @MembraneArea ) * (3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / (DuringBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS DuringBPTempCorrPermeability ,");
                                }
                                if (flowRate == "m3/h" && temperature == "Degree C")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--DURING");
                                    sw.WriteLine("CASE DuringBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((DuringBPFlowRate * 1000.00) / @MembraneArea ) * (1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / (DuringBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS DuringBPTempCorrPermeability ,");
                                }
                                if (flowRate == "m3/h" && temperature == "Degree F")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--DURING");
                                    sw.WriteLine("CASE DuringBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((DuringBPFlowRate * 1000.00) / @MembraneArea ) * (3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / (DuringBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    sw.WriteLine(tabSpace2 + "END AS DuringBPTempCorrPermeability ,");
                                }
                                if (flowRate == "gpm" && temperature == "Degree F")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--DURING");
                                    sw.WriteLine("CASE DuringBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((DuringBPFlowRate * 1440.00) / @MembraneArea ) * (3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / (DuringBPTMP * 0.994)  --gfd/psi");
                                    sw.WriteLine(tabSpace2 + "END AS DuringBPTempCorrPermeability ,");
                                }
                                if (flowRate == "gpm" && temperature == "Degree C")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--DURING");
                                    sw.WriteLine("CASE DuringBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((DuringBPFlowRate * 1440.00) / @MembraneArea ) * (1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / (DuringBPTMP * 0.994)   --gfd/psi");
                                    sw.WriteLine(tabSpace2 + "END AS DuringBPTempCorrPermeability ,");
                                }

                            }

                            //check if AfterBPFlowrate exists
                            if (flowRateExists[2] == true && tmpExists[2] == true)
                            {
                                if (flowRate == "L/s" && temperature == "Degree C")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((AfterBPFlowRate * 3600.00) / @MembraneArea ) * (1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / (-1.0 * AfterBPTMP * 0.01 * 0.994)  --lmh/bar");
                                    if (temperatureExists[2] == true || temperatureExists[3] == true)
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability");
                                    }
                                    else
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability ,");
                                    }
                                }
                                if (flowRate == "L/s" && temperature == "Degree F")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((AfterBPFlowRate * 3600.00) / @MembraneArea ) * (3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / (-1.0 * AfterBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    if (temperatureExists[2] == true || temperatureExists[3] == true)
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability");
                                    }
                                    else
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability ,");
                                    }
                                }
                                if (flowRate == "m3/h" && temperature == "Degree C")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((AfterBPFlowRate * 1000.00) / @MembraneArea ) * (1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / (-1.0 * AfterBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    if (temperatureExists[2] == true || temperatureExists[3] == true)
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability");
                                    }
                                    else
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability ,");
                                    }
                                }
                                if (flowRate == "m3/h" && temperature == "Degree F")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((AfterBPFlowRate * 1000.00) / @MembraneArea ) * (3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / (-1.0 * AfterBPTMP * 0.01 * 0.994)   --lmh/bar");
                                    if (temperatureExists[2] == true || temperatureExists[3] == true)
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability");
                                    }
                                    else
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability ,");
                                    }
                                }
                                if (flowRate == "gpm" && temperature == "Degree F")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((AfterBPFlowRate * 1440.00) / @MembraneArea ) * (3.21006 - (0.05894 * PermeateTemperature) + (0.000504116 * POWER(PermeateTemperature, 2)) - (0.00000171468 *  POWER(PermeateTemperature, 3))) / (-1.0 * AfterBPTMP * 0.994)   --gfd/psi");
                                    if (temperatureExists[2] == true || temperatureExists[3] == true)
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability");
                                    }
                                    else
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability ,");
                                    }
                                }
                                if (flowRate == "gpm" && temperature == "Degree C")
                                {
                                    sw.Write(sw.NewLine);
                                    sw.WriteLine("--AFTER");
                                    sw.WriteLine("CASE AfterBPTMP");
                                    sw.WriteLine(tabSpace2 + "WHEN CAST(0 AS FLOAT) then NULL");
                                    sw.WriteLine(tabSpace2 + "ELSE");
                                    sw.WriteLine(tabSpace2 + tabSpace1 + "((AfterBPFlowRate * 1440.00) / @MembraneArea ) * (1.784 - (0.0575 * PermeateTemperature) + (0.0011 * POWER(PermeateTemperature, 2)) - (0.00001 *  POWER(PermeateTemperature, 3))) / (-1.0 * AfterBPTMP * 0.994)   --gfd/psi");
                                    if (temperatureExists[2] == true || temperatureExists[3] == true)
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability");
                                    }
                                    else
                                    {
                                        sw.WriteLine(tabSpace2 + "END AS AfterBPTempCorrPermeability ,");
                                    }
                                }
                            }
                        }

                        if (flowRateExists[0] == true || flowRateExists[1] == true || flowRateExists[2] == true)
                        {
                            sw.WriteLine("--End of computed Values.");
                        }

                        //check if there is permeate or feed temperature in common and put in train

                        if (temperatureExists[0] == true && temperatureExists[2] == false)
                        {
                            sw.Write(sw.NewLine);
                            sw.WriteLine("PermeateTemperature");
                        }
                        //no permeate temp and feedtemp exists in the common but not in the trains then add it.
                        else if (temperatureExists[0] == false && temperatureExists[1] == true && temperatureExists[3] == false)
                        {
                            sw.Write(sw.NewLine);
                            sw.WriteLine("FeedTemperature");
                        }
                        sw.Write(sw.NewLine);
                        sw.WriteLine("FROM TempData" + worksheetNames[j]);
                        sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                        sw.Write(sw.NewLine);

                        //check if BeforeBPFlowrate exists
                        if (flowRateExists[0] == true && tmpExists[0] == true)
                        {
                            sw.WriteLine("--Check for bad Permeability values");
                            sw.Write(sw.NewLine);
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("UPDATE   ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j]);
                            }
                            else
                            {
                                sw.WriteLine("UPDATE   ' + @SiteName + '_PD" + worksheetNames[j]);
                            }
                            sw.WriteLine("SET       BeforeBPPermeability = NULL");
                            if (flowRate == "gpm")
                            {
                                sw.WriteLine("WHERE  (BeforeBPPermeability <= 0) OR (BeforeBPPermeability > 50)");
                            }
                            else
                            {

                                sw.WriteLine("WHERE  (BeforeBPPermeability <= 0) OR (BeforeBPPermeability > 1250)");

                            }
                            sw.Write(sw.NewLine);
                        }

                        //check if DuringBPFlowrate exists
                        if (flowRateExists[1] == true && tmpExists[1] == true)
                        {
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("UPDATE   ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j]);
                            }
                            else
                            {
                                sw.WriteLine("UPDATE   ' + @SiteName + '_PD" + worksheetNames[j]);
                            }
                            sw.WriteLine("SET       DuringBPPermeability = NULL");
                            if (flowRate == "gpm")
                            {
                                sw.WriteLine("WHERE   (DuringBPPermeability <= 0) OR (DuringBPPermeability > 50)");
                            }
                            else
                            {
                                //L/s
                                sw.WriteLine("WHERE   (DuringBPPermeability <= 0) OR (DuringBPPermeability > 1250)");
                            }
                            sw.Write(sw.NewLine);
                        }
                        //check if AfterBPFlowrate exists
                        if (flowRateExists[2] == true && tmpExists[2] == true)
                        {
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("UPDATE   ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j]);
                            }
                            else
                            {
                                sw.WriteLine("UPDATE   ' + @SiteName + '_PD" + worksheetNames[j]);
                            }

                            sw.WriteLine("SET       AfterBPPermeability = NULL");
                            if (flowRate == "gpm")
                            {
                                sw.WriteLine("WHERE   (AfterBPPermeability <= 0) OR (AfterBPPermeability > 50)");
                            }
                            else
                            {
                                //L/s
                                sw.WriteLine("WHERE   (AfterBPPermeability <= 0) OR (AfterBPPermeability > 1250)");
                            }
                            sw.Write(sw.NewLine);
                        }

                        if (temperatureExists[0] == true || temperatureExists[1] == true)
                        {

                            //check if BeforeBPFlowrate exists
                            if (flowRateExists[0] == true && tmpExists[0] == true)
                            {
                                sw.WriteLine("--Check for bad Temperature corrected Permeability values");
                                sw.Write(sw.NewLine);
                                if (aoNumberExists == true)
                                {
                                    sw.WriteLine("UPDATE   ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j]);
                                }
                                else
                                {
                                    sw.WriteLine("UPDATE   ' + @SiteName + '_PD" + worksheetNames[j]);
                                }
                                sw.WriteLine("SET       BeforeBPTempCorrPermeability = NULL");
                                if (flowRate == "gpm")
                                {
                                    sw.WriteLine("WHERE   (BeforeBPTempCorrPermeability <= 0) OR ( BeforeBPTempCorrPermeability > 50) ");
                                }
                                else
                                {
                                    //L/s
                                    sw.WriteLine("WHERE   (BeforeBPTempCorrPermeability <= 0) OR ( BeforeBPTempCorrPermeability > 1250) ");
                                }
                                sw.Write(sw.NewLine);
                            }
                            //check if DuringBPFlowrate exists
                            if (flowRateExists[1] == true && tmpExists[1] == true)
                            {
                                if (aoNumberExists == true)
                                {
                                    sw.WriteLine("UPDATE   ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j]);
                                }
                                else
                                {
                                    sw.WriteLine("UPDATE   ' + @SiteName + '_PD" + worksheetNames[j]);
                                }
                                sw.WriteLine("SET       DuringBPTempCorrPermeability = NULL");
                                if (flowRate == "gpm")
                                {
                                    sw.WriteLine("WHERE   (DuringBPTempCorrPermeability <= 0) OR (DuringBPTempCorrPermeability > 50)");
                                }
                                else
                                {
                                    //L/s
                                    sw.WriteLine("WHERE   (DuringBPTempCorrPermeability <= 0) OR (DuringBPTempCorrPermeability > 1250)");
                                }
                                sw.Write(sw.NewLine);
                            }

                            //check if AfterBPFlowrate exists
                            if (flowRateExists[2] == true && tmpExists[2] == true)
                            {
                                if (aoNumberExists == true)
                                {
                                    sw.WriteLine("UPDATE   ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j]);
                                }
                                else
                                {
                                    sw.WriteLine("UPDATE   ' + @SiteName + '_PD" + worksheetNames[j]);
                                }
                                sw.WriteLine("SET       AfterBPTempCorrPermeability = NULL");
                                if (flowRate == "gpm")
                                {
                                    sw.WriteLine("WHERE   (AfterBPTempCorrPermeability <= 0) OR (AfterBPTempCorrPermeability > 50)");
                                }
                                else
                                {
                                    //L/s
                                    sw.WriteLine("WHERE   (AfterBPTempCorrPermeability <= 0) OR (AfterBPTempCorrPermeability > 1250)");
                                }
                            }
                        }

                        sw.Write(sw.NewLine);
                        sw.WriteLine("SET NOCOUNT OFF --reenable count messages");
                        sw.Write(sw.NewLine);

                        sw.WriteLine(tabSpace2 + " '");
                        sw.WriteLine(tabSpace2 + "--PRINT (@nSQL)");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the stored procedure, up_fromTmptoProductionData" + worksheetNames[j] + ".'");
                        sw.WriteLine(tabSpace1 + "END");

                        //decrement counter to traverse through the last tag indexes of each sheet
                        totalNumberOfTrains = totalNumberOfTrains - 1;
                    }                    
                }
            }
            sw.Write(sw.NewLine);


            if (totalMit1Rows > 0)
            {
                totalMitSheets = NumberofWorksheetsWithName("MIT");
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MIT in it
                    if (worksheetNames[j].ToString().Contains("MIT"))
                    {
                        sw.WriteLine(tabSpace1 + "IF EXISTS (SELECT * FROM dbo.sysobjects");
                        sw.WriteLine(tabSpace2 + "WHERE id = object_id(N'[dbo].[up_fromTmptoProductionData" + worksheetNames[j] + "]')");
                        sw.WriteLine(tabSpace2 + "AND OBJECTPROPERTY(id, N'IsProcedure') = 1)");
                        sw.WriteLine(tabSpace1 + "PRINT 'Cannot create the stored procedure, up_fromTmptoProductionData" + worksheetNames[j] + " because it already exists.'");
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("--MIT Stored Procedures");
                        sw.WriteLine(tabSpace2 + "SET @nSQL = '");
                        if (siteAssigned == "Sandeep")
                        {
                            sw.WriteLine("--SJ");
                        }
                        else if (siteAssigned == "Saima")
                        {
                            sw.WriteLine("--SB");
                        }
                        else if (siteAssigned == "Edison")
                        {
                            sw.WriteLine("--EC");
                        }
                        else
                        {
                            sw.WriteLine("--DM");
                        }

                        sw.WriteLine("--' + CAST(DATENAME(MONTH, GETDATE()) AS VARCHAR) + ' ' + CAST(DATEPART(DAY, GETDATE())AS VARCHAR) + ', ' +  CAST(DATEPART(YEAR, GETDATE())AS VARCHAR) + '");
                        if (areaSquareFeetChecked == true)
                        {
                            sw.WriteLine("--Units in imperial");
                        }
                        else
                        {
                            sw.WriteLine("--Units in metric");
                        }

                        sw.WriteLine("CREATE PROCEDURE up_fromTmptoProductionData" + worksheetNames[j]);
                        sw.WriteLine("AS");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("SET NOCOUNT ON -- Capture status meassages");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("--delete rows that have no date and time value");
                        sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                        sw.WriteLine("WHERE DateandTime IS NULL");
                        sw.Write(sw.NewLine);
                        //remove - causing data gaps
                        /*
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that have duplicate date and time stamps");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM TempData" + worksheetNames[j] + " GROUP BY DateandTime having Count(*)>1)");
                        }
                        else
                        {
                            sw.WriteLine("--delete rows that have duplicate date and time stamps");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempData" + worksheetNames[j] + " GROUP BY CAST (LEFT(DateandTime,17) AS DATETIME) having Count(*)>1)");
                        }
                        sw.Write(sw.NewLine);
                        */

                        /*
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " WHERE DateandTime IN (SELECT DateandTime FROM TempData" + worksheetNames[j] + "))");
                        }
                        else
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " WHERE DateandTime IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempData" + worksheetNames[j] + "))");
                        }
                        sw.Write(sw.NewLine);
                        */
                         //Adding new statement                        
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ")");
                            }
                            else
                            {
                                sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_PD" + worksheetNames[j] + ")");
                            }
          
                        }
                        else
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ")");
                            }
                            else
                            {
                                sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_PD" + worksheetNames[j] + ")");
                            }

                           
                        }
                        
                        sw.Write(sw.NewLine);
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine("INSERT INTO ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " (");
                        }
                        else
                        {
                            sw.WriteLine("INSERT INTO ' + @SiteName + '_PD" + worksheetNames[j] + " (");
                        }
                        sw.WriteLine("DateandTime , ");

                        //only put in Pressure Difference if there is start pressure and end pressure
                        if (pressureDifferenceExists == false)
                        {
                            for (int i = startListMit; i <= endListMit; i++)
                            {
                                //remove sheetname from the tagname when printing the tag
                                startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                                if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                                {
                                    if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("StartPressure"))
                                    {
                                        for (int k = startListMit; k <= endListMit; k++)
                                        {
                                            //remove sheetname from the tagname when printing the tag
                                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                                            //only want to go through each worksheet (ie MIT1 - should only go through MIT1 tags)
                                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[k]["Tag Name"].ToString().Substring(0, startPoint))
                                            {
                                                if (databaseTagsTable.Rows[k]["Tag Name"].ToString().Contains("EndPressure") || databaseTagsTable.Rows[k]["Tag Name"].ToString().Contains("FinishPressure"))
                                                {
                                                    sw.WriteLine("PressureDifference , ");
                                                    // break;
                                                }
                                            }

                                        }
                                        // break;
                                    }

                                }
                            }
                        }

                        
                        for (int i = startListMit; i <= endListMit; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the MIT then put no comma
                                if (i == totalMitTags[totalMitSheets - 1] && totalMitSheets >= 1)
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                                }
                                else
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                }
                            }

                        }
                        sw.WriteLine(")");
                        sw.Write(sw.NewLine);

                        sw.WriteLine("SELECT DISTINCT");
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("DateandTime , ");
                        }
                        else
                        {
                            //OPC Trend
                            sw.WriteLine("CAST (LEFT(DateandTime,17) AS DATETIME) ,");
                        }
                        //add pressure difference calculation
                        //only put in Pressure Difference if there is start pressure and end pressure
                        if (pressureDifferenceExists == false)
                        {
                            for (int i = startListMit; i <= endListMit; i++)
                            {
                                startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                                if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                                {
                                    if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("StartPressure"))
                                    {
                                        for (int k = startListMit; k <= endListMit; k++)
                                        {
                                            startPoint = databaseTagsTable.Rows[k]["Tag Name"].ToString().IndexOf(".");
                                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[k]["Tag Name"].ToString().Substring(0, startPoint))
                                            {
                                                if (databaseTagsTable.Rows[k]["Tag Name"].ToString().Contains("FinishPressure"))
                                                {
                                                    sw.WriteLine("(StartPressure - FinishPressure) AS PressureDifference  ,");
                                                }
                                                else if (databaseTagsTable.Rows[k]["Tag Name"].ToString().Contains("EndPressure"))
                                                {
                                                    sw.WriteLine("(StartPressure - EndPressure) AS PressureDifference  ,");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                       
                        for (int i = startListMit; i <= endListMit; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the MIT then put no comma
                                if (i == totalMitTags[totalMitSheets - 1] && totalMitSheets >= 1)
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                                }
                                else
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                }
                            }
                        }

                        sw.Write(sw.NewLine);
                        sw.WriteLine("FROM TempData" + worksheetNames[j]);
                        sw.Write(sw.NewLine);
                        sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                        sw.Write(sw.NewLine);
                        sw.WriteLine("SET NOCOUNT OFF --reenable count messages");
                        sw.WriteLine(tabSpace2 + "'");
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the stored procedure, up_fromTmptoProductionData" + worksheetNames[j] + ".'");
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                        //decrement counter to traverse through the last tag indexes of each sheet
                        totalMitSheets = totalMitSheets - 1;

                    }
                }
            }

            //MC Stored Procedure
            if (totalMCRows > 0)
            {
                totalMCSheets = NumberofWorksheetsWithName("MC");
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MC in it
                    if (worksheetNames[j].ToString().Contains("MC"))
                    {
                        sw.WriteLine(tabSpace1 + "IF EXISTS (SELECT * FROM dbo.sysobjects");
                        sw.WriteLine(tabSpace2 + "WHERE id = object_id(N'[dbo].[up_fromTmptoProductionData" + worksheetNames[j] + "]')");
                        sw.WriteLine(tabSpace2 + "AND OBJECTPROPERTY(id, N'IsProcedure') = 1)");
                        sw.WriteLine(tabSpace1 + "PRINT 'Cannot create the stored procedure, up_fromTmptoProductionData" + worksheetNames[j] + " because it already exists.'");
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("--MC Stored Procedures");
                        sw.WriteLine(tabSpace2 + "SET @nSQL = '");
                        if (siteAssigned == "Sandeep")
                        {
                            sw.WriteLine("--SJ");
                        }
                        else if (siteAssigned == "Saima")
                        {
                            sw.WriteLine("--SB");
                        }
                        else if (siteAssigned == "Edison")
                        {
                            sw.WriteLine("--EC");
                        }
                        else
                        {
                            sw.WriteLine("--DM");
                        }

                        sw.WriteLine("--' + CAST(DATENAME(MONTH, GETDATE()) AS VARCHAR) + ' ' + CAST(DATEPART(DAY, GETDATE())AS VARCHAR) + ', ' +  CAST(DATEPART(YEAR, GETDATE())AS VARCHAR) + '");
                        if (areaSquareFeetChecked == true)
                        {
                            sw.WriteLine("--Units in imperial");
                        }
                        else
                        {
                            sw.WriteLine("--Units in metric");
                        }

                        sw.WriteLine("CREATE PROCEDURE up_fromTmptoProductionData" + worksheetNames[j]);
                        sw.WriteLine("AS");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("SET NOCOUNT ON -- Capture status meassages");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("--delete rows that have no date and time value");
                        sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                        sw.WriteLine("WHERE DateandTime IS NULL");
                        sw.Write(sw.NewLine);
                        //remove - causing data gaps
                        /*
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that have duplicate date and time stamps");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM TempData" + worksheetNames[j] + " GROUP BY DateandTime having Count(*)>1)");
                        }
                        else
                        {
                            sw.WriteLine("--delete rows that have duplicate date and time stamps");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempData" + worksheetNames[j] + " GROUP BY CAST (LEFT(DateandTime,17) AS DATETIME) having Count(*)>1)");
                        }
                        sw.Write(sw.NewLine);
                        */

                        /*
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " WHERE DateandTime IN (SELECT DateandTime FROM TempData" + worksheetNames[j] + "))");
                        }
                        else
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " WHERE DateandTime IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempData" + worksheetNames[j] + "))");
                        }
                        sw.Write(sw.NewLine);
                        */
                        //Adding new statement                        
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ")");
                            }
                            else
                            {
                                sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_PD" + worksheetNames[j] + ")");
                            }

                        }
                        else
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ")");
                            }
                            else
                            {
                                sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_PD" + worksheetNames[j] + ")");
                            }


                        }

                        sw.Write(sw.NewLine);
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine("INSERT INTO ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " (");
                        }
                        else
                        {
                            sw.WriteLine("INSERT INTO ' + @SiteName + '_PD" + worksheetNames[j] + " (");
                        }
                        sw.WriteLine("DateandTime , ");


                        for (int i = startListMC; i <= endListMC; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the MC then put no comma
                                if (i == totalMCTags[totalMCSheets - 1] && totalMCSheets >= 1)
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                                }
                                else
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                }
                            }

                        }
                        sw.WriteLine(")");
                        sw.Write(sw.NewLine);

                        sw.WriteLine("SELECT DISTINCT");
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("DateandTime , ");
                        }
                        else
                        {
                            //OPC Trend
                            sw.WriteLine("CAST (LEFT(DateandTime,17) AS DATETIME) ,");
                        }

                        for (int i = startListMC; i <= endListMC; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the MC then put no comma
                                if (i == totalMCTags[totalMCSheets - 1] && totalMCSheets >= 1)
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                                }
                                else
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                }
                            }
                        }

                        sw.Write(sw.NewLine);
                        sw.WriteLine("FROM TempData" + worksheetNames[j]);
                        sw.Write(sw.NewLine);
                        sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                        sw.Write(sw.NewLine);
                        sw.WriteLine("SET NOCOUNT OFF --reenable count messages");
                        sw.WriteLine(tabSpace2 + "'");
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the stored procedure, up_fromTmptoProductionData" + worksheetNames[j] + ".'");
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                        //decrement counter to traverse through the last tag indexes of each sheet
                        totalMCSheets = totalMCSheets - 1;

                    }
                }
            }

            //RC Stored Procedure
            if (totalRCRows > 0)
            {
                totalRCSheets = NumberofWorksheetsWithName("RC");
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with RC in it
                    if (worksheetNames[j].ToString().Contains("RC"))
                    {
                        sw.WriteLine(tabSpace1 + "IF EXISTS (SELECT * FROM dbo.sysobjects");
                        sw.WriteLine(tabSpace2 + "WHERE id = object_id(N'[dbo].[up_fromTmptoProductionData" + worksheetNames[j] + "]')");
                        sw.WriteLine(tabSpace2 + "AND OBJECTPROPERTY(id, N'IsProcedure') = 1)");
                        sw.WriteLine(tabSpace1 + "PRINT 'Cannot create the stored procedure, up_fromTmptoProductionData" + worksheetNames[j] + " because it already exists.'");
                        sw.WriteLine(tabSpace1 + "ELSE");
                        sw.WriteLine(tabSpace1 + "BEGIN");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("--RC Stored Procedures");
                        sw.WriteLine(tabSpace2 + "SET @nSQL = '");
                        if (siteAssigned == "Sandeep")
                        {
                            sw.WriteLine("--SJ");
                        }
                        else if (siteAssigned == "Saima")
                        {
                            sw.WriteLine("--SB");
                        }
                        else if (siteAssigned == "Edison")
                        {
                            sw.WriteLine("--EC");
                        }
                        else
                        {
                            sw.WriteLine("--DM");
                        }

                        sw.WriteLine("--' + CAST(DATENAME(MONTH, GETDATE()) AS VARCHAR) + ' ' + CAST(DATEPART(DAY, GETDATE())AS VARCHAR) + ', ' +  CAST(DATEPART(YEAR, GETDATE())AS VARCHAR) + '");
                        if (areaSquareFeetChecked == true)
                        {
                            sw.WriteLine("--Units in imperial");
                        }
                        else
                        {
                            sw.WriteLine("--Units in metric");
                        }

                        sw.WriteLine("CREATE PROCEDURE up_fromTmptoProductionData" + worksheetNames[j]);
                        sw.WriteLine("AS");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("SET NOCOUNT ON -- Capture status meassages");
                        sw.Write(sw.NewLine);
                        sw.WriteLine("--delete rows that have no date and time value");
                        sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                        sw.WriteLine("WHERE DateandTime IS NULL");
                        sw.Write(sw.NewLine);
                        //remove - causing data gaps
                        /*
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that have duplicate date and time stamps");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM TempData" + worksheetNames[j] + " GROUP BY DateandTime having Count(*)>1)");
                        }
                        else
                        {
                            sw.WriteLine("--delete rows that have duplicate date and time stamps");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempData" + worksheetNames[j] + " GROUP BY CAST (LEFT(DateandTime,17) AS DATETIME) having Count(*)>1)");
                        }
                        sw.Write(sw.NewLine);
                        */

                        /*
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " WHERE DateandTime IN (SELECT DateandTime FROM TempData" + worksheetNames[j] + "))");
                        }
                        else
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " WHERE DateandTime IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempData" + worksheetNames[j] + "))");
                        }
                        sw.Write(sw.NewLine);
                        */
                        //Adding new statement                        
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ")");
                            }
                            else
                            {
                                sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_PD" + worksheetNames[j] + ")");
                            }

                        }
                        else
                        {
                            sw.WriteLine("--delete rows that already exist in the production table");
                            sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                            if (aoNumberExists == true)
                            {
                                sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + ")");
                            }
                            else
                            {
                                sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_PD" + worksheetNames[j] + ")");
                            }


                        }

                        sw.Write(sw.NewLine);
                        if (aoNumberExists == true)
                        {
                            sw.WriteLine("INSERT INTO ' + @SiteName + '_' + @AONumber + '_PD" + worksheetNames[j] + " (");
                        }
                        else
                        {
                            sw.WriteLine("INSERT INTO ' + @SiteName + '_PD" + worksheetNames[j] + " (");
                        }
                        sw.WriteLine("DateandTime , ");


                        for (int i = startListRC; i <= endListRC; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the RC then put no comma
                                if (i == totalRCTags[totalRCSheets - 1] && totalRCSheets >= 1)
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                                }
                                else
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                }
                            }

                        }
                        sw.WriteLine(")");
                        sw.Write(sw.NewLine);

                        sw.WriteLine("SELECT DISTINCT");
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine("DateandTime , ");
                        }
                        else
                        {
                            //OPC Trend
                            sw.WriteLine("CAST (LEFT(DateandTime,17) AS DATETIME) ,");
                        }

                        for (int i = startListRC; i <= endListRC; i++)
                        {
                            //remove sheetname from the tagname when printing the tag
                            startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                            if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                            {
                                //when you get to the last tag for the RC then put no comma
                                if (i == totalRCTags[totalRCSheets - 1] && totalRCSheets >= 1)
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                                }
                                else
                                {
                                    sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                }
                            }
                        }

                        sw.Write(sw.NewLine);
                        sw.WriteLine("FROM TempData" + worksheetNames[j]);
                        sw.Write(sw.NewLine);
                        sw.WriteLine("DELETE FROM TempData" + worksheetNames[j]);
                        sw.Write(sw.NewLine);
                        sw.WriteLine("SET NOCOUNT OFF --reenable count messages");
                        sw.WriteLine(tabSpace2 + "'");
                        sw.WriteLine(tabSpace2 + "--PRINT @nSQL");
                        sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                        sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                        sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                        sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                        sw.WriteLine(tabSpace2 + "ELSE");
                        sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the stored procedure, up_fromTmptoProductionData" + worksheetNames[j] + ".'");
                        sw.WriteLine(tabSpace1 + "END");
                        sw.Write(sw.NewLine);
                        //decrement counter to traverse through the last tag indexes of each sheet
                        totalRCSheets = totalRCSheets - 1;

                    }
                }
            }


            //Common 

            if (totalCommonRows > 0)
            {
                sw.WriteLine(tabSpace1 + "IF EXISTS (SELECT * FROM dbo.sysobjects");
                sw.WriteLine(tabSpace2 + "WHERE id = object_id(N'[dbo].[up_fromTmptoProductionDataCommon]')");
                sw.WriteLine(tabSpace2 + "AND OBJECTPROPERTY(id, N'IsProcedure') = 1)");
                sw.WriteLine(tabSpace1 + "PRINT 'Cannot create the stored procedure, up_fromTmptoProductionDataCommon because it already exists.'");
                sw.WriteLine(tabSpace1 + "ELSE");
                sw.WriteLine(tabSpace1 + "BEGIN");
                sw.Write(sw.NewLine);
                sw.WriteLine(tabSpace2 + "SET @nSQL = '");
                if (siteAssigned == "Sandeep")
                {
                    sw.WriteLine("--SJ");
                }
                else if (siteAssigned == "Saima")
                {
                    sw.WriteLine("--SB");
                }
                else if (siteAssigned == "Edison")
                {
                    sw.WriteLine("--EC");
                }
                else
                {
                    sw.WriteLine("--DM");
                }
                sw.WriteLine("--' + CAST(DATENAME(MONTH, GETDATE()) AS VARCHAR) + ' ' + CAST(DATEPART(DAY, GETDATE())AS VARCHAR) + ', ' +  CAST(DATEPART(YEAR, GETDATE())AS VARCHAR) + '");
                if (areaSquareFeetChecked == true)
                {
                    sw.WriteLine("--Units in imperial");
                }
                else
                {
                    sw.WriteLine("--Units in metric");
                }
                sw.Write(sw.NewLine);
                sw.WriteLine("CREATE PROCEDURE up_fromTmptoProductionDataCommon");
                sw.WriteLine("AS");
                sw.Write(sw.NewLine);
                sw.WriteLine("SET NOCOUNT ON -- Capture status meassages");
                sw.Write(sw.NewLine);
                sw.WriteLine("--delete rows that have no date and time value");
                sw.WriteLine("DELETE FROM TempDataCommon");
                sw.WriteLine("WHERE DateandTime IS NULL");
                sw.Write(sw.NewLine);
                /*
                if (softwareType == "IFIX")
                {
                    sw.WriteLine("--delete rows that have duplicate date and time stamps");
                    sw.WriteLine("DELETE FROM TempDataCommon");
                    sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM TempDataCommon GROUP BY DateandTime having Count(*)>1)");
                }
                else
                {
                    sw.WriteLine("--delete rows that have duplicate date and time stamps");
                    sw.WriteLine("DELETE FROM TempDataCommon");
                    sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempDataCommon GROUP BY CAST (LEFT(DateandTime,17) AS DATETIME) having Count(*)>1)");
                }
                sw.Write(sw.NewLine);
                 */
                /*
                if (softwareType == "IFIX")
                {
                    sw.WriteLine("--delete rows that already exist in the production table");
                    sw.WriteLine("DELETE FROM TempDataCommon");
                    sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PDCommon WHERE DateandTime IN (SELECT DateandTime FROM TempDataCommon))");
                }
                else
                {
                    sw.WriteLine("--delete rows that already exist in the production table");
                    sw.WriteLine("DELETE FROM TempDataCommon");
                    sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PDCommon WHERE DateandTime IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempDataCommon))");
                }

                sw.Write(sw.NewLine);
                */

                 //Adding new statement                        
                if (softwareType == "IFIX")
                {
                    sw.WriteLine("--delete rows that already exist in the production table");
                    sw.WriteLine("DELETE FROM TempDataCommon");
                    if (aoNumberExists == true)
                    {
                        sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PDCommon)");
                    }
                    else
                    {
                        sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_PDCommon)");
                    }
                }
                else
                {
                    sw.WriteLine("--delete rows that already exist in the production table");
                    sw.WriteLine("DELETE FROM TempDataCommon");
                    if (aoNumberExists == true)
                    {
                        sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PDCommon)");
                    }
                    else
                    {
                        sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_PDCommon)");
                    }
                        
                }
                sw.Write(sw.NewLine);
                if (aoNumberExists == true)
                {
                    sw.WriteLine("INSERT INTO ' + @SiteName + '_' + @AONumber + '_PDCommon (");
                }
                else
                {
                    sw.WriteLine("INSERT INTO ' + @SiteName + '_PDCommon (");
                }
                //add common tags
                sw.WriteLine("DateandTime , ");
                for (int i = startListTrain; i <= endListTrain; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    if (commonTags[0] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateFlow"))
                        {
                            sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + " ,");

                        }
                    }
                    if (commonTags[1] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateTurbidity"))
                        {
                            sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + " ,");

                        }
                    }
                    if (commonTags[2] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("MembraneTankLevel"))
                        {
                            sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + " ,");

                        }
                    }


                }

                //write the common tags 
                for (int i = startListCommon; i <= endListCommon; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    //Last tag will have no comma at the end
                    if (i == endListCommon)
                    {
                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));

                    }
                    else
                    {
                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");

                    }

                }
                sw.WriteLine(")");
                sw.Write(sw.NewLine);

                sw.WriteLine("SELECT DISTINCT");
                if (softwareType == "IFIX")
                {
                    sw.WriteLine("DateandTime , ");
                }
                else
                {
                    //OPC Trend
                    sw.WriteLine("CAST (LEFT(DateandTime,17) AS DATETIME) ,");
                }
                //add tags
                for (int i = startListTrain; i <= endListTrain; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    if (commonTags[0] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateFlow"))
                        {
                            sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + " ,");

                        }
                    }
                    if (commonTags[1] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("PermeateTurbidity"))
                        {
                            sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + " ,");

                        }
                    }
                    if (commonTags[2] == false)
                    {
                        if (databaseTagsTable.Rows[i]["Tag Name"].ToString().Contains("MembraneTankLevel"))
                        {
                            sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, 3) + " ,");

                        }
                    }


                }

                //write the common tags 
                for (int i = startListCommon; i <= endListCommon; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    //Last tag will have no comma at the end
                    if (i == endListCommon)
                    {
                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));

                    }
                    else
                    {
                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");

                    }

                }
                sw.Write(sw.NewLine);
                sw.WriteLine("FROM TempDataCommon");
                sw.Write(sw.NewLine);
                sw.WriteLine("DELETE FROM TempDataCommon");
                sw.Write(sw.NewLine);
                sw.WriteLine("SET NOCOUNT OFF --reenable count messages");
                sw.WriteLine(tabSpace2 + "'");
                sw.WriteLine(tabSpace2 + "--PRINT (@nSQL)");
                sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                sw.WriteLine(tabSpace2 + "ELSE");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the stored procedure, up_fromTmptoProductionDataCommon.'");
                sw.WriteLine(tabSpace1 + "END");
            }

            //Daily 

            if (totalDailyRows > 0)
            {
                sw.WriteLine(tabSpace1 + "IF EXISTS (SELECT * FROM dbo.sysobjects");
                sw.WriteLine(tabSpace2 + "WHERE id = object_id(N'[dbo].[up_fromTmptoProductionDataDaily]')");
                sw.WriteLine(tabSpace2 + "AND OBJECTPROPERTY(id, N'IsProcedure') = 1)");
                sw.WriteLine(tabSpace1 + "PRINT 'Cannot create the stored procedure, up_fromTmptoProductionDataDaily because it already exists.'");
                sw.WriteLine(tabSpace1 + "ELSE");
                sw.WriteLine(tabSpace1 + "BEGIN");
                sw.Write(sw.NewLine);
                sw.WriteLine(tabSpace2 + "SET @nSQL = '");

                if (siteAssigned == "Sandeep")
                {
                    sw.WriteLine("--SJ");
                }
                else if (siteAssigned == "Saima")
                {
                    sw.WriteLine("--SB");
                }
                else if (siteAssigned == "Edison")
                {
                    sw.WriteLine("--EC");
                }
                else
                {
                    sw.WriteLine("--DM");
                }
                sw.WriteLine("--' + CAST(DATENAME(MONTH, GETDATE()) AS VARCHAR) + ' ' + CAST(DATEPART(DAY, GETDATE())AS VARCHAR) + ', ' +  CAST(DATEPART(YEAR, GETDATE())AS VARCHAR) + '");
                if (areaSquareFeetChecked == true)
                {
                    sw.WriteLine("--Units in imperial");
                }
                else
                {
                    sw.WriteLine("--Units in metric");
                }
                sw.Write(sw.NewLine);
                sw.WriteLine("CREATE PROCEDURE up_fromTmptoProductionDataDaily");
                sw.WriteLine("AS");
                sw.Write(sw.NewLine);
                sw.WriteLine("SET NOCOUNT ON -- Capture status meassages");
                sw.Write(sw.NewLine);
                sw.WriteLine("--delete rows that have no date and time value");
                sw.WriteLine("DELETE FROM TempDataDaily");
                sw.WriteLine("WHERE DateandTime IS NULL");
                sw.Write(sw.NewLine);
                /*
                if (softwareType == "IFIX")
                {
                    sw.WriteLine("--delete rows that have duplicate date and time stamps");
                    sw.WriteLine("DELETE FROM TempDataDaily");
                    sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM TempDataDaily GROUP BY DateandTime having Count(*)>1)");
                }
                else
                {
                    sw.WriteLine("--delete rows that have duplicate date and time stamps");
                    sw.WriteLine("DELETE FROM TempDataDaily");
                    sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempDataDaily GROUP BY CAST (LEFT(DateandTime,17) AS DATETIME) having Count(*)>1)");
                }
                sw.Write(sw.NewLine);
                 */
                /*
                if (softwareType == "IFIX")
                {
                    sw.WriteLine("--delete rows that already exist in the production table");
                    sw.WriteLine("DELETE FROM TempDataDaily");
                    sw.WriteLine("WHERE DateandTime IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PDDaily WHERE DateandTime IN (SELECT DateandTime FROM TempDataDaily))");
                }
                else
                {
                    sw.WriteLine("--delete rows that already exist in the production table");
                    sw.WriteLine("DELETE FROM TempDataDaily");
                    sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) IN (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PDDaily WHERE DateandTime IN (SELECT CAST (LEFT(DateandTime,17) AS DATETIME) FROM TempDataDaily))");
                }
                sw.Write(sw.NewLine);
                 */
                //Adding new statement                        
                if (softwareType == "IFIX")
                {
                    sw.WriteLine("--delete rows that already exist in the production table");
                    sw.WriteLine("DELETE FROM TempDataDaily");
                    if (aoNumberExists == true)
                    {
                        sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PDDaily)");
                    }
                    else
                    {
                        sw.WriteLine("WHERE DateandTime = ANY (SELECT DateandTime FROM ' + @SiteName + '_PDDaily)");
                    }
                }
                else
                {
                    sw.WriteLine("--delete rows that already exist in the production table");
                    sw.WriteLine("DELETE FROM TempDataDaily");
                    if (aoNumberExists == true)
                    {
                        sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_' + @AONumber + '_PDDaily)");
                    }
                    else
                    {
                        sw.WriteLine("WHERE CAST (LEFT(DateandTime,17) AS DATETIME) = ANY (SELECT DateandTime FROM ' + @SiteName + '_PDDaily)");
                    }

                }
                sw.Write(sw.NewLine);
                if (aoNumberExists == true)
                {
                    sw.WriteLine("INSERT INTO ' + @SiteName + '_' + @AONumber + '_PDDaily (");
                }
                else
                {
                    sw.WriteLine("INSERT INTO ' + @SiteName + '_PDDaily (");
                }
                //add common tags
                sw.WriteLine("DateandTime , ");

                for (int i = startListDaily; i <= endListDaily; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    if (i == endListDaily)
                    {
                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));
                       
                    }
                    else
                    {
                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");

                    }
                }
                sw.WriteLine(")");
                sw.Write(sw.NewLine);

                sw.WriteLine("SELECT DISTINCT");
                if (softwareType == "IFIX")
                {
                    sw.WriteLine("DateandTime , ");
                }
                else
                {
                    //OPC Trend
                    sw.WriteLine("CAST (LEFT(DateandTime,17) AS DATETIME) ,");
                }
                /*Not using this - Was used for recovery calculation - Feb 15, 2011
                string sumoftrains = "";
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    sumoftrains = sumoftrains + "TotalDailyPermeateFlow" + worksheetNames[j];
                }

                 */
                //add tags


                for (int i = startListDaily; i <= endListDaily; i++)
                {
                    //remove sheetname from the tagname when printing the tag
                    startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                    if (i == endListDaily)
                    {
                     
                        /* Don't do recovery this way - taking it out
                        //check if we need to include PlantRecovery 
                        //if feed flow is true
                        if (totalDailyFlowExists[0] == true)
                        {
                            //if Plant Daily Permeate Flow is true
                            if (totalDailyFlowExists[3] == true)
                            {
                                sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                sw.WriteLine("CASE");
                                sw.WriteLine(tabSpace2 + "WHEN TotalDailyPlantFeedFlow <= 0 THEN NULL");
                                sw.WriteLine(tabSpace2 + "WHEN TotalDailyPlantFeedFlow < TotalDailyPlantPermeateFlow  THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "(TotalDailyPlantPermeateFlow / TotalDailyPlantFeedFlow * 100) END AS PlantRecovery");
                            }
                            else if (totalDailyFlowExists[4] == true)
                            {
                                sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                sw.WriteLine("CASE");
                                sw.WriteLine(tabSpace2 + "WHEN TotalDailyPlantFeedFlow <= 0 THEN NULL");
                                sw.WriteLine(tabSpace2 + "WHEN TotalDailyPlantFeedFlow < TotalPlantDailyPermeateFlow  THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "(TotalPlantDailyPermeateFlow / TotalDailyPlantFeedFlow * 100) END AS PlantRecovery");

                            }
                            else
                            {
                                sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                sw.WriteLine("CASE");
                                sw.WriteLine(tabSpace2 + "WHEN TotalDailyPlantFeedFlow <= 0 THEN NULL");
                                sw.WriteLine(tabSpace2 + "WHEN TotalDailyPlantFeedFlow < (" + sumoftrains + ")  THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");

                                sw.WriteLine(tabSpace2 + tabSpace1 + "((" + sumoftrains + ") / TotalDailyPlantFeedFlow * 100) END AS PlantRecovery");
                            }
                        }

                            //Reject Flow is true
                        else if (totalDailyFlowExists[1] == true)
                        {
                            //if Daily Plant Permeate Flow is true
                            if (totalDailyFlowExists[3] == true)
                            {
                                sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                sw.WriteLine("CASE");
                                sw.WriteLine(tabSpace2 + "WHEN (TotalDailyPlantRejectFlow + TotalDailyPermeateFlow) <= 0 THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "((TotalDailyPlantPermeateFlow / (TotalDailyPlantRejectFlow + TotalDailyPlantPermeateFlow)) * 100) END AS PlantRecovery");
                            }
                            //if Plant daily permeate flow is true
                            else if (totalDailyFlowExists[4] == true)
                            {
                                sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                sw.WriteLine("CASE");
                                sw.WriteLine(tabSpace2 + "WHEN (TotalDailyPlantRejectFlow + TotalDailyPermeateFlow) <= 0 THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "((TotalPlantDailyPermeateFlow / (TotalDailyPlantRejectFlow + TotalPlantDailyPermeateFlow)) * 100) END AS PlantRecovery");

                            }
                            else
                            {
                                sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");
                                sw.WriteLine("CASE");
                                sw.WriteLine(tabSpace2 + "WHEN TotalDailyPlantRejectFlow <= 0 THEN NULL");
                                sw.WriteLine(tabSpace2 + "ELSE");
                                sw.WriteLine(tabSpace2 + tabSpace1 + "(((" + sumoftrains + ") / (TotalDailyPlantRejectFlow + " + sumoftrains + ")) * 100) END AS PlantRecovery");
                            }
                        }

                        */
                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1));                     
                    }
                    else
                    {
                        sw.WriteLine(databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(startPoint + 1) + " ,");

                    }
                }

                sw.Write(sw.NewLine);
                sw.WriteLine("FROM TempDataDaily");
                sw.Write(sw.NewLine);
                sw.WriteLine("DELETE FROM TempDataDaily");
                sw.Write(sw.NewLine);
                sw.WriteLine("SET NOCOUNT OFF --reenable count messages");
                sw.WriteLine(tabSpace2 + "'");
                sw.Write(sw.NewLine);
                sw.WriteLine(tabSpace2 + "--PRINT  @nSQL");
                sw.WriteLine(tabSpace2 + "EXECUTE (@nSQL)");
                sw.WriteLine(tabSpace2 + "-- Save any non-zero @@ERROR value.");
                sw.WriteLine(tabSpace2 + "SET @ErrorSave = @@ERROR");
                sw.WriteLine(tabSpace2 + "IF (@ErrorSave <> 0)");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Error Number:' + STR(@ErrorSave)");
                sw.WriteLine(tabSpace2 + "ELSE");
                sw.WriteLine(tabSpace2 + tabSpace1 + "PRINT 'Successfully created the stored procedure, up_fromTmptoProductionDataDaily.'");
                sw.WriteLine(tabSpace1 + "END");

            }
            sw.Write(sw.NewLine);
            sw.WriteLine("END");
            sw.WriteLine("GO");
            sw.Write(sw.NewLine);

            //update DataCleaner tables
            sw.WriteLine("DECLARE");
            sw.WriteLine("@DatabaseName VARCHAR (100),");
            sw.WriteLine("@AONumber VARCHAR (100),");
            sw.WriteLine("@SiteName VARCHAR (100),");
            sw.WriteLine("@DTSNamesRow  BIT,");
            sw.WriteLine("@ErrorSave INT,");
            sw.WriteLine("@TempTablesRows BIT,");
            sw.WriteLine("@IncomingEmailRow BIT,");
            sw.WriteLine("@DTSInformationRows BIT,");
            sw.WriteLine("@nSQL VARCHAR (5000)");
            sw.Write(sw.NewLine);
            sw.WriteLine("--set these for every new site");
            if (aoNumberExists == true)
            {
                sw.WriteLine("SET @DatabaseName = '" + aoNumber + siteName + "'");
            }
            else
            {
                sw.WriteLine("SET @DatabaseName = '" + siteName + "'");
            }
            sw.WriteLine("SET @AONumber = '" + aoNumber + "'");
            sw.WriteLine("SET @SiteName = '" + siteName + "'");
            sw.Write(sw.NewLine);
            sw.WriteLine("SET @DTSNamesRow  = 0 -- 0 - does not exist");
            sw.WriteLine("SET @TempTablesRows = 0 --0 - does not exist");
            sw.WriteLine("SET @IncomingEmailRow = 0");
            sw.WriteLine("SET @DTSInformationRows = 0");
            sw.WriteLine("SET @ErrorSave = 0");
            sw.WriteLine("SET @nSQL = ''");
            sw.Write(sw.NewLine);
            sw.WriteLine("IF EXISTS (SELECT [DatabaseName] FROM [DataCleaner].[dbo].[ZenoTracDatabasesAndTempTables]");
            sw.WriteLine("WHERE [DatabaseName] = @DatabaseName)");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "SET @TempTablesRows = 1");
            sw.WriteLine(tabSpace1 + "PRINT 'Cannot Insert the rows because the database name, ' + @DatabaseName + ' already exists in the ZenoTracDatabasesAndTempTables table .'");
            sw.WriteLine("END");
            sw.Write(sw.NewLine);

            sw.WriteLine("IF @TempTablesRows  = 0");
            sw.WriteLine("BEGIN");
            sw.Write(sw.NewLine);
            sw.WriteLine(tabSpace1 + "SET @nSQL = 'INSERT INTO [DataCleaner].[dbo].[ZenoTracDatabasesAndTempTables]([DatabaseName], [TableName], [StoredProcedure])");
            if (totalCommonRows > 0)
            {
                sw.WriteLine(tabSpace1 + "SELECT '''  + @DatabaseName +  ''', ''TempDataCommon'',''up_fromTmptoProductionDataCommon''");

            }
            if (totalDailyRows > 0 && totalCommonRows > 0)
            {
                sw.WriteLine(tabSpace1 + "UNION ALL");
                sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempDataDaily'',''up_fromTmptoProductionDataDaily''");
            }
            else if (totalDailyRows > 0)
            {
                sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempDataDaily'',''up_fromTmptoProductionDataDaily''");
            }
            if (totalTrain1Rows > 0 && (totalDailyRows > 0 || totalCommonRows > 0))
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with ZW in it
                    if (worksheetNames[j].ToString().Contains("ZW"))
                    {
                        sw.WriteLine(tabSpace1 + "UNION ALL");
                        sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempData" + worksheetNames[j] + "'',''up_fromTmptoProductionData" + worksheetNames[j] + "''");

                    }
                }
            }
            else if (totalTrain1Rows > 0)
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with ZW in it
                    if (worksheetNames[j].ToString().Contains("ZW"))
                    {
                        sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempData" + worksheetNames[j] + "'',''up_fromTmptoProductionData" + worksheetNames[j] + "''");

                    }
                }
            }

            if (totalMit1Rows > 0 && (totalTrain1Rows > 0 || totalDailyRows > 0 || totalCommonRows > 0))
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MIT in it
                    if (worksheetNames[j].ToString().Contains("MIT"))
                    {
                        sw.WriteLine(tabSpace1 + "UNION ALL");
                        sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempData" + worksheetNames[j] + "'',''up_fromTmptoProductionData" + worksheetNames[j] + "''");

                    }
                }
            }
            else if (totalMit1Rows > 0)
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MIT in it
                    if (worksheetNames[j].ToString().Contains("MIT"))
                    {
                        sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempData" + worksheetNames[j] + "'',''up_fromTmptoProductionData" + worksheetNames[j] + "''");

                    }
                }
            }


            if ( totalMCRows > 0 && (totalMit1Rows > 0 || totalTrain1Rows > 0 || totalDailyRows > 0 || totalCommonRows > 0))
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MC in it
                    if (worksheetNames[j].ToString().Contains("MC"))
                    {
                        sw.WriteLine(tabSpace1 + "UNION ALL");
                        sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempData" + worksheetNames[j] + "'',''up_fromTmptoProductionData" + worksheetNames[j] + "''");

                    }
                }
            }
            else if (totalMCRows > 0)
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MC in it
                    if (worksheetNames[j].ToString().Contains("MC"))
                    {
                        sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempData" + worksheetNames[j] + "'',''up_fromTmptoProductionData" + worksheetNames[j] + "''");

                    }
                }
            }

            if (totalRCRows > 0 && (totalMCRows > 0 || totalMit1Rows > 0 || totalTrain1Rows > 0 || totalDailyRows > 0 || totalCommonRows > 0))
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with RC in it
                    if (worksheetNames[j].ToString().Contains("RC"))
                    {
                        sw.WriteLine(tabSpace1 + "UNION ALL");
                        sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempData" + worksheetNames[j] + "'',''up_fromTmptoProductionData" + worksheetNames[j] + "''");

                    }
                }
            }
            else if (totalRCRows > 0)
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with RC in it
                    if (worksheetNames[j].ToString().Contains("RC"))
                    {
                        sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''TempData" + worksheetNames[j] + "'',''up_fromTmptoProductionData" + worksheetNames[j] + "''");

                    }
                }
            }

            sw.WriteLine(tabSpace1 + "'");
            sw.Write(sw.NewLine);
            sw.WriteLine(tabSpace1 + "--PRINT @nSQL");
            sw.WriteLine(tabSpace1 + "EXECUTE (@nSQL)");
            sw.WriteLine(tabSpace1 + "-- Save any non-zero @@ERROR value.");
            sw.WriteLine(tabSpace1 + "SET @ErrorSave = @@ERROR");
            sw.WriteLine(tabSpace1 + "IF (@ErrorSave <> 0)");
            sw.WriteLine(tabSpace2 + "PRINT 'Error Number:' + STR(@ErrorSave)");
            sw.WriteLine(tabSpace1 + "ELSE");
            sw.WriteLine(tabSpace2 + "PRINT 'ZenoTracDatabasesAndTempTables table updated successfully.'");
            sw.Write(sw.NewLine);
            sw.WriteLine("END");


            sw.Write(sw.NewLine);

            sw.WriteLine("/*");
            sw.WriteLine("*Update ZenoTracSitesDTSNames Table if the row does not exist");
            sw.WriteLine("*/");
            sw.WriteLine("IF EXISTS (SELECT [SiteName] FROM [DataCleaner].[dbo].[ZenoTracSitesDTSNames]");
            sw.WriteLine("WHERE [SiteName] = @DatabaseName)");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "SET @DTSNamesRow   = 1");
            sw.WriteLine(tabSpace1 + "PRINT 'Cannot insert the row because the database name, ' + @DatabaseName + ' already exists in the ZenoTracSitesDTSNames table .'");
            sw.WriteLine("END");
            sw.Write(sw.NewLine);

            sw.WriteLine("IF @DTSNamesRow  = 0");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "SET @nSQL = '");
            sw.WriteLine(tabSpace1 + "INSERT INTO [DataCleaner].[dbo].[ZenoTracSitesDTSNames]([SiteName], [DTSName], [Status])");
            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''' + @DatabaseName + ''',''no''");
            sw.WriteLine(tabSpace1 + "'");
            sw.WriteLine(tabSpace1 + "--PRINT @nSQL");
            sw.WriteLine(tabSpace1 + "EXECUTE (@nSQL)");
            sw.WriteLine(tabSpace1 + "-- Save any non-zero @@ERROR value.");
            sw.WriteLine(tabSpace1 + "SET @ErrorSave = @@ERROR");
            sw.WriteLine(tabSpace1 + "IF (@ErrorSave <> 0)");
            sw.WriteLine(tabSpace2 + "PRINT 'Error Number:' + STR(@ErrorSave)");
            sw.WriteLine(tabSpace1 + "ELSE");
            sw.WriteLine(tabSpace2 + "PRINT 'ZenoTracSitesDTSNames table updated successfully.'");
            sw.WriteLine("END");
            sw.Write(sw.NewLine);


            sw.WriteLine("/*");
            sw.WriteLine("*Update tblDailyIncomingEmailReport Table if the row does not exist");
            sw.WriteLine("*/");
            //sw.WriteLine("/*");
            sw.WriteLine("IF EXISTS (SELECT [eSubject] FROM [DataCleaner].[dbo].[tblDailyIncomingEmailReport]");
            sw.WriteLine("WHERE [eSubject] = @DatabaseName)");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "SET @IncomingEmailRow = 1");
            sw.WriteLine(tabSpace1 + "PRINT 'Cannot insert the row because the database name, ' + @DatabaseName + ' already exists in the tblDailyIncomingEmailReport table .'");
            sw.WriteLine("END");
            sw.Write(sw.NewLine);

            sw.WriteLine("IF @IncomingEmailRow  = 0");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "SET @nSQL = '");
            sw.WriteLine(tabSpace1 + "INSERT INTO [DataCleaner].[dbo].[tblDailyIncomingEmailReport]([eFrom], [eSubject], [Email], [eCount], [SiteName], [Attachment], [ActivationDate], [SiteAssignedTo], [ProjectNumber], [CreateCIT])");
            if (siteAssigned == "Dave")
            {
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + '@zenon.com'',''' + @DatabaseName + ''', ''no'', 0, ''' + @SiteName + ' ' + @AONumber + ' (NEW - Not in QA yet)' + ''', ''NO'', CONVERT(CHAR(10), GETDATE(),101) , 1362, ''' + @AONumber + ''', ''NO''" + "--DATE CONVERTED TO MM/DD/YYYY");
                }
                else
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + '@zenon.com'',''' + @DatabaseName + ''', ''no'', 0, ''' + @SiteName + ' ' + @AONumber + ' (NEW - Not in QA yet)' + ''', ''NO'', CONVERT(CHAR(10), GETDATE(),101) , 1362, ''' + @SiteName + ''', ''NO''" + "--DATE CONVERTED TO MM/DD/YYYY");
                }
            }
            else if (siteAssigned == "Saima")
            {
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + '@zenon.com'',''' + @DatabaseName + ''', ''no'', 0, ''' + @SiteName + ' ' + @AONumber + ' (NEW - Not in QA yet)' + ''', ''NO'', CONVERT(CHAR(10), GETDATE(),101) , 2480, ''' + @AONumber + ''', ''NO''" + "--DATE CONVERTED TO MM/DD/YYYY");
                }
                else
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + '@zenon.com'',''' + @DatabaseName + ''', ''no'', 0, ''' + @SiteName + ' ' + @AONumber + ' (NEW - Not in QA yet)' + ''', ''NO'', CONVERT(CHAR(10), GETDATE(),101) , 2480, ''' + @SiteName + ''', ''NO''" + "--DATE CONVERTED TO MM/DD/YYYY");
                }
            }
            else if (siteAssigned == "Edison")
            {
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + '@zenon.com'',''' + @DatabaseName + ''', ''no'', 0, ''' + @SiteName + ' ' + @AONumber + ' (NEW - Not in QA yet)' + ''', ''NO'', CONVERT(CHAR(10), GETDATE(),101) , 10866, ''' + @AONumber + ''', ''NO''" + "--DATE CONVERTED TO MM/DD/YYYY");
                }
                else
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + '@zenon.com'',''' + @DatabaseName + ''', ''no'', 0, ''' + @SiteName + ' ' + @AONumber + ' (NEW - Not in QA yet)' + ''', ''NO'', CONVERT(CHAR(10), GETDATE(),101) , 10866, ''' + @SiteName + ''', ''NO''" + "--DATE CONVERTED TO MM/DD/YYYY");
                }
            }
            else if (siteAssigned == "Sandeep")
            {
                if (aoNumberExists == true)
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + '@zenon.com'',''' + @DatabaseName + ''', ''no'', 0, ''' + @SiteName + ' ' + @AONumber + ' (NEW - Not in QA yet)' + ''', ''NO'', CONVERT(CHAR(10), GETDATE(),101) , 1937, ''' + @AONumber + ''', ''NO''" + "--DATE CONVERTED TO MM/DD/YYYY");
                }
                else
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + '@zenon.com'',''' + @DatabaseName + ''', ''no'', 0, ''' + @SiteName + ' ' + @AONumber + ' (NEW - Not in QA yet)' + ''', ''NO'', CONVERT(CHAR(10), GETDATE(),101) , 1937, ''' + @SiteName + ''', ''NO''" + "--DATE CONVERTED TO MM/DD/YYYY");
                }
            }

            sw.WriteLine(tabSpace1 + "'");
            sw.WriteLine(tabSpace1 + "--PRINT @nSQL");
            sw.WriteLine(tabSpace1 + "EXECUTE (@nSQL)");
            sw.WriteLine(tabSpace1 + "-- Save any non-zero @@ERROR value.");
            sw.WriteLine(tabSpace1 + "SET @ErrorSave = @@ERROR");
            sw.WriteLine(tabSpace1 + "IF (@ErrorSave <> 0)");
            sw.WriteLine(tabSpace2 + "PRINT 'Error Number:' + STR(@ErrorSave)");
            sw.WriteLine(tabSpace1 + "ELSE");
            sw.WriteLine(tabSpace2 + "PRINT 'tblDailyIncomingEmailReport table updated successfully.'");
            sw.WriteLine("END");
            //sw.WriteLine("*/");
            sw.Write(sw.NewLine);


            sw.WriteLine("/*");
            sw.WriteLine("*Update tblDTSInformation Table if the rows do not exist");
            sw.WriteLine("*/");
            sw.WriteLine("IF EXISTS (SELECT [DBName] FROM [DataCleaner].[dbo].[tblDTSInformation]");
            sw.WriteLine("WHERE [DBName] = @DatabaseName)");
            sw.WriteLine("BEGIN");
            sw.WriteLine(tabSpace1 + "SET @DTSInformationRows   = 1");
            sw.WriteLine(tabSpace1 + "PRINT 'Cannot insert the rows because the database name of ' + @DatabaseName + ' already exists in the table tblDTSInformation.'");
            sw.WriteLine("END");
            sw.Write(sw.NewLine);
            sw.WriteLine("IF @DTSInformationRows  = 0");
            sw.WriteLine(tabSpace1 + "BEGIN");
            sw.WriteLine("--Common and Daily");
            sw.WriteLine(tabSpace1 + "SET @nSQL = '");
            sw.WriteLine(tabSpace1 + "INSERT INTO [DataCleaner].[dbo].[tblDTSInformation]([DBName], [DestLoc], [DestDTSLoc], [SiteName], [FileName], [File8_3Name], [DTSName], [SPName], [EmailSubject], [Status])");
            if (totalCommonRows > 0)
            {
                if (softwareType == "IFIX")
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\'  + @DatabaseName + 'DTS\\'', ''' + @DatabaseName + ''', ''Common.CSV'', ''CMN.CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionDataCommon'', ''' + @DatabaseName + ''', ''No''");
                }
                else if (softwareType == "OPC Trend")
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\'  + @DatabaseName + 'DTS\\'', ''' + @DatabaseName + ''', ''CMN'', ''CMN.CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionDataCommon'', ''' + @DatabaseName + ''', ''No''");
                }
            }
            if (totalDailyRows > 0 && totalCommonRows > 0)
            {
                sw.WriteLine(tabSpace1 + "UNION ALL");
                if (softwareType == "IFIX")
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''Daily.CSV'', ''DA.CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionDataDaily'',''' + @DatabaseName + ''', ''No''");
                }
                else if (softwareType == "OPC Trend")
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''Daily'', ''DA.CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionDataDaily'',''' + @DatabaseName + ''', ''No''");
                }
            }
            else if (totalDailyRows > 0)
            {
                if (softwareType == "IFIX")
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''Daily.CSV'', ''DA.CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionDataDaily'',''' + @DatabaseName + ''', ''No''");
                }
                else if (softwareType == "OPC Trend")
                {
                    sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''Daily'', ''DA.CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionDataDaily'',''' + @DatabaseName + ''', ''No''");

                }
            }

            if (totalTrain1Rows > 0 && (totalDailyRows > 0 || totalCommonRows > 0))
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with ZW in it
                    if (worksheetNames[j].ToString().Contains("ZW"))
                    {
                        sw.WriteLine(tabSpace1 + "UNION ALL");
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + ".CSV'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                        else if (softwareType == "OPC Trend")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + "'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                    }
                }
            }
            else if (totalTrain1Rows > 0)
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with ZW in it
                    if (worksheetNames[j].ToString().Contains("ZW"))
                    {
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + ".CSV'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                        else if (softwareType == "OPC Trend")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + "'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                    }
                }
            }

            if (totalMit1Rows > 0 && (totalTrain1Rows > 0 || totalDailyRows > 0 || totalCommonRows > 0))
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MIT in it
                    if (worksheetNames[j].ToString().Contains("MIT"))
                    {
                        sw.WriteLine(tabSpace1 + "UNION ALL");
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + ".CSV'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                        else if (softwareType == "OPC Trend")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + "'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                    }
                }
            }
            else if (totalMit1Rows > 0)
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MIT in it
                    if (worksheetNames[j].ToString().Contains("MIT"))
                    {
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + ".CSV'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                        else if (softwareType == "OPC Trend")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + "'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                    }
                }
            }
            if (totalMCRows > 0 && (totalMit1Rows > 0 || totalTrain1Rows > 0 || totalDailyRows > 0 || totalCommonRows > 0))
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MC in it
                    if (worksheetNames[j].ToString().Contains("MC"))
                    {
                        sw.WriteLine(tabSpace1 + "UNION ALL");
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + ".CSV'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                        else if (softwareType == "OPC Trend")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + "'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                    }
                }
            }
            else if (totalMCRows > 0)
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with MC in it
                    if (worksheetNames[j].ToString().Contains("MC"))
                    {
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + ".CSV'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                        else if (softwareType == "OPC Trend")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + "'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                    }
                }
            }

            if (totalRCRows > 0 && (totalMCRows > 0 || totalMit1Rows > 0 || totalTrain1Rows > 0 || totalDailyRows > 0 || totalCommonRows > 0))
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with RC in it
                    if (worksheetNames[j].ToString().Contains("RC"))
                    {
                        sw.WriteLine(tabSpace1 + "UNION ALL");
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + ".CSV'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                        else if (softwareType == "OPC Trend")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + "'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                    }
                }
            }
            else if (totalRCRows > 0)
            {
                for (int j = 0; j < worksheetNames.Length; j++)
                {
                    //get worksheet name with RC in it
                    if (worksheetNames[j].ToString().Contains("RC"))
                    {
                        if (softwareType == "IFIX")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + ".CSV'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                        else if (softwareType == "OPC Trend")
                        {
                            sw.WriteLine(tabSpace1 + "SELECT ''' + @DatabaseName + ''',''C:\\Copy of Emails\\' + @DatabaseName + '\\'', ''C:\\Copy of Emails\\' + @DatabaseName + 'DTS\\'',''' + @DatabaseName + ''', ''" + worksheetNames[j] + "'', ''" + worksheetNames[j] + ".CSV'',''' + @DatabaseName + ''', ''up_fromTmptoProductionData" + worksheetNames[j] + "'',''' + @DatabaseName + ''', ''No''");
                        }
                    }
                }
            }


            sw.WriteLine(tabSpace1 + "'");
            sw.WriteLine(tabSpace1 + "--PRINT @nSQL");
            sw.WriteLine(tabSpace1 + "EXECUTE (@nSQL)");
            sw.WriteLine(tabSpace1 + "-- Save any non-zero @@ERROR value.");
            sw.WriteLine(tabSpace1 + "SET @ErrorSave = @@ERROR");
            sw.WriteLine(tabSpace1 + "IF (@ErrorSave <> 0)");
            sw.WriteLine(tabSpace2 + "PRINT 'Error Number:' + STR(@ErrorSave)");
            sw.WriteLine(tabSpace1 + "ELSE");
            sw.WriteLine(tabSpace2 + "PRINT 'tblDTSInformation table updated successfully.'");
            sw.Write(sw.NewLine);
            sw.WriteLine("END");
            sw.Write(sw.NewLine);
            sw.WriteLine("GO");
            sw.Write(sw.NewLine);
            sw.Close();
        }

        /// <summary>
        /// Returns the first worksheet name that appears in the Excel file.  Searches the sheets from left to right.
        /// </summary>
        /// <param name="searchString">The string value to search for.</param>
        /// <returns></returns>
        private string SearchFirstNameOfWorkSheet(string searchString)
        {
            for (int j = 0; j < worksheetNames.Length; j++)
            {

                if (worksheetNames[j].ToString().Contains(searchString))
                {
                    string firstName = worksheetNames[j].ToString();
                    return firstName;

                }
            }
            return "";
        }

        /// <summary>
        /// Count the number of worksheets with the given search string.  Searches the sheets from left to right.
        /// </summary>
        /// <param name="searchString">The string value to search for.</param>
        /// <returns></returns>
        private int NumberofWorksheetsWithName(string searchString)
        {
            int countSheets = 0;

            for (int j = 0; j < worksheetNames.Length; j++)
            {
                if (worksheetNames[j].ToString().Contains(searchString))
                {
                    countSheets = countSheets + 1;
                }
            }
            return countSheets;
        }
        /// <summary>
        /// Returns the location for the last tag in the tab that you search for. If there are muliple tabs for ie ZW then goes from left to right.
        /// </summary>
        /// <param name="searchString">The string value to search for.</param>
        /// <returns></returns>
        private int [] GetNumberofTagsForWorksheet(string worksheetName)
        {

                  //declare variables
            DataTable databaseTagsTable = m_excelData.Copy();

            //delete rows that have digital tags
            DataRow[] getDigitalRows = null;

            //find the digital rows that have "ready"
            getDigitalRows = FindRowsInDataTable(databaseTagsTable, "ready");


            foreach (DataRow dr in getDigitalRows)
            {
                //removed the data ready rows since these are digital tags
                databaseTagsTable.Rows.Remove(dr);

            }

            int []totalTags;
            int numberofTags = 0;
            int startPoint = 0;
            int arraySize = NumberofWorksheetsWithName(worksheetName);

            totalTags = new int[arraySize];
            for (int j = 0; j < worksheetNames.Length; j++)
            {

                if (worksheetNames[j].ToString().Contains(worksheetName))
                {

                    for (int i = 0; i <= databaseTagsTable.Rows.Count - 1; i++)
                    {
                        //remove sheetname from the tagname when printing the tag
                        startPoint = databaseTagsTable.Rows[i]["Tag Name"].ToString().IndexOf(".");
                        //add the first ZW train first
                        if (worksheetNames[j].ToString() == databaseTagsTable.Rows[i]["Tag Name"].ToString().Substring(0, startPoint))
                        {
                            numberofTags = numberofTags + 1;
                        }
                    }
                    if (arraySize >= 1)
                    {
                        //the max length of the array holds the last index for the first tab
                        totalTags[arraySize - 1] = numberofTags;
                        arraySize = arraySize - 1;
                    }
                }
            }            
            return totalTags;
                  
        }


        /// <summary>
        /// Gets start, end, and total rows of given data table.
        /// </summary>
        /// <param name="databaseTable"></param>
        /// <param name="startCommon"></param>
        /// <param name="endCommon"></param>
        /// <param name="totalCommonTags"></param>
        /// <param name="startDaily"></param>
        /// <param name="endDaily"></param>
        /// <param name="totalDailyTags"></param>
        /// <param name="startTrain"></param>
        /// <param name="endTrain"></param>
        /// <param name="totalTrainTags"></param>
        /// <param name="startTrain1"></param>
        /// <param name="endTrain1"></param>
        /// <param name="totalTrain1Tags"></param>
        /// <param name="startMit1"></param>
        /// <param name="endMit1"></param>
        /// <param name="totalMit1Tags"></param>
        private void GetStartEndIndexesForTagsInDataTable(ref DataTable databaseTable, ref int startCommon, ref int endCommon, ref int totalCommonTags, ref int startDaily, ref int endDaily, ref int totalDailyTags, ref int startTrain, ref int endTrain, ref int totalTrainTags, ref int startTrain1, ref int endTrain1, ref int totalTrain1Tags, ref int startMit, ref int endMit, ref int totalMitTags, ref int startMit1, ref int endMit1, ref int totalMit1Tags, ref int startMC, ref int endMC, ref int totalMCTags, ref int startMC1, ref int endMC1, ref int totalMC1Tags, ref int startRC, ref int endRC, ref int totalRCTags, ref int startRC1, ref int endRC1, ref int totalRC1Tags)
        {

            //delete rows that have digital tags
            DataRow[] getDigitalRows = null;

            //find the digital rows that have "ready"
            getDigitalRows = FindRowsInDataTable(databaseTable, "ready");


            foreach (DataRow dr in getDigitalRows)
            {
                //removed the data ready rows since these are digital tags
                databaseTable.Rows.Remove(dr);

            }

            //sort data table alphabetically
            //databaseTable.DefaultView.Sort = string.Format("{0}", "Tag Name", "ASC");
           // databaseTable = databaseTable.DefaultView.Table;

            int mitNameLength = SearchFirstNameOfWorkSheet("MIT").Length;
            int trainNameLength = SearchFirstNameOfWorkSheet("ZW").Length;
            int commonNameLength = SearchFirstNameOfWorkSheet("Common").Length;
            int dailyNameLength = SearchFirstNameOfWorkSheet("Daily").Length;
            int mcNameLength = SearchFirstNameOfWorkSheet("MC").Length;
            int rcNameLength = SearchFirstNameOfWorkSheet("RC").Length;

            foreach (DataRow dr in databaseTable.Rows)
            {
                if (trainNameLength > 0)
                {
                    //get train start indexes
                    if ((dr["Tag Name"].ToString().Substring(0, 2) == "ZW") || (dr["Tag Name"].ToString().Substring(0, 2) == "zw"))
                    {
                        endTrain = databaseTable.Rows.IndexOf(dr);
                        totalTrainTags = totalTrainTags + 1;

                    }
                }

                if (mitNameLength > 0)
                {
                    //get train start indexes
                    if ((dr["Tag Name"].ToString().Substring(0, 3) == "MIT") || (dr["Tag Name"].ToString().Substring(0, 3) == "mit"))
                    {
                        endMit = databaseTable.Rows.IndexOf(dr);
                        totalMitTags = totalMitTags + 1;

                    }
                }
                //for future when you have multiple common and daily tabs
                if (commonNameLength > 0)
                {
                    //get common and daily start indexes
                    if ((dr["Tag Name"].ToString().Substring(0, 6) == "Common") || (dr["Tag Name"].ToString().Substring(0, 6) == "common") || (dr["Tag Name"].ToString().Substring(0, 3) == "CMN"))
                    {
                        endCommon = databaseTable.Rows.IndexOf(dr);
                        totalCommonTags = totalCommonTags + 1;
                    }
                }

                if (dailyNameLength > 0)
                {
                    if ((dr["Tag Name"].ToString().Substring(0, 5) == "Daily") || (dr["Tag Name"].ToString().Substring(0, 5) == "daily") || (dr["Tag Name"].ToString().Substring(0, 2) == "DA"))
                    {
                        endDaily = databaseTable.Rows.IndexOf(dr);
                        totalDailyTags = totalDailyTags + 1;

                    }
                }

                if (mcNameLength > 0)
                {
                    //get train start indexes
                    if ((dr["Tag Name"].ToString().Substring(0, 2) == "MC") || (dr["Tag Name"].ToString().Substring(0, 2) == "mc"))
                    {
                        endMC = databaseTable.Rows.IndexOf(dr);
                        totalMCTags = totalMCTags + 1;

                    }
                }

                if (rcNameLength > 0)
                {
                    //get train start indexes
                    if ((dr["Tag Name"].ToString().Substring(0, 2) == "RC") || (dr["Tag Name"].ToString().Substring(0, 2) == "rc"))
                    {
                        endRC = databaseTable.Rows.IndexOf(dr);
                        totalRCTags = totalRCTags + 1;

                    }
                }

          

                //get MIT index for train 1
                if (mitNameLength > 0)
                {
                    if ((dr["Tag Name"].ToString().Substring(0, mitNameLength) == SearchFirstNameOfWorkSheet("MIT")) || (dr["Tag Name"].ToString().Substring(0, mitNameLength) == SearchFirstNameOfWorkSheet("mit")))
                    {
                        endMit1 = databaseTable.Rows.IndexOf(dr);
                        totalMit1Tags = totalMit1Tags + 1;
                    }
                }
                if (trainNameLength > 0)
                {
                    if ((dr["Tag Name"].ToString().Substring(0, trainNameLength) == SearchFirstNameOfWorkSheet("ZW")) || (dr["Tag Name"].ToString().Substring(0, trainNameLength) == SearchFirstNameOfWorkSheet("zw")))
                    {
                        endTrain1 = databaseTable.Rows.IndexOf(dr);
                        totalTrain1Tags = totalTrain1Tags + 1;
                    }
                }

                if (mcNameLength > 0)
                {
                    if ((dr["Tag Name"].ToString().Substring(0, mcNameLength) == SearchFirstNameOfWorkSheet("MC")) || (dr["Tag Name"].ToString().Substring(0, mcNameLength) == SearchFirstNameOfWorkSheet("mc")))
                    {
                        endMC1 = databaseTable.Rows.IndexOf(dr);
                        totalMC1Tags = totalMC1Tags + 1;
                    }
                }

                if (rcNameLength > 0)
                {
                    if ((dr["Tag Name"].ToString().Substring(0, rcNameLength) == SearchFirstNameOfWorkSheet("RC")) || (dr["Tag Name"].ToString().Substring(0, rcNameLength) == SearchFirstNameOfWorkSheet("rc")))
                    {
                        endRC1 = databaseTable.Rows.IndexOf(dr);
                        totalRC1Tags = totalRC1Tags + 1;
                    }
                }

            }

            //start indexes
            startTrain = endTrain - (totalTrainTags - 1);
            startMit= endMit- (totalMitTags - 1);
            startCommon = endCommon - (totalCommonTags - 1);
            startDaily = endDaily - (totalDailyTags - 1);
            startMit1 = endMit1 - (totalMit1Tags - 1);
            startTrain1 = endTrain1 - (totalTrain1Tags - 1);
            startMC = endMC - (totalMCTags - 1);
            startMC1 = endMC1 - (totalMC1Tags - 1);
            startRC = endRC - (totalRCTags - 1);
            startRC1 = endRC1 - (totalRC1Tags - 1);            

        }
    }
}

       

 
