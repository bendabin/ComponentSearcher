#define testReadTable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Data.OleDb;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Windows.Forms;

namespace ComponentSearcher
{       
    class MyExcel
    {        
        //enum 
        enum ComboOp { START, CONTAINS, ALL };

        enum WriteModeType { QUANTITY, EDIT, NEW_ENTRIES };

        //importing DLL's into class 

        //End task
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool EndTask(IntPtr hWnd, bool fShutDown, bool fForce);
               
        // The FindWindow function retrieves a handle to the top-level 
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Find window by Caption only. Note you must pass IntPtr.Zero as the first parameter.

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        //Get process window thread process ID API
        [DllImport("user32.dll", SetLastError=true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        //Set last error to zero, to zero out error state of application
        [DllImport("user32.dll", SetLastError = true)]
        static extern void SetLastErrorEx(uint dwErrCode, uint dwType);


        /*          Declare some properties here to pass to other classes 
         * 
         * 
         */

        //declare class objects
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static Excel.Range myRange;
        ProcessExcelVars hWndVal = new ProcessExcelVars();

        CultureInfo ci;

        DataTable ExcelSpreadVals = new DataTable();

        //declare global static variables here
        private static int lastRow = 0;
        private static int lastCol = 0;

        static int hWnd = 0;

        static string ExcelPath = null; //This stores the excel path value for finding spreadsheet's
                                        //location  
        
       //Initialise Excel spreadsheet and save directing into local variable
        public void ExcelInit(string path)
        {
            //Initialise Excel spreadsheet by loading values into a
            try
            {
                if (path != null && path != string.Empty)
                {
                    //Check to see if spreadsheet open and closes successfully 

                    MyApp = new Excel.Application();

                    //store value main excel window value in ProcessExcelVars class
                    hWnd = MyApp.Hwnd;
                    PartsListClass.Hwind = hWnd;

                    MyApp.Visible = false;
                    MyBook = MyApp.Workbooks.Open(path);
                    MySheet = (Excel.Worksheet)MyBook.Sheets[1];            // Explict cast is not required here                   
                    lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    lastCol = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    //Quit Excel application and release all references 
                    //and allow Garbage collector to handle the cleaning process
                    MyBook.Close(false, Type.Missing, Type.Missing);
                    MyApp.Quit();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.FinalReleaseComObject(MySheet);
                    Marshal.FinalReleaseComObject(MyBook);
                    Marshal.FinalReleaseComObject(MyApp);
                    MyApp = null;
                    MyBook = null;
                    MySheet = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    //Kill excel process that was ctreated in memory 
                    KillProcessByMainWindows(hWnd);

                    ExcelPath = path;   //current excel path to be stored in variable                   
                }
            }
            catch (Exception ex)
            {
                ExcelPath = null;
                //MessageBox.Show(ex.ToString(), "Error");
                //catch exception message and store 
                throw new ExcelExceptionMessage(ex.Message); //throw exception string to be caught in Form1 class        
            }
        }

        //public DataTable ReadExcelValues(string locationStr, string keyWord, int searchOption, int wordSearchFilter)
        public DataTable ReadExcelValues(string locationStr, string keyWord, int searchOption, int wordSearchFilter, int SelectSearchMode)
        {
            //array for storing the filtered results datatable! 
            int newRowIndex = 0;

            //declaring filtering flags
            Boolean KeywordFilter = false;
            Boolean CabinetLocFilter = false;
            Boolean contains = false;

            //declare new table for storing searched values 
            DataTable excelSearchVars = new DataTable();

            excelSearchVars.Clear();      //clear excel table which stores search results            

            //Add columns to the data and display in spreadsheet
            excelSearchVars.Columns.Add("Part");
            excelSearchVars.Columns.Add("Pack");
            excelSearchVars.Columns.Add("Cabinet");
            excelSearchVars.Columns.Add("Row");
            excelSearchVars.Columns.Add("Drawer");
            excelSearchVars.Columns.Add("Section");
            excelSearchVars.Columns.Add("Part Info");
            excelSearchVars.Columns.Add("Supplier");
            excelSearchVars.Columns.Add("Alt Parts");
            excelSearchVars.Columns.Add("SMD Marking");
            excelSearchVars.Columns.Add("Quantity");
            excelSearchVars.Columns.Add("Index");
            
                           
            
            //open and create new excel data table based on search results
            if (ExcelPath != null && ExcelPath != string.Empty)
            {
                MyApp = new Excel.Application();

                //store value main excel window value in ProcessExcelVars class
                hWnd = MyApp.Hwnd;
                PartsListClass.Hwind = hWnd;

                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(ExcelPath);
                MySheet = (Excel.Worksheet)MyBook.Sheets[1];            // Explict cast is not required here                   
                lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                lastCol = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                //Create a new write range for the spreadsheet entries to be added too
                int StartCellRow = 1;
                int LasttCellRow = lastRow;

                //Get Excel value range of cells 
                Excel.Range CellStart = (Excel.Range)MySheet.Cells[StartCellRow, 1];
                Excel.Range CellEnd = (Excel.Range)MySheet.Cells[LasttCellRow, lastCol];
                Excel.Range MySheetRng = MySheet.get_Range(CellStart, CellEnd);

                //Store values inside object array
                object[,] data = MySheetRng.Cells.Value2;

                //close Excel spreadsheet
                MyBook.Close(false, Type.Missing, Type.Missing);

                MyApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.FinalReleaseComObject(MySheet);
                Marshal.FinalReleaseComObject(MyBook);
                Marshal.FinalReleaseComObject(MyApp);
                MyApp = null;
                MyBook = null;
                MySheet = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //Kill excel process that was ctreated in memory 
                KillProcessByMainWindows(hWnd);

                //Add an extra column for the index 
                int colSize = lastCol + 1;

                List<string> deConstructerStr = new List<string>();
                //split comma in string
                Char delimiter = ',';
                String[] cabinetCompareStr = locationStr.Split(delimiter);

                for (int i = 0; i < lastRow; i++)
                {
                    //see what mode is selected using full filters or none. 
                    if (SelectSearchMode == 0)  //will search spreadsheet based on keywords and location information
                    {

                        if (keyWord == "*")
                            KeywordFilter = true;
                        else
                        {
                            //extract either part number or part information column and convert to string
                            string rowString = (data[i + 1, 1 + searchOption] ?? String.Empty).ToString();

                            //Check what option is selected for word seach filter 
                            if (wordSearchFilter == 0)  //search for strings starting with keyword 
                                contains = rowString.StartsWith(keyWord, true, ci);
                            else                        //search for strings containing keyword
                                contains = rowString.IndexOf(keyWord, StringComparison.OrdinalIgnoreCase) >= 0;

                            //check if keyword in box matches part number column on spreadsheet
                            if (contains && !(String.IsNullOrWhiteSpace(keyWord)))
                                KeywordFilter = true;
                            else
                                KeywordFilter = false;
                        }

                        if (cabinetCompareStr[0].ToString() != "*")
                        {
                            //construct an array based on number of elements selected in the form option boxes
                            for (int j = 0; j < cabinetCompareStr.Length; j++)
                            {
                                //deConstructerStr.Add(ExcelSpreadVals.Rows[i][2 + j].ToString());
                                deConstructerStr.Add((data[i + 1, 3 + j] ?? String.Empty).ToString());
                            }
                            //check if guven
                            CabinetLocFilter = deConstructerStr.SequenceEqual(cabinetCompareStr);

                            deConstructerStr.Clear(); //reset and clear array
                        }
                        else
                            CabinetLocFilter = true;
                    }

                    //If option 1 selected search datasheet based on location information only
                    if(SelectSearchMode == 1)
                    {
                        KeywordFilter = true;
                        //construct an array based on number of elements selected in the form option boxes
                        for (int a = 0; a < cabinetCompareStr.Length; a++)
                        {
                            //deConstructerStr.Add(ExcelSpreadVals.Rows[i][2 + j].ToString());
                            deConstructerStr.Add((data[i + 1, 3 + a] ?? String.Empty).ToString());
                        }
                        //check if guven
                        CabinetLocFilter = deConstructerStr.SequenceEqual(cabinetCompareStr);

                        deConstructerStr.Clear(); //reset and clear array

                    }

                    if (CabinetLocFilter == true && KeywordFilter == true)
                    {
                        excelSearchVars.Rows.Add();
                        for (int a = 0; a < excelSearchVars.Columns.Count; a++)
                        {
                            if (a < lastCol)
                                excelSearchVars.Rows[newRowIndex][a] = data[i + 1, a + 1];
                            else
                                excelSearchVars.Rows[newRowIndex][a] = 1 + i;   //record the index
                        }
                        newRowIndex++;  //increment index of new table and 
                    }
                }            
        }             
           return excelSearchVars; //return table 


        }


        //This is without the edit features and stuff
        public Boolean WriteExcelValues2(DataGridView ExcelDataVals)
        {
            try
            {
                if (ExcelPath != null && ExcelPath != string.Empty)
                {
                    MyApp = new Excel.Application();

                  //  string index = ExcelDataVals.Rows[0].Cells[11].Value.ToString();

                    //store value main excel window value in ProcessExcelVars class
                    hWnd = MyApp.Hwnd;
                    PartsListClass.Hwind = hWnd;

                    MyApp.Visible = false;
                    MyBook = MyApp.Workbooks.Open(ExcelPath);
                    MySheet = (Excel.Worksheet)MyBook.Sheets[1];            // Explict cast is not required here                   
                    lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    lastCol = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    //Create a new write range for the spreadsheet entries to be added too
                    int StartCellRow = 1;
                    int LasttCellRow = lastRow;

                    //Get Excel value range of cells 
                    Excel.Range CellStart = (Excel.Range)MySheet.Cells[StartCellRow, 1];
                    Excel.Range CellEnd = (Excel.Range)MySheet.Cells[LasttCellRow, lastCol];
                    Excel.Range MySheetRng = MySheet.get_Range(CellStart, CellEnd);

                    int NextCellRow = LasttCellRow + 1;

                    //Write edited value to excel spreadsheet one column at a time!
                    for (int i = 0; i < ExcelDataVals.Rows.Count; i++)
                    {                        
                        for (int j = 0; j < ExcelDataVals.Columns.Count - 1; j++)
                        {
                            MySheet.Cells[NextCellRow, 1 + j] = ExcelDataVals.Rows[i].Cells[j].Value;    //write value to spreadsheet
                        }
                        ++NextCellRow;
                    }

                    MyBook.Save();              //save quantity amount to excel spreadsheet

                    //Quit Excel application and release all references 
                    //and allow Garbage collector to handle the cleaning process
                    MyBook.Close(false, Type.Missing, Type.Missing);

                    MyApp.Quit();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.FinalReleaseComObject(MySheet);
                    Marshal.FinalReleaseComObject(MyBook);
                    Marshal.FinalReleaseComObject(MyApp);
                    MyApp = null;
                    MyBook = null;
                    MySheet = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    //Kill excel process that was ctreated in memory 
                    KillProcessByMainWindows(hWnd);

                }
            }
            catch
            { }
            return true;
        }


        public Boolean WriteFunctionTest(DataGridView ExcelDataVals, int ModeSel)
        {
            List<string> cabinetCompareStr = new List<string>();
            List<string> cabinetCompareStr2 = new List<string>();

            int RowIndex = 0;
            try
            {
                if (ExcelPath != null && ExcelPath != string.Empty)
                {
                    MyApp = new Excel.Application();

                    //store value main excel window value in ProcessExcelVars class
                    hWnd = MyApp.Hwnd;
                    PartsListClass.Hwind = hWnd;

                    MyApp.Visible = false;
                    MyBook = MyApp.Workbooks.Open(ExcelPath);
                    MySheet = (Excel.Worksheet)MyBook.Sheets[1];            // Explict cast is not required here                   
                    lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    lastCol = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    //Create a new write range for the spreadsheet entries to be added too
                    int StartCellRow = 1;
                    int LasttCellRow = lastRow;

                    //Get Excel value range of cells 
                    Excel.Range CellStart = (Excel.Range)MySheet.Cells[StartCellRow, 1];
                    Excel.Range CellEnd = (Excel.Range)MySheet.Cells[LasttCellRow, lastCol];
                    Excel.Range MySheetRng = MySheet.get_Range(CellStart, CellEnd);

                    //Store values inside object array
                    object[,] data = MySheetRng.Cells.Value2;


                    switch (ModeSel)
                    {
                        case 0:         //Add New entries to spreadsheet
                            RowIndex = LasttCellRow + 1;

                            //Write edited value to excel spreadsheet one column at a time!
                            for (int i = 0; i < ExcelDataVals.Rows.Count; i++)
                            {
                                for (int j = 0; j < ExcelDataVals.Columns.Count - 1; j++)
                                {
                                    MySheet.Cells[RowIndex, 1 + j] = ExcelDataVals.Rows[i].Cells[j].Value;    //write value to spreadsheet
                                }
                                ++RowIndex;
                            }
                            break;

                        case 1:     //Edit Existing entries to spreadsheet
                            //retrieve index adress of the modified or edited row
                            RowIndex = Convert.ToInt32(ExcelDataVals.CurrentRow.Cells[11].Value.ToString());
                            //Write edited value to excel spreadsheet one column at a time!
                            for (int i = 0; i < ExcelDataVals.CurrentRow.Cells.Count - 1; i++)
                                MySheet.Cells[RowIndex, 1 + i] = ExcelDataVals.CurrentRow.Cells[i].Value;    //write value to spreadsheet
                            break;

                        case 3:
                            //This will store values to which the table will be sorted
                            List<string> arrayElementsX = new List<string>();
                            List<string> currentRowString = new List<string>();
                            List<string> compareRowString = new List<string>();

                            List<string> newElements = new List<string>();
                            List<string> tempElements = new List<string>();

                            String[] xIndexExtract;
                            char[] delimiter = { ',' };

                            ///int nextRowIndex = 1;
                            Boolean indexOccurance = false;
                            string indexCount = null;


                            //MySheetRng.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                            //MySheetRng.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                            //MySheetRng.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                            //MySheetRng.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                            //MySheetRng.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                            //MySheetRng.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone

                            //MySheetRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


                            //get index of item from spreadsheet.
                            RowIndex = Convert.ToInt32(ExcelDataVals.CurrentRow.Cells[11].Value.ToString());

                            //delete from excel spreadsheet
                            MySheetRng = MySheetRng.Cells.Rows[RowIndex];
                            MySheetRng.Delete();
                            MyBook.Save();

                            //delete from datagridview 
                            ExcelDataVals.Rows.RemoveAt(ExcelDataVals.CurrentRow.Index);

                            //update column and row information from excel spreadsheet
                            //Get update value of rows
                            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                            lastCol = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                            //Get Excel value range of cells 
                            CellStart = (Excel.Range)MySheet.Cells[StartCellRow, 1];
                            CellEnd = (Excel.Range)MySheet.Cells[LasttCellRow, lastCol];
                            MySheetRng = MySheet.get_Range(CellStart, CellEnd);

                            //Reload and find new adress index's 
                            for (int c = 0; c < ExcelDataVals.Rows.Count; c++)
                            {
                                for (int i = 0; i < lastRow; i++)   //delete selected entries
                                {
                                    for (int j = 0; j < lastCol; j++)
                                    {
                                         currentRowString.Add((data[i + 1, 1 + j] ?? String.Empty).ToString());
                                         compareRowString.Add((ExcelDataVals.Rows[c].Cells[j].Value ?? String.Empty).ToString());
                                    }
                                    indexOccurance = currentRowString.SequenceEqual(compareRowString);

                                    indexCount = (i + 1).ToString();

                                    if (indexOccurance)
                                    {
                                        arrayElementsX.Add(c.ToString() + "," + indexCount);
                                    }
                                    currentRowString.Clear();
                                    compareRowString.Clear();
                                }

                                for (int y = 0; y < arrayElementsX.Count; y++)
                                {
                                    //split the assorting array into pieces
                                    xIndexExtract = arrayElementsX[y].Split(delimiter);

                                    if (xIndexExtract[0] == c.ToString())
                                    {
                                        tempElements.Add(xIndexExtract[1]); //add last index to temp array
                                    }
                                }

                                //check how many elements contain in temp elements 
                                if (tempElements.Count == 1)
                                {
                                    //add contents to new Element array
                                    newElements.Add(tempElements[0]);
                                    tempElements.Clear();
                                }

                                //more filtering through is required
                                if (tempElements.Count > 1)
                                {
                                    int index = 0;
                                    Boolean flagOccurance = false;

                                    do
                                    {
                                        if (flagOccurance == true)
                                            flagOccurance = false;  //switch flag off 
                                                                    //loop through newElements array and check that selected value isn't present inside

                                        for (int i = 0; i < newElements.Count; i++)
                                        {
                                            //check previous index value in newElements array                                      
                                            if (newElements[i] == tempElements[index])
                                            {
                                                index++;    //increment index
                                                flagOccurance = true;
                                                break; //break from for loop and increment
                                            }

                                        }

                                    } while (flagOccurance);

                                    //add temp element to array 
                                    newElements.Add(tempElements[index]);

                                }
                                tempElements.Clear();

                                //write new Index to following function
                                ExcelDataVals.Rows[c].Cells[11].Value = newElements[c];
                            }                      

                          break;
                    }

                    MyBook.Save();              //save quantity amount to excel spreadsheet

                    //Quit Excel application and release all references 
                    //and allow Garbage collector to handle the cleaning process
                    MyBook.Close(false, Type.Missing, Type.Missing);

                    MyApp.Quit();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.FinalReleaseComObject(MySheet);
                    Marshal.FinalReleaseComObject(MyBook);
                    Marshal.FinalReleaseComObject(MyApp);
                    MyApp = null;
                    MyBook = null;
                    MySheet = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    //Kill excel process that was ctreated in memory 
                    KillProcessByMainWindows(hWnd);

                }
            }
            catch (Exception ex)
            {
                //Quit Excel application and release all references 
                //and allow Garbage collector to handle the cleaning process
                MyBook.Close(false, Type.Missing, Type.Missing);

                MyApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.FinalReleaseComObject(MySheet);
                Marshal.FinalReleaseComObject(MyBook);
                Marshal.FinalReleaseComObject(MyApp);
                MyApp = null;
                MyBook = null;
                MySheet = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //Kill excel process that was ctreated in memory 
                KillProcessByMainWindows(hWnd);
              
            }       
        
            return true;
        }

        private int spreadSheetValsDetect(DataGridView ExcelDataVals)
        {
            List<string> deConstructerStr = new List<string>();
            List<string> deConstructerStr2 = new List<string>();
            //ExcelDataVals.Rows.RemoveAt(11);    //remove index column
            Boolean OccuranceFilter = false;
            int index = 0;


            MyApp = new Excel.Application();

            //store value main excel window value in ProcessExcelVars class
            hWnd = MyApp.Hwnd;
            PartsListClass.Hwind = hWnd;

            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(ExcelPath);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];            // Explict cast is not required here                   
            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            lastCol = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

            //Create a new write range for the spreadsheet entries to be added too
            int StartCellRow = 1;
            int LasttCellRow = lastRow;

            //Get Excel value range of cells 
            Excel.Range CellStart = (Excel.Range)MySheet.Cells[StartCellRow, 1];
            Excel.Range CellEnd = (Excel.Range)MySheet.Cells[LasttCellRow, lastCol];
            Excel.Range MySheetRng = MySheet.get_Range(CellStart, CellEnd);

            //Store values inside object array
            object[,] data = MySheetRng.Cells.Value2;

            OccuranceFilter = deConstructerStr.SequenceEqual(deConstructerStr2);

            for (int i = 0; i < lastRow; i++)
            {
                for (int j = 0; j < lastCol; j++)
                {
                    deConstructerStr.Add((data[i + 1, 1 + j] ?? String.Empty).ToString());
                    deConstructerStr2.Add((ExcelDataVals.CurrentRow.Cells[j].Value ?? String.Empty).ToString());
                }
            }

            //Quit Excel application and release all references 
            //and allow Garbage collector to handle the cleaning process
            MyBook.Close(false, Type.Missing, Type.Missing);

            MyApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.FinalReleaseComObject(MySheet);
            Marshal.FinalReleaseComObject(MyBook);
            Marshal.FinalReleaseComObject(MyApp);
            MyApp = null;
            MyBook = null;
            MySheet = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //Kill excel process that was ctreated in memory 
            KillProcessByMainWindows(hWnd);

            return 0;

        }


        
        public Boolean WriteExcelValues(DataGridViewRow ExcelDataVals)
        {
            try
            {
                //retrieve index adress of the modified or edited row
                int index = Convert.ToInt32(ExcelDataVals.Cells[11].Value.ToString());

                //Open Excel app for writing new data to 
                MyApp = new Excel.Application();
                hWnd = MyApp.Hwnd;
                PartsListClass.Hwind = hWnd;

                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(ExcelPath);
                MySheet = (Excel.Worksheet)MyBook.Sheets[1];            // Explict cast is not required here   

                //Write edited value to excel spreadsheet one column at a time!
                for (int i = 0; i < ExcelDataVals.Cells.Count - 1; i++)
                    MySheet.Cells[index, 1 + i] = ExcelDataVals.Cells[i].Value;    //write value to spreadsheet

                MyBook.Save();              //save quantity amount to excel spreadsheet

                //Quit Excel application and release all references 
                //and allow Garbage collector to handle the cleaning process
                MyBook.Close(false, Type.Missing, Type.Missing);

                MyApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.FinalReleaseComObject(MySheet);
                Marshal.FinalReleaseComObject(MyBook);
                Marshal.FinalReleaseComObject(MyApp);
                MyApp = null;
                MyBook = null;
                MySheet = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //Kill excel process that was ctreated in memory 
                KillProcessByMainWindows(hWnd);
            }
            catch
            { }
            return true;
        }
        














        /*      

    //This function will organise values and store them into a binding list
    private void StorePartsList(object [,]dataVals, int row, int col)
    {
        if(ExcelSpreadVals.Rows.Count > 0)
        {
            ExcelSpreadVals.Clear(); //clear data if array is full
        }
        //Add an extra column for the index 
        int colSize = col + 1;

        //Place values inside dataTable 
        for (int Colindex = 0; Colindex < colSize; Colindex++)
            ExcelSpreadVals.Columns.Add();
        for (int i = 0; i < row; i++)
        {
            ExcelSpreadVals.Rows.Add();
            for (int j = 0; j < colSize; j++)
            {
                if (j < col)
                    ExcelSpreadVals.Rows[i][j] = dataVals[i + 1, j + 1];
                else
                    ExcelSpreadVals.Rows[i][j] = i + 1; //save index size in another column 
            }
        }

    }


    */




        /*

                public DataTable ReadExcelValues(string locationStr, string keyWord, int searchOption, int wordSearchFilter)
                {
                    //declaring filtering flags
                    Boolean KeywordFilter = false;
                    Boolean CabinetLocFilter = false;
                    Boolean contains = false;

                    int newRowIndex = 0;
                    DataTable excelSearchVars = new DataTable();

                    excelSearchVars.Clear();        //clear excel table which stores search results

                    //Add data colomns 
                    excelSearchVars.Columns.Add("Part");
                    excelSearchVars.Columns.Add("Pack");
                    excelSearchVars.Columns.Add("Cabinet");
                    excelSearchVars.Columns.Add("Row");
                    excelSearchVars.Columns.Add("Drawer");
                    excelSearchVars.Columns.Add("Section");
                    excelSearchVars.Columns.Add("Parts_Info");
                    excelSearchVars.Columns.Add("Supplier");
                    excelSearchVars.Columns.Add("AltParts");
                    excelSearchVars.Columns.Add("Quantity");
                    excelSearchVars.Columns.Add("Index");

                    //###########TEST FEATURE######################################

                    //Char keyWordDelimiter = ' '; // use space as a delimitor

                    //String[] keyWordSpliter = keyWord.Split(keyWordDelimiter);

                    //#############################################################

                    List<string> deConstructerStr = new List<string>();
                    //split comma in string
                    Char delimiter = ',';
                    String[] cabinetCompareStr = locationStr.Split(delimiter);

                    //loop through all rows in excel spreadsheet
                    for (int i = 0; i < ExcelSpreadVals.Rows.Count; i++)
                    {
                        //check if search box contains special character 
                        if (keyWord == "*")
                            KeywordFilter = true;
                        else
                        {
                            //search by either part number or information this code will search through a selected column
                            //depending on the index option chosen by the optionBox
                            string rowString = (ExcelSpreadVals.Rows[i][searchOption] ?? String.Empty).ToString();

                            //Check what option is selected for word seach filter
                            if (wordSearchFilter == 0)  //search for strings starting with keyword
                                contains = rowString.StartsWith(keyWord, true, ci);
                            else                        //search for strings containing keyword
                                contains = rowString.IndexOf(keyWord, StringComparison.OrdinalIgnoreCase) >= 0;

                            //check if keyword in box matches part number column on spreadsheet
                            if (contains && !(String.IsNullOrWhiteSpace(keyWord)))
                                KeywordFilter = true;
                            else
                                KeywordFilter = false;
                        }

                        //SEARCH AND FILTER ENTRIES BASED ON A SPECIFIC LOCATION
                        if (cabinetCompareStr[0].ToString() != "*")
                        {
                            //construct an array based on number of elements selected in the form option boxes
                            for (int j = 0; j < cabinetCompareStr.Length; j++)
                            {
                                deConstructerStr.Add(ExcelSpreadVals.Rows[i][2 + j].ToString());
                            }
                            //check if guven
                            CabinetLocFilter = deConstructerStr.SequenceEqual(cabinetCompareStr);

                            deConstructerStr.Clear(); //reset and clear array
                        }
                        else
                            CabinetLocFilter = true;

                        //Now find results when filters are true

                        if (CabinetLocFilter == true && KeywordFilter == true)
                        {
                            excelSearchVars.Rows.Add();
                            for (int a = 0; a < excelSearchVars.Columns.Count; a++)
                            {
                                excelSearchVars.Rows[newRowIndex][a] = ExcelSpreadVals.Rows[i][a];
                            }
                            newRowIndex++;  //increment index of new table and 
                        }
                    }                    
                        return excelSearchVars;
                }
                */



        public Boolean UpdateQuantityValues(DataGridView resultValuesGridView, int amount)
        {

            int index = Convert.ToInt32(resultValuesGridView.SelectedRows[0].Cells[10].Value); //retrieve address index of row

            resultValuesGridView.SelectedRows[0].Cells[9].Value = amount; //write quantity to datagridview
            ExcelSpreadVals.Rows[index - 1][9] = amount;

            //Open Excel app for writing new data to 
            MyApp = new Excel.Application();
            hWnd = MyApp.Hwnd;
            PartsListClass.Hwind = hWnd;

            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(ExcelPath);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];            // Explict cast is not required here   


            MySheet.Cells[index, 10] = Convert.ToString(amount);    //write quantity amount to column 

            MyBook.Save();              //save quantity amount to excel spreadsheet

            //Quit Excel application and release all references 
            //and allow Garbage collector to handle the cleaning process
            MyBook.Close(false, Type.Missing, Type.Missing);

            MyApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.FinalReleaseComObject(MySheet);
            Marshal.FinalReleaseComObject(MyBook);
            Marshal.FinalReleaseComObject(MyApp);
            MyApp = null;
            MyBook = null;
            MySheet = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //Kill excel process that was ctreated in memory 
            KillProcessByMainWindows(hWnd);




            return true;
        }

        public Boolean EditComponentValues(DataGridView resultValuesGridView)
        {
            int RowIndex = Convert.ToInt32(resultValuesGridView.SelectedRows[0].Cells[10].Value); //retrieve address index of row

            //Open Excel app for writing new data to 
            MyApp = new Excel.Application();
            hWnd = MyApp.Hwnd;
            PartsListClass.Hwind = hWnd;

            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(ExcelPath);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];            // Explict cast is not required here   

            //loop through all columns minus the hidden adress index column
            int colSize = resultValuesGridView.Columns.Count - 1;
            //write values to Excel spreadsheet           

            for (int colIndex = 0; colIndex < colSize; colIndex++)
            {
                string colValues = Convert.ToString(resultValuesGridView.SelectedRows[0].Cells[colIndex].Value);
                MySheet.Cells[RowIndex, colIndex + 1] = colValues;
                ExcelSpreadVals.Rows[RowIndex - 1][colIndex] = colValues;
            }

            MyBook.Save();              //save quantity amount to excel spreadsheet
            //Quit Excel application and release all references 
            //and allow Garbage collector to handle the cleaning process
            MyBook.Close(false, Type.Missing, Type.Missing);

            MyApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.FinalReleaseComObject(MySheet);
            Marshal.FinalReleaseComObject(MyBook);
            Marshal.FinalReleaseComObject(MyApp);
            MyApp = null;
            MyBook = null;
            MySheet = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //Kill excel process that was ctreated in memory 
            KillProcessByMainWindows(hWnd);
         
        
      



            /*
            try
            {
                int index = Convert.ToInt32(resultValuesGridView.SelectedRows[0].Cells[10].Value); //retrieve address index of row
                // Open Excel app for writing new data to
 
             MyApp = new Excel.Application();
             hWnd = MyApp.Hwnd;
             PartsListClass.Hwind = hWnd;

                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(ExcelPath);
                MySheet = (Excel.Worksheet)MyBook.Sheets[1];            // Explict cast is not required here   

                for (int col = 0; col < resultValuesGridView.Columns.Count - 1; col++)
                {
                    MySheet.Cells[index][col] = resultValuesGridView.SelectedRows[0].Cells[col].Value;      //write quantity amount to entire row 
                }
                MyBook.Save();              //save quantity amount to excel spreadsheet

                //Quit Excel application and release all references 
                //and allow Garbage collector to handle the cleaning process
                MyBook.Close(false, Type.Missing, Type.Missing);

                MyApp.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.FinalReleaseComObject(MySheet);
                Marshal.FinalReleaseComObject(MyBook);
                Marshal.FinalReleaseComObject(MyApp);
                MyApp = null;
                MyBook = null;
                MySheet = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //Kill excel process that was ctreated in memory 
                KillProcessByMainWindows(hWnd);
                return true;
            }
            catch(Exception ex)
            {

            }
            */

            return true;
        }
        //Kill any Excel process that have been opened by this program 
        //Get the process ID number of the process and kill it from the memory
        public bool KillProcessByMainWindows(int hWnd)
        {
            uint processID;
            GetWindowThreadProcessId((IntPtr)hWnd, out processID);

            if (processID == 0) return false;
            try
            {
                Process.GetProcessById((int)processID).Kill();
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

    }


}
