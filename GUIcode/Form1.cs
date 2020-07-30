#define SearchTest 

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace ComponentSearcher
{
    public partial class Form1 : Form
    {
        //enums
        enum ComboOp { START, ALL, CONTAINS, NOTHING};
        enum ComboSearchOp {PARTNUM, PARTINFO, LOC};
        enum ComboLocType {ALLOCATED, SPARE};
        enum ComboCabinetType { ALL, CABINET, ROW, DRAWER, SECTION};
        enum WriteModeType{QUANTITY, EDIT, NEW_ENTRIES};

        //structures 
        public struct saveVars
        {             
            public string varPath; 
        };   

        public struct comboBoxPrevVals
        {
            public bool searchOptionBoxSave;
            public bool spareAllocateBoxSave;
            public bool locTypeBoxSave;
            public bool prtSearchComboSave;

            public bool prevCabinetBox;
            public bool prevRowBox;
            public bool preDrawerBox;
            public bool prevSectionBox;
        };

        //save button states 
        public struct buttonCheckPrev
        {
            public bool AddPtCheckedState;
        };

        //class declarations 
        saveVars varSave;

        comboBoxPrevVals locationVars;
        buttonCheckPrev buttonCheckStatus;

        //Global Variables decalarations 
        static int rowStart = 0;
        int rowIndexCounter = 0;

        //Class object declarations 
        private static MyExcel AppExcel = new MyExcel();
        public static DataTable table = new DataTable();     // this is to store variables
        public static DataTable stockTable = new DataTable(); //storing stock variables 

        //#####################Program begins here##############################
        public Form1()
            {
                InitializeComponent();

                NewCompDataGrid.DataSource = null;

                //disable single cell selection and only select row on datagridview
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
           
                //retrieve stored settings from memory
                //and load into structure 
                varSave.varPath = Properties.Settings.Default.DirPath;          
                saveCheckBox.Checked = Properties.Settings.Default.checkStatus;

                //display path location in text box
                PathDirBox.Text = varSave.varPath;    
                PathDirBox.Enabled = true;
           
                //Check if path was previously saved and open excel 
                if (saveCheckBox.Checked == true)
                {
                    PathDirBox.Enabled = false;
                    try
                    {
                        AppExcel.ExcelInit(varSave.varPath); //save path excel spreadsheet string
                    }
                    catch(Exception ex) //jump here if there is a problem with Excel path
                    {
                        MessageBox.Show(ex.ToString(), "Error");
                        PathDirBox.Clear();
                        PathDirBox.Enabled = true;
                        saveCheckBox.Checked = false;
                    }
                }

                //Initialise datagridview
                spreadsheetDatagridView1Inialise(dataGridView1);
                spreadsheetDatagridView1Inialise(NewCompDataGrid);

                //lock tab pages
                tabPage2.Enabled = false;               

                for (int i = 1; i <= 100; i++)
                {
                    string[] numbers = { i.ToString() };
                    comboCabinetBox.Items.AddRange(numbers);
                    cabinetComboAdd.Items.AddRange(numbers);
                }

            //setting combobox control defaults

            searchOptionBox.SelectedIndex = 0;
            PartSearchComboBox1.SelectedIndex = 1;
            SpareAllocatedComboBox.SelectedIndex = 0;
            locTypeComboBox.SelectedIndex = 0;

            //setting parts list combo box defaults
            cabinetComboAdd.SelectedIndex = 0;
            RowComboAdd.SelectedIndex = 0;
            drawerComboAdd.SelectedIndex = 0;
            sectionComboAdd.SelectedIndex = 0;

            this.dataGridView1.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);

        }

        //################## Related to opening up spreadsheet and saving directory location ############
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.InitialDirectory = "";
                openFileDialog1.Filter = "Excel files (*.xls, *.xlsx)|*.xls;*.xlsx";
                openFileDialog1.FilterIndex = 2;

                //save Excel path location string 
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    PathDirBox.Text = openFileDialog1.FileName;

                    //save path location into a structure variable 
                    varSave.varPath = openFileDialog1.FileName;

                    try
                    {
                        AppExcel.ExcelInit(varSave.varPath); //save path excel spreadsheet string
                                                             //Get string from File Dialog and send to check function to see if programs are current running

                        //clear contents on datagrid
                        slctedPart.Clear();
                        cabinetBox.Clear();
                        rowBox.Clear();
                        drawerBox.Clear();
                        sectionBox.Clear();
                        quantity.Clear();

                        spreadsheetDatagridView1Inialise(dataGridView1);
                }
                    catch (Exception ex) //jump here if there is a problem with Excel path
                    {
                        MessageBox.Show(ex.ToString(), "Error");
                        PathDirBox.Clear();
                        PathDirBox.Enabled = true;
                        varSave.varPath = null;
                    }
                    
            }

        
            /*
            //Keep displaying dialog box until satisfactory option has been selected  
            do
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //Get string from File Dialog and send to check function to see if programs are current running
                    textBox1.Text = openFileDialog1.FileName;

                    //save path location into a structure variable 
                    varSave.varPath = openFileDialog1.FileName;

                }   
                else
                {
                    //Break from loop if user wishes to cancel the open file dialog box option
                    break;      
                }
            } while (openExcel(varSave.varPath) == false);      //check to see if file is working properly
        */

        }

        private void saveCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (saveCheckBox.Checked)
            {
                if (String.IsNullOrWhiteSpace(varSave.varPath) != true)
                {
                    PathDirBox.Enabled = false;
                    browseButt.Enabled = false;
                    Properties.Settings.Default.DirPath = varSave.varPath;
                }
                else
                {
                    MessageBox.Show("No path selected");
                    PathDirBox.Enabled = true;
                    browseButt.Enabled = true;
                    saveCheckBox.Checked = false;
                }
            }
            else
            {
                PathDirBox.Enabled = true;
                browseButt.Enabled = true;
                Properties.Settings.Default.DirPath = null;
                //textBox1.Clear();
            }
            //save path location into memory and checked box status
            Properties.Settings.Default.checkStatus = saveCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        //################# Search and option controls and filters ############################################ 

            private void searchLocationButton_Click(object sender, EventArgs e)
            {
                int searchOption = 0; //stores variable for choose part option

                //Will construct a string based on the options selected 
                string optionStr = null;

                if (varSave.varPath != null && varSave.varPath != string.Empty)
                {
                    //search by part number or part information?
                    if (searchOptionBox.SelectedIndex == 0)
                        searchOption = 0; //search by part number
                    else if (searchOptionBox.SelectedIndex == 1)
                        searchOption = 6; //search by part information
                    else
                        searchOption = 9;  //search by SMD markings 

                    try
                    {

                        switch (locTypeComboBox.SelectedIndex)
                        {
                            case (int)ComboCabinetType.CABINET:
                                optionStr = comboCabinetBox.SelectedItem.ToString();
                                break;

                            case (int)ComboCabinetType.ROW:
                                optionStr = comboCabinetBox.SelectedItem.ToString() + "," + comboRowBox.SelectedItem.ToString();

                                break;
                            case (int)ComboCabinetType.DRAWER:
                                optionStr = comboCabinetBox.SelectedItem.ToString() + "," + comboRowBox.SelectedItem.ToString() + ","
                                    + comboDrawerBox.SelectedItem.ToString();
                                break;
                            case (int)ComboCabinetType.SECTION:
                                optionStr = comboCabinetBox.SelectedItem.ToString() + "," + comboRowBox.SelectedItem.ToString() + ","
                                    + comboDrawerBox.SelectedItem.ToString() + "," + comboSectionBox.SelectedItem.ToString();
                                break;

                            case (int)ComboCabinetType.ALL:
                                optionStr = "*"; // send special charcter if search is done all characters
                                break;

                            default:
                                break;
                        }


                        //send string and commands to read excel class
                        dataGridView1.DataSource = AppExcel.ReadExcelValues(optionStr, searchWordBox.Text, searchOption, PartSearchComboBox1.SelectedIndex, 0);
                        //dataGridView1.DataSource = AppExcel.ReadExcelValues(searchWordBox.Text);

                        if (dataGridView1.Rows.Count > 0)  //check if function returned a table that's not null
                        {
                            if (editModeEnable.Checked == true)
                            {
                                //editModeControls.Enabled = true;
                                EditModeHighlightGrid(dataGridView1.SelectedRows[0].Index);
                            }
                            else
                            {
                                //Higlight first line and display current items at location
                                highlightGrid(0);
                            }

                            //Display Found Entries as a string
                            resultsFound.Text = dataGridView1.RowCount.ToString();

                            //change focus to grid display
                            dataGridView1.Focus();
                        }
                        else
                        {
                            //editModeControls.Enabled = false;

                            //Clear boxes if nothing is found
                            slctedPart.Clear();
                            cabinetBox.Clear();
                            rowBox.Clear();
                            drawerBox.Clear();
                            sectionBox.Clear();
                            quantity.Clear();
                            resultsFound.Text = "No Results Found";
                        }

                    }
                    catch (Exception ex)
                    {
                        SpareAllocateButt.Enabled = false;
                        MessageBox.Show("One or more of the boxes must not be left blank!");
                    }
                }
                else
                {
                    SpareAllocateButt.Enabled = false;
                    MessageBox.Show("File is not valid please open valid document!", "ERROR");
                }

            }

            private void newSearchButt_Click(object sender, EventArgs e)
            {
                //initialisegrid
                spreadsheetDatagridView1Inialise(dataGridView1);

                labelMessage.Text = null;
                resultsFound.Text = null;

                //clear text boxes
                slctedPart.Clear();
                cabinetBox.Clear();
                rowBox.Clear();
                drawerBox.Clear();
                sectionBox.Clear();
                quantity.Clear();

            }

            private void locTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
            {
                switch (locTypeComboBox.SelectedIndex)
                {
                    case (int)ComboCabinetType.ALL:

                        cabinetControlsPanel.Enabled = false;   //disable all cabinet boxes
                        //set boxes display index
                        comboCabinetBox.SelectedIndex = -1;
                        comboRowBox.SelectedIndex = -1;
                        comboDrawerBox.SelectedIndex = -1;
                        comboSectionBox.SelectedIndex = -1;
                        break;

                    case (int)ComboCabinetType.CABINET:

                        cabinetControlsPanel.Enabled = true;   //enable cabinet boxes
                        //enable disable/individual boxes in panel
                        comboCabinetBox.Enabled = true;
                        comboRowBox.Enabled = false;
                        comboDrawerBox.Enabled = false;
                        comboSectionBox.Enabled = false;
                        break;
                    case (int)ComboCabinetType.ROW:

                        cabinetControlsPanel.Enabled = true;   //enable cabinet boxes
                        //enable disable/individual boxes in panel
                        comboCabinetBox.Enabled = true;
                        comboRowBox.Enabled = true;
                        comboDrawerBox.Enabled = false;
                        comboSectionBox.Enabled = false;
                        break;
                    case (int)ComboCabinetType.DRAWER:

                        cabinetControlsPanel.Enabled = true;   //enable cabinet boxes
                        //enable disable/individual boxes in panel
                        comboCabinetBox.Enabled = true;
                        comboRowBox.Enabled = true;
                        comboDrawerBox.Enabled = true;
                        comboSectionBox.Enabled = false;
                        break;
                    case (int)ComboCabinetType.SECTION:

                        cabinetControlsPanel.Enabled = true;   //enable cabinet boxes
                        //enable disable/individual boxes in panel
                        comboCabinetBox.Enabled = true;
                        comboRowBox.Enabled = true;
                        comboDrawerBox.Enabled = true;
                        comboSectionBox.Enabled = true;
                        break;

                    default:
                        break;
                }
            }

            private void SpareAllocatedComboBox_SelectedIndexChanged(object sender, EventArgs e)
            {

                switch (SpareAllocatedComboBox.SelectedIndex)
                {
                    case (int)ComboLocType.ALLOCATED:
                        searchWordBox.Enabled = true;
                        searchOptionBox.Enabled = true;
                        PartSearchComboBox1.Enabled = true;
                        searchWordBox.Clear();
                        break;
                    case (int)ComboLocType.SPARE:   //if SPARE OPTION SELECTED
                        searchWordBox.Enabled = false;
                        searchOptionBox.Enabled = false;
                        PartSearchComboBox1.Enabled = false;
                        searchWordBox.Text = "SPARE";           //use this as keyword
                        PartSearchComboBox1.SelectedIndex = 1;  //search by contain
                        searchOptionBox.SelectedIndex = 0;      //search by part number
                        break;

                    default:
                        break;
                }

            }

            private void searchWordBox_KeyDown(object sender, KeyEventArgs e)
            {
                //if Enter key has been pressed then perform 
                //search button action 
                if (e.KeyCode == Keys.Enter)
                {
                    searchLocationButton.PerformClick();
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    return;
                }
            }


        //############################ Data Grid results control  ########################################################## 

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows.Count <= 0)  //if no data is in cells
            {
                return;
            }

            if (editModeEnable.Checked != true) //check if edit mode is on
            {
                //get current index of selected row and display on screen
                int index = dataGridView1.CurrentCell.RowIndex;
                slctedPart.Text = dataGridView1.Rows[index].Cells[0].Value.ToString();
                displayItemLocation(index);
            }

            if (editModeEnable.Checked && editChange_Butt.Enabled != true)    //make sure button is not visible
            {
                EditModeHighlightGrid(dataGridView1.CurrentCell.RowIndex);
            }
            if (editModeEnable.Checked && editChange_Butt.Enabled == true)
                dataGridView1.ClearSelection();                
        }    

   
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

            int index = 0;
            int incIndex = 0;

            //if (editModeEnable.Checked != true)
            //{

                //Check if down key was pushed 
                if (e.KeyCode == Keys.Down)
                {
                    //does datagrid contain any data?
                    if (dataGridView1.RowCount > 0)
                    {
                        //Get current value of the cell 
                        index = dataGridView1.CurrentCell.RowIndex;

                        //take a copy of the index and get next position
                        incIndex = index;
                        ++incIndex;

                        //Check that the next position is not greater than the number of 
                        //available rows 
                        if (incIndex != dataGridView1.RowCount)
                        {
                            //select and highlight the current position on the datagridview
                            dataGridView1.Rows[index].Selected = false;
                            dataGridView1.Rows[++index].Selected = true;

                            //Display items on GUI pass current position to function 
                            displayItemLocation(index);
                        }

                    }


                }
                else if (e.KeyCode == Keys.Up)
                {

                    if (dataGridView1.SelectedRows.Count > 0)
                    {
                        //Get current value of the cell 
                        index = dataGridView1.CurrentCell.RowIndex;

                        //take a copy of the index and get previous position
                        incIndex = index;
                        --incIndex;

                        if (incIndex >= 0)
                        {
                            dataGridView1.Rows[index].Selected = false;
                            dataGridView1.Rows[--index].Selected = true;

                            slctedPart.Text = dataGridView1.Rows[index].Cells[0].Value.ToString();

                            //Display items on GUI
                            displayItemLocation(index);
                        }
                    }
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    searchWordBox.Focus();
                }
                else
                {
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                }
            //}
        }
                
        private void displayItemLocation(int index)
        {
            //Display part information in box
            slctedPart.Text = dataGridView1.Rows[index].Cells[0].Value.ToString();

            //Display items on GUI
            cabinetBox.Text = dataGridView1.Rows[index].Cells[2].Value.ToString();
            rowBox.Text = dataGridView1.Rows[index].Cells[3].Value.ToString();
            drawerBox.Text = dataGridView1.Rows[index].Cells[4].Value.ToString();
            sectionBox.Text = dataGridView1.Rows[index].Cells[5].Value.ToString();
            quantity.Text = dataGridView1.Rows[index].Cells[10].Value.ToString();
        }

        private void spreadsheetDatagridView1Inialise(DataGridView grid)
        {

            DataTable InitTable = new DataTable();

            //Create table for adding values to spreadsheet
            InitTable.Columns.Add("Part");
            InitTable.Columns.Add("Pack");
            InitTable.Columns.Add("Cabinet");
            InitTable.Columns.Add("Row");
            InitTable.Columns.Add("Drawer");
            InitTable.Columns.Add("Section");
            InitTable.Columns.Add("Part Info");
            InitTable.Columns.Add("Supplier");
            InitTable.Columns.Add("Alt Parts");
            InitTable.Columns.Add("SMD Marking");
            InitTable.Columns.Add("Quantity", typeof(int));
            InitTable.Columns.Add("Index");                       

            grid.RowHeadersVisible = false;
            grid.DataSource = InitTable;
            grid.AutoGenerateColumns = true;

            //******For future reference if checkboxes are required********************
            /*
            DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
            grid.Columns.Add(chk);
            */       

            DataGridViewColumn column0 = grid.Columns[0];
            column0.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column0.Width = 150;
            column0.SortMode = DataGridViewColumnSortMode.NotSortable;
            column0.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column1 = grid.Columns[1];
            column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column1.Width = 100;
            column1.SortMode = DataGridViewColumnSortMode.NotSortable;
            column1.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column2 = grid.Columns[2];
            column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column2.Width = 80;
            column2.SortMode = DataGridViewColumnSortMode.NotSortable;
            column2.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column3 = grid.Columns[3];
            column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column3.Width = 80;
            column3.SortMode = DataGridViewColumnSortMode.NotSortable;
            column3.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column4 = grid.Columns[4];
            column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column4.Width = 80;
            column4.SortMode = DataGridViewColumnSortMode.NotSortable;
            column4.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column5 = grid.Columns[5];
            column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column5.Width = 80;
            column5.SortMode = DataGridViewColumnSortMode.NotSortable;
            column5.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column6 = grid.Columns[6];
            column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column6.Width = 200;
            column6.SortMode = DataGridViewColumnSortMode.NotSortable;
            column6.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column7 = grid.Columns[7];
            column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column7.Width = 100;
            column7.SortMode = DataGridViewColumnSortMode.NotSortable;
            column7.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column8 = grid.Columns[8];
            column8.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column8.Width = 150;
            column8.SortMode = DataGridViewColumnSortMode.NotSortable;
            column8.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column9 = grid.Columns[9];
            column9.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column9.Width = 80;
            column9.SortMode = DataGridViewColumnSortMode.NotSortable;
            column9.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column10 = grid.Columns[10];
            column10.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column10.Width = 70;
            column10.SortMode = DataGridViewColumnSortMode.NotSortable;
            column10.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            DataGridViewColumn column11 = grid.Columns[11];
            column11.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column11.Width = 70;
            column11.Visible = false;
            column11.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowLocationClick = e.RowIndex;           //detect which row has been selected must be greater than 0

            if (rowLocationClick >= 0)
            {
                //if double clicked modify and edit the selected item
                if (editModeEnable.Checked == true && editChangeModeControls.Enabled != true) //check if edit mode is enabled
                {
                    preRowTable.Reset();        //This will store the selected previous value

                    int index = dataGridView1.SelectedRows[0].Index; //save index of selected part 

                    //take a copy of previous value
                    foreach (DataGridViewColumn cols in dataGridView1.Columns)
                        preRowTable.Columns.Add();

                    try
                    {
                        for (int x = 0; x < dataGridView1.Rows.Count; x++)
                        {
                            if (x != index)
                                dataGridView1.Rows[x].Visible = false;
                            if (x == index) //save index of previous value
                            {
                                preRowTable.Rows.Add();
                                for (int a = 0; a < preRowTable.Columns.Count; a++) //save previous value of row
                                {
                                    preRowTable.Rows[0][a] = (dataGridView1.Rows[index].Cells[a].Value ?? String.Empty).ToString();
                                }
                            }

                        }

                        //clear grid and make writable
                        dataGridView1.ClearSelection();
                        dataGridView1.ReadOnly = false;

                        //Enable disable controls
                        editChangeModeControls.Enabled = true;
                        editModeControls.Enabled = false;
                        quantityControlsPanel.Enabled = false;
                        searchControlPanel.Enabled = false;

                        //disable buttons
                        editModeEnable.Enabled = false;

                        //make various panels visible and invisible
                        editModeControls.Visible = false;
                        editChangeModeControls.Visible = true;
                        quantityControlsPanel.Visible = false;

                        searchControlPanel.Visible = false;

                        //disable cell click event handers 
                        this.dataGridView1.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
                        // this.dataGridView1.CellDoubleClick -= dataGridView1_CellDoubleClick;
                        //clear grid
                        dataGridView1.ClearSelection();
                        //SpareAllocateButt.Enabled = false;  

                        tabPage2.Enabled = false;

                        if (changeAcceptEditPart.Visible)
                        {
                            changeAcceptEditPart.Enabled = false;                                            
                        }
                    }
                    catch
                    {
                        //SpareAllocateButt.Enabled = false;
                        //int y = 0;
                    }                 
                   
                }
            }

        }

        private void EditModeHighlightGrid(int index)
        {
            dataGridView1.MultiSelect = true;
            dataGridView1.ClearSelection();

            //Highlight all the cells you don't want to change
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (i != index)
                    dataGridView1.Rows[i].Selected = true;
                if (i == index)
                {
                    SpareAllocateButt.Enabled = true;
                }
            }
            //Display location items in boxex  
            displayItemLocation(index);
        }

        private void highlightGrid(int index)   //highlight and display locations
        {
            if (dataGridView1.Rows.Count > 0) //make sure there is data displayed on grid before proceding
            {
                //view data select line
                dataGridView1.Rows[index].Selected = true;

                //Display location items in boxex  
                displayItemLocation(index);
            }
        }

        //############################ Edit and modify controls ###########################################################         

            private void editModeEnable_CheckedChanged(object sender, EventArgs e)
            {
                try
                {
                    //Check which option has been selected
                    if (editModeEnable.Checked == true)
                    {                       

                        DialogResult result = MessageBox.Show("You are about to enter Edit Mode do you wish to proceed?", "Warning", MessageBoxButtons.YesNo);

                        switch (result)
                        {
                            case DialogResult.Yes:
                                {
                                    //enable tab2
                                    tabPage2.Enabled = true;
                                    //enable edit features 
                                    editModeControls.Enabled = true;

                                    //enable edit mode controls
                                    editModePanel.Visible = true;

                                   // addLocationSparePanel.Enabled = true;

                                //this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
                                this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
                                    
                                //check for data on grid 
                                    if (dataGridView1.Rows.Count > 0)
                                    {
                                        //highlight all cells except for the selected one 
                                        EditModeHighlightGrid(dataGridView1.SelectedRows[0].Index);
                                    }
                                    break;
                                }
                            case DialogResult.No:
                                {
                                    //enable tab2
                                    tabPage2.Enabled = false;

                                    //displays current row and cancels controls 
                                    dataGridView1.MultiSelect = false;
                                    dataGridView1.ReadOnly = true;

                                    editModePanel.Visible = false;

                                    //disable edit mode controls 
                                    editModeControls.Enabled = false;

                                    //disable event handler edit button event handler before changing state to avoid jumping to event
                                    this.editModeEnable.CheckedChanged -= editModeEnable_CheckedChanged;
                                    editModeEnable.Checked = false;
                                    //renable event handler
                                    this.editModeEnable.CheckedChanged += editModeEnable_CheckedChanged;                                        

                                    this.dataGridView1.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);

                                    editModeControls.Enabled = false;
                                    if (dataGridView1.Rows.Count > 0)
                                    {
                                        //Higlight grid
                                        highlightGrid(dataGridView1.CurrentRow.Index);
                                    }
                                    break;
                                }
                        }
                    return;
                    }

                        if(editModeEnable.Checked == false)
                        {
                            //display result dialog box
                            DialogResult result2 = MessageBox.Show("You are about to leave Edit Mode do you wish to continue?","Warning", MessageBoxButtons.YesNo);

                            switch (result2)
                            {
                                case DialogResult.Yes :

                                    //disable event handler edit button event handler before changing state to avoid jumping to event
                                    this.editModeEnable.CheckedChanged -= editModeEnable_CheckedChanged;
                                    editModeEnable.Checked = false;
                                    //renable event handler
                                    this.editModeEnable.CheckedChanged += editModeEnable_CheckedChanged;

                                    dataGridView1.MultiSelect = false;
                                    dataGridView1.ReadOnly = true;

                                    //disable edit mode controls 
                                    editModeControls.Enabled = false;

                                    editModePanel.Visible = false;

                                    //enable tab2
                                    tabPage2.Enabled = false;

                                    this.dataGridView1.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);


                                    if (dataGridView1.Rows.Count > 0)   //display and higlight current gird if edit mode not enabled
                                    {                                   
                                        //Higlight grid current index
                                        highlightGrid(dataGridView1.CurrentRow.Index);
                                    }
                                    break;                                     

                                case DialogResult.No :

                                    this.editModeEnable.CheckedChanged -= editModeEnable_CheckedChanged;
                                    editModeEnable.Checked = true;
                                    this.editModeEnable.CheckedChanged += editModeEnable_CheckedChanged;

                                    //disable edit mode controls 
                                    editModeControls.Enabled = true;
                                    //enable tab2
                                     tabPage2.Enabled = true;

                                    editModePanel.Visible = true;
                            break;
                            }

                        return;
                        }                        
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("Unable to enter edit mode please ensure sufficient information is displayed in grid","Error!");

                        //Edit mode events and check box status changes
                        this.editModeEnable.CheckedChanged -= editModeEnable_CheckedChanged;                        

                        editModeEnable.Checked = false;

                        //renable event handler
                        this.editModeEnable.CheckedChanged += editModeEnable_CheckedChanged;

                        //disable edit mode controls
                        editModeControls.Enabled = false;

                        editModePanel.Visible = false;

                         //this.dataGridView1.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);

                        //disable double click event handler 
                        //this.dataGridView1.CellDoubleClick -= dataGridView1_CellDoubleClick;    
                        this.dataGridView1.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);

            }
            }         
        
                    DataTable preRowTable = new DataTable();     //save variables to this property      


        private void editChange_Butt_Click(object sender, EventArgs e)
        {
            int y = 0;
            DialogResult result = MessageBox.Show("Would you like to proceed with the following changes?", "Save Changes", MessageBoxButtons.OKCancel);

            switch (result)
            {
                //if result is yes do the following
                case DialogResult.OK:

                    //make invisble rows repear again!
                    int index = dataGridView1.CurrentRow.Index;

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        //make rows visble again
                        if (i != index)
                        {
                            dataGridView1.Rows[i].Visible = true;   //make all rows visible again
                            dataGridView1.Rows[i].Selected = true;  //selected all rows that aren't equal to the index
                        }
                    }
                    //display location of index
                    displayItemLocation(index);

                    dataGridView1.ReadOnly = true;

                    AppExcel.WriteFunctionTest(dataGridView1, 1);     //send value to function for editing purposes 
                                                                      //disable buttons                       
                                                                      //enable/disable controls


                    //enable/disable controls
                    editModeControls.Enabled = true; //enable edit button controls
                    editChangeModeControls.Enabled = false;

                    editModeControls.Visible = true;
                    editChangeModeControls.Visible = false;
                    quantityControlsPanel.Visible = true;

                    quantityControlsPanel.Enabled = true;

                    if (changeAcceptEditPart.Visible != true)
                    {
                        editModeEnable.Enabled = true;
                        searchControlPanel.Enabled = true;
                        searchControlPanel.Visible = true;
                    }

                    if (changeAcceptEditPart.Visible)
                    {
                        //editModeControls.Enabled = false; //enable edit button controls
                        searchControlPanel.Visible = false;
                    }

                    changeAcceptEditPart.Enabled = true;
                    tabPage2.Enabled = true;

                    this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
                    //this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
                    break;
            }
        }

        private void cancelEditChangeButt_Click(object sender, EventArgs e)
        {
            //make invisble rows repear again!
            int index = dataGridView1.CurrentRow.Index;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //make rows visble again
                if (i != index)
                {
                    dataGridView1.Rows[i].Visible = true;   //make rows visible
                    dataGridView1.Rows[i].Selected = true;  //selected non selected rows
                }
                if (i == index)
                {
                    for (int col = 0; col <= dataGridView1.Columns.Count - 1; col++)
                    {
                        dataGridView1.Rows[index].Cells[col].Value = (preRowTable.Rows[0][col] ?? String.Empty).ToString();
                        //  preRowTable.Rows[0][a] = (dataGridView1.Rows[index].Cells[a].Value ?? String.Empty).ToString();
                    }
                }

            }

            //enable/disable controls
            editModeControls.Enabled = true; //enable edit button controls
            editChangeModeControls.Enabled = false;

            editModeControls.Visible = true;
            editChangeModeControls.Visible = false;
            quantityControlsPanel.Visible = true;    

            quantityControlsPanel.Enabled = true;

            if (changeAcceptEditPart.Visible != true)
            {
                editModeEnable.Enabled = true;
                searchControlPanel.Enabled = true;
                searchControlPanel.Visible = true;
            }

            if (changeAcceptEditPart.Visible)
            {
                //editModeControls.Enabled = false; //enable edit button controls
                searchControlPanel.Visible = false;
            }

            changeAcceptEditPart.Enabled = true;

            //make grid readonly
            dataGridView1.ReadOnly = true;

            tabPage2.Enabled = true;

            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
        }


        private void SpareAllocateButt_Click(object sender, EventArgs e)
            {
                try
                {
                    //check if the datagrid contains data
                    if (dataGridView1.Rows.Count > 0)
                    {
                        int index = dataGridView1.CurrentRow.Index;

                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            //make rows visble again
                            if (i == index)         //write blank items in cell
                            {
                                dataGridView1.Rows[index].Cells[0].Value = "SPARE BIN";
                                dataGridView1.Rows[index].Cells[1].Value = " ";
                                dataGridView1.Rows[index].Cells[6].Value = " ";
                                dataGridView1.Rows[index].Cells[7].Value = " ";
                                dataGridView1.Rows[index].Cells[8].Value = " ";
                                dataGridView1.Rows[index].Cells[9].Value = " ";
                                dataGridView1.Rows[index].Cells[10].Value = " ";
                            }

                        }
                        dataGridView1.ReadOnly = true;
                        editModeEnable.Enabled = true;

                        AppExcel.WriteExcelValues(dataGridView1.CurrentRow);

                        MessageBox.Show("Part successfully written to spreadsheet!", "Message");
                        //MessageBox.Show.Close();
                        /*############
                         * send modified text to function
                         */
                        //AppExcel.WriteExcelValues(dataGridView1.CurrentRow);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Problem writting data to spreadsheet!", "ERROR");
                }
        }

            //********************stock quantity controls***************************************************************
            private void quantityAddBut_Click(object sender, EventArgs e)
            {
                //initialise variables
                string amountStr = null;
                int amount = 0;
                int index = 0;                

                try
                {
                    //get current index of selected row 
                    index = dataGridView1.CurrentRow.Index;

                    if (quantityBox.Value > 0)
                    {
                        //get selected quantity amount from the row
                        amountStr = (dataGridView1.CurrentRow.Cells[10].Value ?? String.Empty).ToString();

                        //check if the quantity slected contains a value
                        if (String.IsNullOrWhiteSpace(amountStr))
                        {
                            amount = 0;
                        }
                        else
                        {
                            amount = Convert.ToInt16(amountStr);
                        }
                        amount = Convert.ToInt16(quantityBox.Value) + amount;                      
                    
                        //write selected row to the function
                        dataGridView1.Rows[index].Cells[10].Value = amount;
                        //AppExcel.WriteExcelValues(dataGridView1.CurrentRow);

                        AppExcel.WriteFunctionTest(dataGridView1, 1);     //send value to function for editing purposes 

                }
                    else
                    {
                        MessageBox.Show("Quantity Box contains no amount!", "Warning");
                    }

                    //check which mode is selected for highlighting items 
                    //check what mode the datagrid is in
                    if (editModeEnable.Checked)
                    {
                        //Highlight all the cells you don't want to change
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (i != index)
                                dataGridView1.Rows[i].Selected = true;
                        }
                     }
                    else
                    {
                        dataGridView1.Rows[index].Selected = true;
                    }

                    quantityBox.Value = 0;  

                    //Display searched items in boxes and set the row hightlighter to zero    
                    displayItemLocation(index);                  
                    
                }
                catch (Exception ex)
                {
                    DialogResult errorResult = MessageBox.Show("Invalid row selection or no data on grid!", "Error!");                    
                }

            }

            private void quantityRemoveButt_Click(object sender, EventArgs e)
            {
                //initialise variables
                string amountStr = null;
                int amount = 0;
                int index = 0;

                try
                {
                    //get current index of selected row 
                    index = dataGridView1.CurrentRow.Index;
                    if (quantityBox.Value > 0)
                    {

                        /*
                        //check which mode the is in operation and get current index adress
                        if (editModeEnable.Checked)
                        {
                            index = dataGridView1.CurrentCell.RowIndex;
                        }
                        else
                        {
                            index = dataGridView1.CurrentRow.Index;
                        }
                        */
                        //get selected quantity amount from the row
                        amountStr = (dataGridView1.CurrentRow.Cells[10].Value ?? String.Empty).ToString();

                        if (String.IsNullOrWhiteSpace(amountStr))
                        {
                            amount = 0;
                        }
                        else
                        {
                            amount = Convert.ToInt16(amountStr);
                        }

                        amount = amount - Convert.ToInt16(quantityBox.Value);
                    
                        if(amount < 0)
                        {
                            MessageBox.Show("Amount selected is more than what's avaiable!", "ERROR");
                            quantityBox.Value = 0;
                        }
                        else
                        {
                            //check which mode the is in operation and get current index adress
                            if (editModeEnable.Checked)
                            {
                                //Highlight all the cells you don't want to change
                                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                {
                                    if (i != index)
                                        dataGridView1.Rows[i].Selected = true;
                                }
                            }
                            else
                            {
                                dataGridView1.Rows[index].Selected = true;
                            }

                            //send updated value to function
                            dataGridView1.Rows[index].Cells[10].Value = amount;

                            //write amount to function
                            //AppExcel.WriteExcelValues(dataGridView1.CurrentRow);
                            AppExcel.WriteFunctionTest(dataGridView1, 1);     //send value to function for editing purposes 
                            quantityBox.Value = 0;         

                            //Display searched items in boxes and set the row hightlighter to zero    
                            displayItemLocation(index);

                    }
                }
            }
                catch (Exception ex)
                {
                    //select datagrid to highlight the first line 
                    //displayItemLocation(dataGridView1.CurrentRow.Index);
                    quantityBox.Value = 0;
                    MessageBox.Show("Invalid Quantity Selection", "Error");
                }
        }

        //#############################Add Entry Mode Controls################################################################
      
            //********************************* add row controls**************************************************************
            private void newEntryButAdd_Click(object sender, EventArgs e)
            {
                if (editModeEnable.Checked)        // Add row mode been enabled?.
                {
                    //int index = 0;
                    int rowNum = (int)rowCountNum.Value;

                    if (table.Columns.Count <= 0)
                    {
                        //Add columns to the data and display in spreadsheet
                        table.Columns.Add("Part");
                        table.Columns.Add("Pack");
                        table.Columns.Add("Cabinet");
                        table.Columns.Add("Row");
                        table.Columns.Add("Drawer");
                        table.Columns.Add("Section");
                        table.Columns.Add("Part Info");
                        table.Columns.Add("Supplier");
                        table.Columns.Add("Alt Parts");
                        table.Columns.Add("SMD Marking");
                        table.Columns.Add("Quantity");
                        table.Columns.Add("Index");
                    }                     

                //Add new row to spreadsheet
                for (int index = 0; index < rowNum; index++)
                {
                    table.Rows.Add();       //add more rows
                }                   

                    //create new table to display columns and rows on 
                    //datagrid view
                    NewCompDataGrid.DataSource = table;

                    if (NewCompDataGrid.SelectedRows.Count > 0)
                    {
                        addLocationSparePanel.Enabled = true;
                        // addLocationButt.Enabled = true;
                        RemoveEntryButt.Enabled = true;
                        ExcelAddEntryButt.Enabled = true;
                        //spareAllocateButton.Enabled = true;

                        //cabinetComboAdd.Enabled = true;
                        //RowComboAdd.Enabled = true;
                        //drawerComboAdd.Enabled = true; 
                        //sectionComboAdd.Enabled = true;

                        //highlight current location on grid
                        //Highlight latest line
                        //view data select line
                        NewCompDataGrid.CurrentRow.Selected = true;
                    }
                    else
                    {
                        addLocationSparePanel.Enabled = false;
                        //addLocationButt.Enabled = false;
                        RemoveEntryButt.Enabled = false;
                        ExcelAddEntryButt.Enabled = false;                     

                    }

                /*


                    //int index = 0;
                    int rowNum = (int)rowCountNum.Value;


                    //Add new row to spreadsheet
                    for (int index = 0; index < rowNum; index++)
                    {
                        table2.Rows.Add();       //add more rows
                    }

                    //create new table to display columns and rows on 
                    //datagrid view
                    NewCompDataGrid.DataSource = table2;



               
                    if (table.Columns.Count <= 0)
                    {
                        table.Columns.Add("Part");
                        table.Columns.Add("Pack");
                        table.Columns.Add("Cabinet");
                        table.Columns.Add("Row");
                        table.Columns.Add("Drawer");
                        table.Columns.Add("Section");
                        table.Columns.Add("Part Info");
                        table.Columns.Add("Supplier");
                        table.Columns.Add("Alt Parts");
                        table.Columns.Add("Quantity");
                        table.Columns.Add("Index");
                    }

                    //Add new row to spreadsheet
                    for (int index = 0; index < rowNum; index++)
                    {
                        table.Rows.Add();       //add more rows
                    }

                    //create new table to display columns and rows on 
                    //datagrid view
                    NewCompDataGrid.DataSource = table;

                    */
            }
            }

            private void RemoveEntryButt_Click(object sender, EventArgs e)
            {
                //Get total number of items selected
                int selectedItemCount = NewCompDataGrid.SelectedRows.Count;
                int index = 0;

                for (int i = 0; i < selectedItemCount; i++)
                {
                    //get address of current selected index 
                    index = NewCompDataGrid.SelectedRows[0].Index;
                    NewCompDataGrid.Rows.RemoveAt(index);
                }

                if (NewCompDataGrid.Rows.Count > 0)
                {
                    addLocationSparePanel.Enabled = true;
                    RemoveEntryButt.Enabled = true;

                    ExcelAddEntryButt.Enabled = true;                   
                    
                    //Highlight latest line
                    //view data select line
                    NewCompDataGrid.CurrentRow.Selected = true;

                }
                else
                {
                    addLocationSparePanel.Enabled = false;

                    RemoveEntryButt.Enabled = false;
                    ExcelAddEntryButt.Enabled = false;                 
                }

        }

            private void resetAddGrid_Click(object sender, EventArgs e)
            {
               // if(NewCompDataGrid.Rows.Count > 0)  //check if grid contains any data
                spreadsheetDatagridView1Inialise(NewCompDataGrid);
                table.Clear();   //clear table as well
                rowCountNum.Value = 0;
                ExcelAddEntryButt.Enabled = false;
                RemoveEntryButt.Enabled = false;
            addLocationSparePanel.Enabled = false;
            }

            private void spareAllocateButton_Click(object sender, EventArgs e)
            {
                // check if grid contains any data
                if (NewCompDataGrid.Rows.Count > 0)
                {                
                    //Get total number of items selected
                    int selectedItemCount = NewCompDataGrid.SelectedRows.Count;
                    int index = 0;

                    for (int i = 0; i < selectedItemCount; i++)
                    {
                        //get address of current selected index 
                        index = NewCompDataGrid.SelectedRows[i].Index;
                        NewCompDataGrid.Rows[index].Cells[0].Value = "SPARE BIN";
                    }
                }


            }
        
            private void addLocationButt_Click(object sender, EventArgs e)
            {
                //check if grid contains data
                if (NewCompDataGrid.Rows.Count > 0)
                {             
                    
                    //Check if location exists and write to grid 
                    string cabinetCompare = cabinetComboAdd.SelectedItem.ToString() + "," + RowComboAdd.SelectedItem.ToString() + "," +
                                            drawerComboAdd.SelectedItem.ToString() + "," + sectionComboAdd.SelectedItem.ToString();

                    //read and check spreadsheet 
                    //AppExcel.ReadExcelValues(cabinetCompare, null, 0, 0, 1);
                    int Count = AppExcel.ReadExcelValues(cabinetCompare, null, 0, 0, 1).Rows.Count;                   

                    //Check if box contains 
                    if (Count > 0)
                    {
                        // MessageBox.Show("This following location is already in used by " + AppExcel.ReadExcelValues(cabinetCompare, null, 0, 0, 1).Rows.Count + 
                        //      " others do you wish to change ?","Warning");
                        //MessageBox.Show("You are about to enter Edit Mode do you wish to proceed?", "Warning", MessageBoxButtons.YesNo);
                        DialogResult result = MessageBox.Show(Count + " locations already in use do you wish to edit or change?", "Warning", MessageBoxButtons.YesNo); 
                        
                        switch(result)
                        {

                            case DialogResult.Yes:
                                dataGridView1.DataSource = AppExcel.ReadExcelValues(cabinetCompare, null, 0, 0, 1);
                            editModeEnable.Checked = true;  //change to edit mode 
                            searchControlPanel.Enabled = false;  //disable edit control panel
                            tabPage2.Enabled = false;
                            tabControl1.SelectedTab = SearchTab;
                            changeAcceptEditPart.Visible = true;
                            editModeEnable.Enabled = false;

                            searchControlPanel.Visible = false;

                            //highlight grid and display current location in grid
                            EditModeHighlightGrid(0);

                            //Display Found Entries as a string
                            resultsFound.Text = Count.ToString();
                            break;

                            case DialogResult.No:
                                break;

                        }
                        //      " others do you wish to change ?","Warning", MessageBoxButtons.YesNo);
                    }
                    else
                    {
                        //write values to grid                
                        NewCompDataGrid.SelectedCells[2].Value = cabinetComboAdd.SelectedItem.ToString();
                        NewCompDataGrid.SelectedCells[3].Value = RowComboAdd.SelectedItem.ToString();
                        NewCompDataGrid.SelectedCells[4].Value = drawerComboAdd.SelectedItem.ToString();
                        NewCompDataGrid.SelectedCells[5].Value = sectionComboAdd.SelectedItem.ToString();
                    }
            }

            }
        
            private void ExcelAddEntryButt_Click(object sender, EventArgs e)
            {
                //check if grid contains any data
                if (NewCompDataGrid.Rows.Count > 0)
                {
                    try
                    {
                        DialogResult result = MessageBox.Show("The following Entries will be written to spreadsheet do you wish to continue?", "Message", MessageBoxButtons.OKCancel);

                        switch (result)
                        {
                            case DialogResult.OK:

                            //AppExcel.WriteExcelValues2(NewCompDataGrid);
                            AppExcel.WriteFunctionTest(NewCompDataGrid, 0); //Zero for adding new entries                         

                            //clear data grid
                            spreadsheetDatagridView1Inialise(NewCompDataGrid);
                            spreadsheetDatagridView1Inialise(dataGridView1);
                            //Clear boxes if nothing is found
                            slctedPart.Clear();
                            cabinetBox.Clear();
                            rowBox.Clear();
                            drawerBox.Clear();
                            sectionBox.Clear();
                            quantity.Clear();
                            resultsFound.Text = "No Results Found";


                            rowCountNum.Value = 0;
                            //disable controls
                            RemoveEntryButt.Enabled = false;
                            ExcelAddEntryButt.Enabled = false;
                            addLocationSparePanel.Enabled = true;                                    
                            table.Clear();   //clear table as well

                            //Show success message
                            MessageBox.Show("Items successfully written to spreadsheet", "Excel Write");
                            break;

                            case DialogResult.Cancel:
                                break;
                        }

                    }
                    catch(Exception ex)
                    {
                        //when there is a problem writing values to spreadsheet
                    }
                }

            }

            private void deleteEntryButt_Click(object sender, EventArgs e)
            {
                //delete entries 
                AppExcel.WriteFunctionTest(dataGridView1, 3); //delete selected entry

                if (dataGridView1.Rows.Count > 0)
                {
                    //Highlight current and existing line
                    EditModeHighlightGrid(0);
                    resultsFound.Text = dataGridView1.Rows.Count.ToString();
            }
                else
                {
                    //Clear boxes if nothing is found
                    slctedPart.Clear();
                    cabinetBox.Clear();
                    rowBox.Clear();
                    drawerBox.Clear();
                    sectionBox.Clear();
                    quantity.Clear();
                    resultsFound.Text = "No Results Found";
                }
            }

        //################# Form Termination controls ######################################################################

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Kill any Excel processes left hanging in the memory
            AppExcel.KillProcessByMainWindows(PartsListClass.Hwind);
        }

        private void changeAcceptEditPart_Click(object sender, EventArgs e)
        {
            labelMessage.Text = null;
            resultsFound.Text = null;

            //clear text boxes
            slctedPart.Clear();
            cabinetBox.Clear();
            rowBox.Clear();
            drawerBox.Clear();
            sectionBox.Clear();
            quantity.Clear();

            changeAcceptEditPart.Visible = false;
            tabPage2.Enabled = true;
            tabControl1.SelectedTab = tabPage2;
            searchControlPanel.Enabled = true;

            searchControlPanel.Visible = true;

            editModeEnable.Enabled = true;
            spreadsheetDatagridView1Inialise(dataGridView1);

        }

        private void searchAllCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if(searchAllCheckBox.Checked)
            {
                SpareAllocatedComboBox.Enabled = false;
                SpareAllocatedComboBox.SelectedIndex = 0;
                searchWordBox.Text = "*";           //Search all
                searchWordBox.Enabled = false;                
            }
            else
            {
                SpareAllocatedComboBox.Enabled = true;
                searchWordBox.Clear();           //Search all
                searchWordBox.Enabled = true;
            }
        }






















#if !SearchTest

        /*
private bool openExcel(string path)
{        

    bool state = false; 
    try
    {
        //initialise excel and open sheet
        AppExcel.ExcelInit(varSave.varPath);

        //initiliase all datagrids 
        spreadsheetDatagridView1Inialise(dataGridView1);
        spreadsheetDatagridView1Inialise(stockListGridView);
        spreadsheetDatagridView1Inialise(NewCompDataGrid);
        state = true;
    }
    catch (ExcelExceptionMessage ex)
    {
        //catch and display error message if Excel has problems initialising 
        DialogResult errorResult = MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OKCancel);
        //store previously good working path in structure 
        varSave.varPath = PartsListClass.PrevDirPath;
        textBox1.Text = varSave.varPath;

        if (errorResult == DialogResult.OK)
        {
            textBox1.Enabled = true;
            saveCheckBox.Checked = false;

            //Kill excel process that was created in memory         
            AppExcel.KillProcessByMainWindows(PartsListClass.Hwind);
            state = false;     

        }
        else
        {
            textBox1.Enabled = true;
            varSave.varPath = PartsListClass.PrevDirPath;
            saveCheckBox.Checked = false;

            //Kill excel process that was created in memory
            AppExcel.KillProcessByMainWindows(PartsListClass.Hwind);
            state = true;    
            //Kill excel process 
            //but dont show open dialog box      
        }
    }
    return state;   //exit subroutine 
}   
*/
        private void saveCheckBox_CheckedChanged(object sender, EventArgs e)
        {   
            if (saveCheckBox.Checked)
            {
                if (String.IsNullOrWhiteSpace(varSave.varPath) != true)
                {
                    textBox1.Enabled = false;
                    browseButt.Enabled = false;
                    Properties.Settings.Default.DirPath = varSave.varPath;
                }
                else
                {
                    MessageBox.Show("No path selected");
                    textBox1.Enabled = true;
                    browseButt.Enabled = true;
                    saveCheckBox.Checked = false;
                }
            }
            else
            {
                textBox1.Enabled = true;
                browseButt.Enabled = true;
                Properties.Settings.Default.DirPath = null;
                //textBox1.Clear();
            }
            //save path location into memory and checked box status
            Properties.Settings.Default.checkStatus = saveCheckBox.Checked;            
            Properties.Settings.Default.Save();
        }

        ////////////////////////////////////////////////////////////////////////////////////////////
        //Search grid and various controls 
        ////////////////////////////////////////////////////////////////////////////////////////////

        /*
        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            //if Enter key has been pressed then perform 
            //search button action 
            if (e.KeyCode == Keys.Enter)
            {
                searchButt.PerformClick();
                e.SuppressKeyPress = true;
                e.Handled = true;
                return;
            }                           
        }
        */
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            int index = 0;
            int incIndex = 0;

            //Check if down key was pushed 
            if (e.KeyCode == Keys.Down)
            {
                //does datagrid contain any data?
                if (dataGridView1.RowCount > 0)
                {
                    //Get current value of the cell 
                    index = dataGridView1.CurrentCell.RowIndex;

                    //take a copy of the index and get next position
                    incIndex = index;   
                    ++incIndex;     

                    //Check that the next position is not greater than the number of 
                    //available rows 
                    if (incIndex != dataGridView1.RowCount)
                    {
                        //select and highlight the current position on the datagridview
                        dataGridView1.Rows[index].Selected = false;
                        dataGridView1.Rows[++index].Selected = true;

                        //Display items on GUI pass current position to function 
                        displayItemLocation(index);
                    }
                }

            }
            else if (e.KeyCode == Keys.Up)
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    //Get current value of the cell 
                    index = dataGridView1.CurrentCell.RowIndex;

                    //take a copy of the index and get previous position
                    incIndex = index;   
                    --incIndex;     

                    if (incIndex >= 0)
                    {
                        dataGridView1.Rows[index].Selected = false;
                        dataGridView1.Rows[--index].Selected = true;

                        slctedPart.Text = dataGridView1.Rows[index].Cells[0].Value.ToString();

                        //Display items on GUI
                        displayItemLocation(index);
                    }                 
                }

            }
            else if(e.KeyCode == Keys.Escape)
            {
                searchWordBox.Focus();
            }                        
            else
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
            }

        }
        
        private void displayItemLocation(int index)
        {
            //Display part information in box
            slctedPart.Text = dataGridView1.Rows[index].Cells[0].Value.ToString();

            //Display items on GUI
            cabinetBox.Text = dataGridView1.Rows[index].Cells[2].Value.ToString();
            rowBox.Text = dataGridView1.Rows[index].Cells[3].Value.ToString();
            drawerBox.Text = dataGridView1.Rows[index].Cells[4].Value.ToString();
            sectionBox.Text = dataGridView1.Rows[index].Cells[5].Value.ToString();
            quantity.Text = dataGridView1.Rows[index].Cells[9].Value.ToString();
        }
        
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridView1.Rows.Count <= 0)  //if no data is in cells
            {
                return;         
            }
            int index = dataGridView1.CurrentCell.RowIndex;
            
            slctedPart.Text = dataGridView1.Rows[index].Cells[0].Value.ToString();

            //enable stock quantity controls
            //quantityAddBut.Enabled = true;
            //quantityRemoveButt.Enabled = true;
            /*
            //Display items on GUI
            cabinetBox.Text = dataGridView1.Rows[index].Cells[2].Value.ToString();
            rowBox.Text = dataGridView1.Rows[index].Cells[3].Value.ToString();
            drawerBox.Text = dataGridView1.Rows[index].Cells[4].Value.ToString();
            sectionBox.Text = dataGridView1.Rows[index].Cells[5].Value.ToString();
            quantity.Text = dataGridView1.Rows[index].Cells[9].Value.ToString();
            */
            //Display items on GUI
            displayItemLocation(index);

            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show(MousePosition);
            }

        }
        /*
        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            DataTable excelTableSel = new DataTable();

           if(e.ClickedItem == editToolStripMenuItem)
           {
                try
                {
                    AppExcel.EditComponentValues(dataGridView1); //send selected values to grid
                }
                catch (Exception ex)
                {
                    DialogResult errorResult = MessageBox.Show("The following contains no quantity do you wish to add following to stock",
                        "Error!", MessageBoxButtons.YesNo);

                    if (errorResult == DialogResult.Yes)
                    {
                        //select datagrid to highlight the first line 
                        displayItemLocation(dataGridView1.CurrentRow.Index);
                    }
                    //enableEditModeButt.Enabled = false;
                }
            }
           if(e.ClickedItem == copyToolStripMenuItem)
           {
                DataObject dataObj = dataGridView1.GetClipboardContent();
                Clipboard.SetDataObject(dataObj, true);
                dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
           }
           if(e.ClickedItem == deleteToolStripMenuItem)
           {
                DialogResult errorResult = MessageBox.Show("Do you wish for following items to be deleted from spreadsheet? ",
                    "Delete Entries", MessageBoxButtons.OKCancel);   
                                
                if (errorResult == DialogResult.OK)
                {
                    AppExcel.DeleteExcelValues(dataGridView1);  //pass values to function 
                } 
            }
           if(e.ClickedItem == addToStockToolStripMenuItem)
            {
                DialogResult errorResult = MessageBox.Show("Do you wish to add the following entries to stock order list?",
                    "Delete Entries", MessageBoxButtons.OKCancel);

                if (errorResult == DialogResult.OK)
                {                   
                    stockAdd(dataGridView1);
                }

            }
        }
      
        */
        ////////////////////////////////////////////////////////////////////////////////////////////
        //This section deals with adding entries to spreadsheet for user functionality!
        ////////////////////////////////////////////////////////////////////////////////////////////    

            
        private void newEntryButAdd_Click(object sender, EventArgs e)
        {
            //int index = 0;
            int rowNum = (int)rowCountNum.Value;

            if (table.Columns.Count <= 0)
            {
                table.Columns.Add("Part");
                table.Columns.Add("Pack");
                table.Columns.Add("Cabinet");
                table.Columns.Add("Row");
                table.Columns.Add("Drawer");
                table.Columns.Add("Section");
                table.Columns.Add("Part Info");
                table.Columns.Add("Supplier");
                table.Columns.Add("Alt Parts");
                table.Columns.Add("Quantity");
                table.Columns.Add("Index");
            }

            //Add new row to spreadsheet
            for (int index = 0; index < rowNum; index++)
            {
                table.Rows.Add();       //add more rows
            }

            //create new table to display columns and rows on 
            //datagrid view
            NewCompDataGrid.DataSource = table;            
        }
 
        private void RemoveEntryButt_Click(object sender, EventArgs e)
        {
            if (NewCompDataGrid.Rows.Count > 0)
            {
                int tableIndex = table.Rows.Count - 1;

                //Check if number rules on table is greater than 0
                //and remove row entries one at a time 
                table.Rows.RemoveAt(tableIndex);                
            }
        }

        private void spreadSheetAdd_Click(object sender, EventArgs e)
        {
            try
            {
                /*
                //Pass new values on table to myexcel class to be 
                //wrriten and stored inside spreadsheet

                if (AppExcel.WriteExcelValues(NewCompDataGrid))
                {
                    //if sucessful clear values off spreadsheet
                    table.Clear();  //remove all data values from table
                    NewCompDataGrid.DataSource = null;
                }
                else
                {

                }
                */
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }            
        }   
     
        private void spreadsheetDatagridView1Inialise(DataGridView grid)
        {
 
            DataTable InitTable = new DataTable();

            //Create table for adding values to spreadsheet
            InitTable.Columns.Add("Part");
            InitTable.Columns.Add("Pack");
            InitTable.Columns.Add("Cabinet");
            InitTable.Columns.Add("Row");
            InitTable.Columns.Add("Drawer");
            InitTable.Columns.Add("Section");
            InitTable.Columns.Add("Part Info");
            InitTable.Columns.Add("Supplier");
            InitTable.Columns.Add("Alt Parts");
            InitTable.Columns.Add("Quantity", typeof(int));
            InitTable.Columns.Add("Index");

            grid.DataSource = InitTable;

            grid.AutoGenerateColumns = true;
            DataGridViewColumn column0 = grid.Columns[0];
            column0.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column0.Width = 150;
            DataGridViewColumn column1 = grid.Columns[1];
            column1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column1.Width = 100;
            DataGridViewColumn column2 = grid.Columns[2];
            column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column2.Width = 100;
            DataGridViewColumn column3 = grid.Columns[3];
            column3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column3.Width = 100;
            DataGridViewColumn column4 = grid.Columns[4];
            column4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column4.Width = 100;
            DataGridViewColumn column5 = grid.Columns[5];
            column5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column5.Width = 100;
            DataGridViewColumn column6 = grid.Columns[6];
            column6.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column6.Width = 100;
            DataGridViewColumn column7 = grid.Columns[7];
            column7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column7.Width = 100;
            DataGridViewColumn column8 = grid.Columns[8];
            column8.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column8.Width = 100;
            DataGridViewColumn column9 = grid.Columns[9];
            column9.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column9.Width = 70;
            DataGridViewColumn column10 = grid.Columns[10];
            column10.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column10.Width = 90;
            column10.Visible = false;

        }

        /////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////Component stock quantity controls 
        /////////////////////////////////////////////////////////////////////////////////////////////

        private void quantityAddBut_Click(object sender, EventArgs e)
        {            
            int amount = 0; 
            try
            {
                //check if quantity value is greater than 0
                if(quantityBox.Value > 0)
                {
                    //extract quantity data from grid and update value
                    amount = Convert.ToInt16(dataGridView1.SelectedRows[0].Cells[9].Value);
                    amount = Convert.ToInt16(quantityBox.Value) + amount;

                    //write amount to function
                    AppExcel.WriteExcelValues(dataGridView1, amount, (int)WriteModeType.QUANTITY);
                }
                else
                {
                    MessageBox.Show("Quantity Box contains no amount!", "Warning");
                }
                quantityBox.Value = 0;

                dataGridView1.Rows[0].Selected = true;
                //select datagrid to highlight the first line 
                dataGridView1.Rows[0].Cells[0].Value.ToString();

                //Display searched items in boxes and set the row hightlighter to zero    
                displayItemLocation(0);



                /*  //Check if a cell has been selected 
                  if (dataGridView1.SelectedRows.Count > 0)
                  {
                      if (quantityBox.Value > 0)
                      {
                          RemoveStockButt.Enabled = true;

                          //Read current quantity amount off Excel cell grid and add new quantity to it
                          amount = Convert.ToInt16(dataGridView1.SelectedRows[0].Cells[9].Value);
                          amount = Convert.ToInt16(quantityBox.Value) + amount;
                          //Add stock quantity amount to spreadsheet 
                          AppExcel.UpdateQuantityValues(dataGridView1, amount);                       
                      }
                      else
                      {
                          MessageBox.Show("Quantity Box contains no amount!", "Warning");
                      }
                      quantityBox.Value = 0;

                      //Display part information in box
                      slctedPart.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();

                      //Display items on GUI
                      cabinetBox.Text = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                      rowBox.Text = dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
                      drawerBox.Text = dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
                      sectionBox.Text = dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
                      quantity.Text = dataGridView1.SelectedRows[0].Cells[9].Value.ToString();
                      //displayItemLocation(0);
                  }
                  */
            }
            catch(Exception ex)
            {
                DialogResult errorResult = MessageBox.Show("The following contains no quantity do you wish to add following to stock", 
                    "Error!", MessageBoxButtons.YesNo);

                if(errorResult == DialogResult.Yes)
                {                    
                    amount = Convert.ToInt16(quantityBox.Value);
                    AppExcel.WriteExcelValues(dataGridView1, amount, (int)WriteModeType.QUANTITY);
                    quantityBox.Value = 0;

                    //select datagrid to highlight the first line 
                    displayItemLocation(dataGridView1.CurrentRow.Index);
                    
                }

            }            
        }

        private void quantityRemoveButt_Click(object sender, EventArgs e)
        {
            int amount;
            try
            {
                    if (quantityBox.Value > 0)
                    {
                        int quantity = Convert.ToInt16(dataGridView1.SelectedRows[0].Cells[9].Value);
                        //Read current quantity amount off Excel cell grid and add new quantity to it
                        amount = quantity -  Convert.ToInt16(quantityBox.Value);

                        if (amount < 0)
                        {
                            if (quantity == 0)
                            {





                            /*
                                DialogResult errorResult = MessageBox.Show("Quantity selected is 0 do you wish to add to stock list?",
                                    "Error!", MessageBoxButtons.YesNo);

                                if(errorResult == DialogResult.Yes)
                                {
                                    //Add parts stock list 
                                }
                                */
                            }
                            else
                            {
                                DialogResult errorResult = MessageBox.Show("Amount selected is more than available amount do you still wish to take remaining from stock?",
                                    "Error!", MessageBoxButtons.YesNo);

                                if (errorResult == DialogResult.Yes)
                                {
                                    amount = 0;
                                    //Add stock quantity amount to spreadsheet 
                                    //AppExcel.UpdateQuantityValues(dataGridView1, amount);
                                }
                            }
                        }
                        else if (amount == 0)
                        {
                            //AppExcel.UpdateQuantityValues(dataGridView1, amount);
                            DialogResult errorResult = MessageBox.Show("No more parts left in stock do you wish to add stock order list?", "Warning", MessageBoxButtons.YesNo);
                        }
                        else
                        {
                            //Add stock quantity amount to spreadsheet 
                           // AppExcel.UpdateQuantityValues(dataGridView1, amount);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Quantity Box contains no amount!", "Warning");
                    }
                    //select datagrid to highlight the first line 
                    //displayItemLocation(dataGridView1.CurrentRow.Index);
                    quantityBox.Value = 0;
                
            }
            catch (Exception ex)
            {
                //select datagrid to highlight the first line 
                //displayItemLocation(dataGridView1.CurrentRow.Index);
                quantityBox.Value = 0;
                MessageBox.Show("Invalid Quantity Selection", "Error");
            }


        }

        private void stockAdd(DataGridView stockListView1)
        {                          
            if (stockTable.Columns.Count <= 0)
            {
                stockTable.Columns.Add("Part");
                stockTable.Columns.Add("Pack");
                stockTable.Columns.Add("Cabinet");
                stockTable.Columns.Add("Row");
                stockTable.Columns.Add("Drawer");
                stockTable.Columns.Add("Section");
                stockTable.Columns.Add("Part Info");
                stockTable.Columns.Add("Supplier");
                stockTable.Columns.Add("Alt Parts");
                stockTable.Columns.Add("Quantity");
                stockTable.Columns.Add("Index");
            }

            for (int rowIndex = 0; rowIndex < stockListView1.SelectedRows.Count; rowIndex++)
            {
                stockTable.Rows.Add();
                for (int colIndex = 0; colIndex < stockListView1.Columns.Count; colIndex++)
                {
                    stockTable.Rows[rowIndex + rowStart][colIndex] = stockListView1.SelectedRows[rowIndex].Cells[colIndex].Value;                    
                }
                rowIndexCounter++; //increment row index to determine starting position

            }
            //change position for row start last row is index position is retained in count value
            rowStart = rowIndexCounter;
            stockListGridView.DataSource = stockTable;
        }

        private void stockListGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            RemoveStockButt.Enabled = true;
        }

        private void RemoveAllStockButt_Click(object sender, EventArgs e)
        {
            RemoveStockButt.Enabled = false;
            spreadsheetDatagridView1Inialise(stockListGridView);
        }
        //////////////////////////////////////////////////////////////
        //Form closing and termination controls
        //////////////////////////////////////////////////////////////

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Kill any Excel processes left hanging in the memory
            AppExcel.KillProcessByMainWindows(PartsListClass.Hwind);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult Result = MessageBox.Show("Do you want to quit program?", "Exit", MessageBoxButtons.YesNo);

            if (Result == DialogResult.Yes)
            {
                //kill any processes if still in memory!
                AppExcel.KillProcessByMainWindows(PartsListClass.Hwind);
                this.Close();
            }
        }

        private void saveEdit_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                //dataGridView1.Enabled = true;
                dataGridView1.ReadOnly = false;
            }
        }

        private void editModeEnable_CheckedChanged(object sender, EventArgs e)
        {





            /*
            if(editModeEnable.Checked)
            {
                
                //enable tab page
                tabPage2.Enabled = true;


                //enable editing options
                deleteToolStripMenuItem.Enabled = true;

                if (dataGridView1.Rows.Count > 0)   //free edit control when data is displayed on the grid
                {
                    editToolStripMenuItem.Enabled = true;
                    dataGridView1.ReadOnly = false;
                }
                else
                {
                    editToolStripMenuItem.Enabled = false;
                    dataGridView1.ReadOnly = true;
                }
            }
            else
            {
                ///enable tab page
                tabPage2.Enabled = false;

                //enable editing options
                deleteToolStripMenuItem.Enabled = false;
                editToolStripMenuItem.Enabled = false;
                dataGridView1.ReadOnly = true;
                editModeEnable.Checked = false;
            }
            */
        }

      
        private void locationComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            switch (locationComboBox.SelectedIndex)
            {
                case (int)ComboLocType.ALLOCATED :
                    searchWordBox.Enabled = true;
                    searchOptionBox.Enabled = true;
                    PartSearchComboBox1.Enabled = true;
                    searchWordBox.Clear();
                    break;
                case (int)ComboLocType.SPARE:   //if SPARE OPTION SELECTED
                    searchWordBox.Enabled = false;
                    searchOptionBox.Enabled = false;
                    PartSearchComboBox1.Enabled = false;
                    searchWordBox.Text = "SPARE";           //use this as keyword
                    PartSearchComboBox1.SelectedIndex = 1;  //search by contain
                    searchOptionBox.SelectedIndex = 0;      //search by part number
                    break;

                default:
                    break;
            }
           
        }

        private void locTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(locTypeComboBox.SelectedIndex)
            {
                case (int)ComboCabinetType.ALL:
                    //disable and clear all location ComboBox features 
                    comboCabinetBox.Enabled = false;
                    comboRowBox.Enabled = false;
                    comboDrawerBox.Enabled = false;
                    comboSectionBox.Enabled = false;
                    comboCabinetBox.SelectedIndex = -1;
                    comboRowBox.SelectedIndex = -1;
                    comboDrawerBox.SelectedIndex = -1;
                    comboSectionBox.SelectedIndex = -1;
                    break;

                case (int)ComboCabinetType.CABINET:
                    comboCabinetBox.Enabled = true;
                    comboRowBox.Enabled = false;
                    comboDrawerBox.Enabled = false;
                    comboSectionBox.Enabled = false;
                    break;
                case (int)ComboCabinetType.ROW:
                    comboCabinetBox.Enabled = true;
                    comboRowBox.Enabled = true;
                    comboDrawerBox.Enabled = false;
                    comboSectionBox.Enabled = false;
                    break;
                case (int)ComboCabinetType.DRAWER:
                    comboCabinetBox.Enabled = true;
                    comboRowBox.Enabled = true;
                    comboDrawerBox.Enabled = true;
                    comboSectionBox.Enabled = false;
                    break;
                case (int)ComboCabinetType.SECTION:
                    comboCabinetBox.Enabled = true;
                    comboRowBox.Enabled = true;
                    comboDrawerBox.Enabled = true;
                    comboSectionBox.Enabled = true;
                    break;

                default:
                    break;
            }
        }

        private void searchLocationButton_Click(object sender, EventArgs e)
        {
            int searchOption; //stores variable for choose part option

            //Will construct a string based on the options selected 
            string optionStr = null;

            //search by part number or part information?
            if (searchOptionBox.SelectedIndex == 0)
                searchOption = 0; //search by part number
            else
                searchOption = 6; //search by part information

            try
            {
                    
                switch (locTypeComboBox.SelectedIndex)
                {
                    case (int)ComboCabinetType.CABINET:
                        optionStr = comboCabinetBox.SelectedItem.ToString();
                        break;

                    case (int)ComboCabinetType.ROW:
                        optionStr = comboCabinetBox.SelectedItem.ToString() + "," + comboRowBox.SelectedItem.ToString();

                        break;
                    case (int)ComboCabinetType.DRAWER:
                        optionStr = comboCabinetBox.SelectedItem.ToString() + "," + comboRowBox.SelectedItem.ToString() + ","
                            + comboDrawerBox.SelectedItem.ToString();
                        break;
                    case (int)ComboCabinetType.SECTION:
                        optionStr = comboCabinetBox.SelectedItem.ToString() + "," + comboRowBox.SelectedItem.ToString() + ","
                            + comboDrawerBox.SelectedItem.ToString() + "," + comboSectionBox.SelectedItem.ToString();
                        break;

                    case (int)ComboCabinetType.ALL:
                        optionStr = "*"; // send special charcter if search is done all characters
                        break;

                    default:
                        break;
                }


                //send string and commands to read excel class
                dataGridView1.DataSource = AppExcel.ReadExcelValues(optionStr, searchWordBox.Text, searchOption, PartSearchComboBox1.SelectedIndex);

                if (dataGridView1.Rows.Count > 0)  //check if function returned a table that's not null
                {
                    //view data select line
                    dataGridView1.Rows[0].Selected = true;
                    //select datagrid to highlight the first line 
                    dataGridView1.Rows[0].Cells[0].Value.ToString();

                    //Display searched items in boxes and set the row hightlighter to zero    
                    displayItemLocation(0);

                    //Display Found Entries as a string
                    resultsFound.Text = dataGridView1.RowCount.ToString();

                    //change focus to grid display
                    dataGridView1.Focus();

                    //enable stock quantity controls 
                    quantityAddBut.Enabled = true;
                    quantityRemoveButt.Enabled = true;
                    quantityBox.Enabled = true;
                }
                else
                {
                    //Clear boxes if nothing is found
                    slctedPart.Clear();
                    cabinetBox.Clear();
                    rowBox.Clear();
                    drawerBox.Clear();
                    sectionBox.Clear();
                    quantity.Clear();
                    resultsFound.Text = "No Results Found";

                    //disable stock quantity controls
                    //enable stock quantity controls
                    quantityAddBut.Enabled = false;
                    quantityRemoveButt.Enabled = false;
                    quantityBox.Enabled = false;
                }
                   
            }
            catch (Exception ex)
            {
                MessageBox.Show("One or more of the boxes must not be left blank!");
            }
            
        }

        //When search enter button is pressed perform search action and supress event!
        private void searchWordBox_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)
            {
                searchLocationButton.PerformClick();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        
        private void newSearchButt_Click(object sender, EventArgs e)
        {
           
            //Initialise Excel spreadsheet for display purposes 
            //Initialise values on dataGrid
            //dataGridView1.DataSource = AppExcel.InitSpreadSheet();
            spreadsheetDatagridView1Inialise(dataGridView1);
            //dataGridView1.Enabled = false;

            labelMessage.Text = null;
            resultsFound.Text = null;
            //clear text boxes
            if (locationComboBox.SelectedIndex != 1)
                searchWordBox.Clear();

            slctedPart.Clear();
            cabinetBox.Clear();
            rowBox.Clear();
            drawerBox.Clear();
            sectionBox.Clear();
            quantity.Clear();

            //disable stock quantity controls
            quantityAddBut.Enabled = false;
            quantityRemoveButt.Enabled = false;
            quantityBox.Enabled = false;
        }

#endif
    }
}

  



          
 
            
            





        

