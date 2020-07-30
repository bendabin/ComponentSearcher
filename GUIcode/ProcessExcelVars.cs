using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;

namespace ComponentSearcher
{
    /*
     *   This class will store any values and variables that are sent to it by MyExcel Class
     *   and will allow other classes such as Form1 class to access these variables for use 
     */
    class ProcessExcelVars
    {             
        //store Excel spreadsheet values inside a bindings list
        //Store the individual values in a binding list
        public string Part { get; set; }
        public string Pack { get; set; }
        public string Cabinet { get; set; }
        public string Row { get; set; }
        public string Drawer { get; set; }
        public string Section { get; set; }
        public string Parts_Info { get; set; }
        public string Supplier { get; set; }   
        public string AltParts { get; set; }
        public string Quantity { get; set; }
    }

    class PartsListClass
    {
        //store previous good working path
        public static string PrevDirPath = null;

        //intialise Hwnd value for main window to zero and store variables inside this class
        public static int Hwind = 0;

        //Store Excel values inside the following class 
        public static List<ProcessExcelVars> ExcelSorted = new List<ProcessExcelVars>();

        //store Excel table write values inside the following public static list inside class
        public static List<ProcessExcelVars> ExcelWrite = new List<ProcessExcelVars>();


    }

    //Creating a Custom ICompare to check empty string values by using default String.Compare
    //The first checks will return -1 instead of 1 or 1 instead of -1, if using the standard string comparison.
    public class EmptyStringsAreLast : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            if (String.IsNullOrEmpty(y) && !String.IsNullOrEmpty(x))
            {
                return -1;
            }
            else if (!String.IsNullOrEmpty(y) && String.IsNullOrEmpty(x))
            {
                return 1;
            }
            else
            {
                return String.Compare(x, y);
            }
        }
    }
}
