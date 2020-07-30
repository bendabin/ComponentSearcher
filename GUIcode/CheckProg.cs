using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComponentSearcher
{
    class CheckProg
    {
        public void checkProgStatus(string dirStr)
        {

            try
            {
                System.IO.StreamReader file_name = new
                System.IO.StreamReader(dirStr);
            }
            catch (Exception ex)
            {
                int y = 0;
                //MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
            }
        }
    }
}
