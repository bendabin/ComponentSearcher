using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ComponentSearch
{
    class PartList
    {
        public string Part { get; set; }
        public string Pack { get; set; }
        public string Cabinet { get; set; }
        public string Row { get; set; }
        public string Drawer { get; set; }
        public string Section { get; set; }        
        public string Parts_Info { get; set; }
        public string Supplier { get; set; }
    }

    class EmpConstants
    {
        private const string DOMAIN_NAME = "xyz.com";
    }
}
