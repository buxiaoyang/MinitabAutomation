using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MinitabAutomation.Models
{
    class RowData
    {
        public string filePath { get; set; }
        public int rowCount { get; set; }
        public int[]  node { get; set; }
        public DateTime[] dataTime { get; set; }
        public ArrayList instances { get; set; }

    }
}
