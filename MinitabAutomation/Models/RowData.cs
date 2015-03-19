using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MinitabAutomation.Models
{
    public class RowData
    {
        public string filePath { get; set; }
        public int rowCount { get; set; }
        public ArrayList node { get; set; }
        public ArrayList dataTime { get; set; }
        public ArrayList instances { get; set; }

        public RowData()
        {
            node = new ArrayList();
            dataTime = new ArrayList();
            instances = new ArrayList();
        }
    }
}
