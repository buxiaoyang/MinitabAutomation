using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MinitabAutomation.Models
{
    public class Instance
    {
        public string limType { get; set; }
        public Double upLimit { get; set; }
        public Double lowerLimit { get; set; }
        public string name { get; set; }
        public string unit { get; set; }
        public string title { get; set; }
        public ArrayList pictures { get; set; }
        public ArrayList data { get; set; }

        public Instance()
        {
            pictures = new ArrayList();
            data = new ArrayList();
            limType = "";
            upLimit = Double.NaN;
            lowerLimit = Double.NaN;
            name = "";
            unit = "";
            title = "";
        }
    }
}
