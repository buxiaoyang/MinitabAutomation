using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MinitabAutomation.Models
{
    class Instance
    {
        public string limType { get; set; }
        public Double upLimit { get; set; }
        public Double lowerLimit { get; set; }
        public string name { get; set; }
        public string unit { get; set; }

        public Double[] data { get; set; }
    }
}
