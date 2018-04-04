using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MockApi.Models
{
    public class LeaveNotifyData
    {
        public string To { get; set; }

        public IEnumerable<string> CC { get; set; }
    }
}
