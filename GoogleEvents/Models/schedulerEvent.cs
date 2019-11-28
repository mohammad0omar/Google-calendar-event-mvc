using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GoogleEvents.Models
{
    public class schedulerEvent
    {
        public int id { get; set; }
        public string text { get; set; }
        public DateTime start_date { get; set; }
        public DateTime end_date { get; set; }
        public string user { get; set; }
        public string gid { get; set; }
    }
}