
using GoogleEvents.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace GoogleEvents
{
    public class CalendarContext : DbContext
    {
        public CalendarContext()
            : base()
        { }
        public DbSet<schedulerEvent> Events { get; set; }
    }
}