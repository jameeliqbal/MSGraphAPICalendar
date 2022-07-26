using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MSGraphAPICalendar
{
    public class AppConfiguration
    {
        public IEnumerable<string> Scopes { get; set; }
        public string AppId { get; set; }
        public string NumberCalendarDays { get; set; }
    }
}
