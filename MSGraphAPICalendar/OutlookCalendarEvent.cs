using Microsoft.Graph;

namespace MSGraphAPICalendar.User.Calendar.API.Models
{
    public class OutlookCalendarEvent
    {
        public string Subject { get; set; }
        public string EventOrganizer { get; set; }
        public DateTimeTimeZone Start { get; set; }
        public DateTimeTimeZone End { get; set; }
    }
}