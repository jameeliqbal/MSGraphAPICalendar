using MSGraphAPICalendar.User.Calendar.API.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace MSGraphAPICalendar.User.Calendar.API.Services
{
    public interface IOutlookCalendar
    {
        Task<List<OutlookCalendarEvent>> GetCalendarEvents();
        Task<OutlookCalendarEvent> AddEvent(OutlookCalendarEvent newEvent);
    }
}