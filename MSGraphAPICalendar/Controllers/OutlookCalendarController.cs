using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MSGraphAPICalendar.User.Calendar.API.Models;
using MSGraphAPICalendar.User.Calendar.API.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MSGraphAPICalendar.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class OutlookCalendarController : ControllerBase
    {
        private readonly IOutlookCalendar outlookCalendar;
        public OutlookCalendarController(IOutlookCalendar outlookCalendar)
        {
            this.outlookCalendar = outlookCalendar;
        }

        [HttpGet("GetCalendarEvents")]
        public async Task<IActionResult> GetCalendarEvents()
        {
            var calendarEvents = outlookCalendar.GetCalendarEvents();
            if (calendarEvents == null)
                return BadRequest("Calendar request failed");

            return Ok(calendarEvents);
        }

        [HttpPost("AddEvent")]
        public async Task<IActionResult> AddEvent(OutlookCalendarEvent model)
        {
            //model.Start.DateTime = DateTime.Now.AddDays(2).ToString();
            //model.Start.TimeZone = TimeZoneInfo.Local.StandardName;
            //model.End.DateTime = DateTime.Now.AddDays(4).ToString();
            //model.End.TimeZone = TimeZoneInfo.Local.StandardName;

            var calendarEvents = outlookCalendar.AddEvent(model);
            if (calendarEvents == null)
                return BadRequest("Calendar request failed");

            return Ok(calendarEvents);
        }
    }
}
