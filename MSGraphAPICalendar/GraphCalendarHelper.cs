using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TimeZoneConverter;

namespace MSGraphAPICalendar.GraphHelper
{
    public class GraphCalendarHelper
    {
        private static  GraphServiceClient graphClient;
        public static void Initialize(IAuthenticationProvider authProvider)
        {
            graphClient = new GraphServiceClient(authProvider);
        }

        public static async Task<Microsoft.Graph.User> GetMeAsync()
        {
            try
            {
                // GET /me
                return await graphClient.Me
                    .Request()
                    .Select(u => new {
                        u.DisplayName,
                        u.MailboxSettings
                    })
                    .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
                return null;
            }
        }

        public static async Task<IEnumerable<Event>> GetCalendarViewAsync(
            DateTime today,
            string timeZone,
            int numberOfDays)
        {
            if (numberOfDays > 31)
                return null;

            var startOfWeek = GetUtcStartOfWeekInTimeZone(today, timeZone);
            var endOfWeek = startOfWeek.AddDays(numberOfDays);

            var viewOptions = new List<QueryOption>
            {
                new QueryOption("startDateTime", startOfWeek.ToString("o")),
                new QueryOption("endDateTime", endOfWeek.ToString("o"))
            };

            try
            {
                var events = await graphClient.Me
                    .CalendarView
                    .Request(viewOptions)
                    .Header("Prefer", $"outlook.timezone=\"{timeZone}\"")
                    .Top(50)
                    .Select(e => new
                    {
                        e.Subject,
                        e.Organizer,
                        e.Start,
                        e.End
                    })
                    .OrderBy("start/dateTime")
                    .GetAsync();

                return events.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
                return null;
            }
        }


        public static async Task<IEnumerable<Event>> GetCurrentWeekCalendarViewAsync(DateTime today, string timeZone)
        {
            // Configure a calendar view for the current week
            var startOfWeek = today.AddDays(-1000);// GetUtcStartOfWeekInTimeZone(today, timeZone);
            var endOfWeek = today;// startOfWeek.AddDays(1030);

            var viewOptions = new List<QueryOption>
                    {
                        new QueryOption("startDateTime", startOfWeek.ToString("o")),
                        new QueryOption("endDateTime", endOfWeek.ToString("o"))
                    };

            try
            {
                var events = await graphClient.Me
                    .Calendar.CalendarView
                    .Request(viewOptions)
                    // Send user time zone in request so date/time in
                    // response will be in preferred time zone
                    .Header("Prefer", $"outlook.timezone=\"{timeZone}\"")
                    // Get max 50 per request
                    .Top(20)
                    // Only return fields app will use
                    .Select(e => new
                    {
                        e.Subject,
                        e.Organizer,
                        e.Start,
                        e.End
                    })
                    // Order results chronologically
                    .OrderBy("start/dateTime")
                    .GetAsync();

                return events.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
                return null;
            }
        }

        public static async Task<Event> AddEvent(Event newEvent)
        {
           return await graphClient.Me.Events
                      .Request()
                      .Header("Prefer", "outlook.timezone=\"Pacific Standard Time\"")
                      .AddAsync(newEvent);
        }

        private static DateTime GetUtcStartOfWeekInTimeZone(DateTime today,
            string timeZoneId)
        {
            TimeZoneInfo userTimeZone = TZConvert.GetTimeZoneInfo(timeZoneId);
            int diff = System.DayOfWeek.Sunday - today.DayOfWeek;
            var unspecifiedStart = DateTime.SpecifyKind(
                today.AddDays(diff), DateTimeKind.Unspecified);
            return TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, userTimeZone);
        }

    }
}
 
