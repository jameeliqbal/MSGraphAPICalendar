using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MSGraphAPICalendar.User.Calendar.API.Models;
using Microsoft.Extensions.Options;
 
using MSGraphAPICalendar.Graph.Authentication;
 

namespace MSGraphAPICalendar.User.Calendar.API.Services
{
    public class OutlookCalendar : IOutlookCalendar
    {
        private readonly IEnumerable<string> _scopes;
        private readonly string _appId;
        private readonly string _numberCalendarDays;
        private readonly IOptions<AppConfiguration> _appConfiguration;

        public OutlookCalendar(IOptions<AppConfiguration> appConfiguration)
        {
            _scopes = appConfiguration.Value.Scopes;
             
            _appId = appConfiguration.Value.AppId;
            _numberCalendarDays = appConfiguration.Value.NumberCalendarDays;
        }

        public async Task<List<OutlookCalendarEvent>> GetCalendarEvents()
        {
            var authProvider = new DeviceCodeAuthProvider(_appId, _scopes);

            var accessToken = authProvider.GetAccessToken().Result;
            Console.WriteLine($"Access token: {accessToken}\n");

            GraphHelper.GraphCalendarHelper.Initialize(authProvider);

            var user = GraphHelper.GraphCalendarHelper.GetMeAsync().Result;

            var events = ListCalendarEvents(
                user.MailboxSettings.TimeZone,
                $"{user.MailboxSettings.DateFormat} {user.MailboxSettings.TimeFormat}",
                System.Convert.ToInt16(_numberCalendarDays));

            List<OutlookCalendarEvent> outlookCalendarEvents = new
                List<OutlookCalendarEvent>();

            foreach (Microsoft.Graph.Event item in events)
            {
                outlookCalendarEvents.Add(new OutlookCalendarEvent()
                {
                    EventOrganizer = item.Organizer.EmailAddress.Name,
                    Subject = item.Subject,
                    Start = item.Start,
                    End = item.End
                });
            }
            return outlookCalendarEvents;
        }

        public async Task<OutlookCalendarEvent> AddEvent(OutlookCalendarEvent newEvent)
        {
            var authProvider = new DeviceCodeAuthProvider(_appId, _scopes);

            var accessToken = authProvider.GetAccessToken().Result;
            Console.WriteLine($"Access token: {accessToken}\n");

            GraphHelper.GraphCalendarHelper.Initialize(authProvider);

            var user = GraphHelper.GraphCalendarHelper.GetMeAsync().Result;

            var addedEvent = AddEventUsingGraph( newEvent);

             return new OutlookCalendarEvent()
                {
                    EventOrganizer = addedEvent.Organizer.EmailAddress.Name,
                    Subject = addedEvent.Subject,
                    Start = addedEvent.Start,
                    End = addedEvent.End
                } ;
            
             
        }

        //public async Task<string> AuthorizeAccessToOutlook()
        //{
        //    var authProvider = new DeviceCodeAuthProvider(_appId, _scopes);

        //    var result = authProvider.GetAccessToken().Result;
        //    if (result.StartsWith("Error"))
        //    {
        //        return result;
        //    }
        //}

        static List<Microsoft.Graph.Event> ListCalendarEvents(
            string userTimeZone,
            string dateTimeFormat,
            int numberOfDays)
        {
            var events = GraphHelper
                .GraphCalendarHelper
                .GetCurrentWeekCalendarViewAsync(
                    DateTime.Today,
                    userTimeZone)//,
                    //numberOfDays)
                .Result.ToList();
            return events;
        }

        static Microsoft.Graph.Event AddEventUsingGraph(OutlookCalendarEvent model)
        {
            var newEvent = new Microsoft.Graph.Event()
            {
                Subject = model.Subject,
                Start = model.Start,
                End = model.End,
                Organizer = new Microsoft.Graph.Recipient { EmailAddress = new Microsoft.Graph.EmailAddress() { Address = model.EventOrganizer } }
            };

            var addedEvent = GraphHelper
                .GraphCalendarHelper
                .AddEvent(newEvent)
                .Result;

            return addedEvent;
        }
    }
}
