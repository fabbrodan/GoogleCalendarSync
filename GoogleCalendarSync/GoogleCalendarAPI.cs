using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleCalendarSync
{
    public class GoogleCalendarAPI
    {
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "GoogleCalendatSync";
        CalendarService service;

        public GoogleCalendarAPI()
        {
            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                credPath += @"\.credentials\sync.json";

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });
        }

        public void NewAppointment(DateTime start, DateTime end, string subject)
        {
            Event @event = new Event();
            @event.Start = new EventDateTime();
            @event.End = new EventDateTime();
            
            @event.Start.DateTime = start;
            @event.End.DateTime = end;
            @event.Summary = subject;

            EventsResource.InsertRequest createRequest = service.Events.Insert(@event, "primary");
        }
    }
}
