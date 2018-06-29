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

            if (!Directory.Exists(Directory.GetCurrentDirectory() + "\\SyncSecret"))
            {
                Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\SyncSecret");
            }
            if (!File.Exists(Directory.GetCurrentDirectory() + "\\SyncSecret\\client_secret.json"))
            {
                File.Copy(AppDomain.CurrentDomain.BaseDirectory + "\\client_secret.json", Directory.GetCurrentDirectory() + "\\SyncSecret\\client_secret.json");
            }

            string secretPath = Directory.GetCurrentDirectory() + "\\SyncSecret\\client_secret.json";
            using (var stream = new FileStream(secretPath, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\.credentials\sync.json";

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

        public void NewAppointment(Event _event)
        {
            EventsResource.InsertRequest createRequest = service.Events.Insert(_event, "primary");
            createRequest.Execute();   
        }

        public void UpdateAppointment(string id, Event _event)
        { 
            EventsResource.UpdateRequest updateRequest = service.Events.Update(_event, "primary", id);
            updateRequest.Execute();
        }

        public void DeleteAppointment(string id)
        {
            EventsResource.DeleteRequest deleteRequest = service.Events.Delete("primary", id);
            try
            {
                deleteRequest.Execute();
            }
            catch
            {
                
            }

        }
    }
}
