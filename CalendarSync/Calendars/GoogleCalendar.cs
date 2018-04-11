using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.IO;
using System.Linq;
using System.Threading;
using CalendarSync.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Requests;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using NLog;
using Calendar = CalendarSync.Data.Calendar;
using Settings = CalendarSync.Data.Settings;

namespace CalendarSync.Calendars
{
    internal class GoogleCalendar
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private CalendarService service;
        public List<Calendar> Calendars { get; private set; }

        // http://stackoverflow.com/questions/27257590/check-if-user-is-already-logged-in
        public bool IsUserAuthorized()
        {
            logger.Trace("Checking user authorization status");

            var stream = new FileStream(Settings.SecretsFile, FileMode.Open, FileAccess.Read);
            var initializer = new GoogleAuthorizationCodeFlow.Initializer
            {
                ClientSecretsStream = stream,
                Scopes = new[] { CalendarService.Scope.Calendar },
                DataStore = new FileDataStore(Settings.DataDir, true)
            };
            var flow = new AuthorizationCodeFlow(initializer);
            var token_task = flow.LoadTokenAsync("user", CancellationToken.None);
            token_task.Wait();

            var token = token_task.Result;
            var is_user_authenticated = token != null && (token.RefreshToken != null || !token.IsExpired(flow.Clock));

            logger.Trace("User authorization status = " + (is_user_authenticated ? "authenticated" : "not authenticated"));

            return is_user_authenticated;
        }

        public void Initialize()
        {
            logger.Trace("Initializing service");
            var sw = Stopwatch.StartNew();

            if (!File.Exists(Settings.SecretsFile))
            {
                logger.Trace("Couldn't find client secrets file");
                throw new ApplicationException("Couldn't find client secrets file");
            }

            var stream = new FileStream(Settings.SecretsFile, FileMode.Open, FileAccess.Read);
            GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets,
                new[] { CalendarService.Scope.Calendar },
                "user", CancellationToken.None,
                new FileDataStore(Settings.DataDir, true))
                .ContinueWith(parent =>
                {
                    stream.Dispose();

                    service = new CalendarService(new BaseClientService.Initializer
                    {
                        HttpClientInitializer = parent.Result,
                        ApplicationName = Settings.ApplicationName,
                    });

                    Calendars =
                        service.CalendarList.List()
                            .Execute()
                            .Items.Select(i => new Calendar { DisplayName = i.Summary, Id = i.Id })
                            .ToList();

                    sw.Stop();
                    logger.Trace("Service initialization done - {0} ms", sw.ElapsedMilliseconds);
                }).Wait();
        }

        public Task<List<Appointment>> GetAppointments(DateTime start, DateTime end, string calendar_id)
        {
            return Task.Factory.StartNew(() =>
            {
                var request = service.Events.List(calendar_id);
                request.TimeMin = start;
                request.TimeMax = end;
                request.ShowDeleted = false;
                var events = request.Execute();

                var appointments = new List<Appointment>();
                if (events.Items.Count > 0)
                    appointments.AddRange(events.Items.Select(e => new Appointment(e)).OrderBy(e => e.Start).ToList());

                while (events.NextPageToken != null)
                {
                    Thread.Sleep(250);
                    request.PageToken = events.NextPageToken;
                    events = request.Execute();

                    if (events.Items.Count > 0)
                        appointments.AddRange(events.Items.Select(e => new Appointment(e)).OrderBy(e => e.Start).ToList());
                }

                return appointments;
            });
        }

        public void Add(IEnumerable<Appointment> items, IProgress<string> progress, string calendar_id)
        {
            var chunks = items.Chunk(50).ToList();
            foreach (var chunk in chunks)
            {
                var br = new BatchRequest(service);
                var appointments = chunk.ToList();

                foreach (var appointment in appointments)
                {
                    var google_event = appointment.ToGoogleEvent();
                    var request = service.Events.Insert(google_event, calendar_id);
                    br.Queue<Event>(request, (r, e, i, m) =>
                    {
                        if (!m.IsSuccessStatusCode)
                            logger.Error("Error: " + e.Message);
                    });

                    progress.Report("Adding " + appointment.Start.ToShortDateString() + " - " + appointment.End + " - " + appointment.Subject);
                }

                logger.Trace("Executing batch of {0} events to add", appointments.Count);
                br.ExecuteAsync().Wait();
                Thread.Sleep(250);
            }
            progress.Report(string.Format("Added {0} items", items.Count()));
            logger.Trace("Done adding items");
        }

        public void Remove(IEnumerable<Appointment> items, IProgress<string> progress, string calendar_id)
        {
            var chunks = items.Chunk(50).ToList();
            foreach (var chunk in chunks)
            {
                var br = new BatchRequest(service);
                var appointments = chunk.ToList();

                foreach (var appointment in appointments)
                {
                    var request = service.Events.Delete(calendar_id, appointment.OriginalId);
                    br.Queue<Event>(request, (r, e, i, m) =>
                    {
                        if (!m.IsSuccessStatusCode)
                            logger.Error("Error: " + e.Message);
                    });

                    progress.Report("Removing " + appointment.Start.ToShortDateString() + " - " + appointment.End + " - " + appointment.Subject);
                }

                logger.Trace("Executing batch of {0} events to remove", appointments.Count);
                br.ExecuteAsync().Wait();
                Thread.Sleep(250);
            }
            progress.Report(string.Format("Removed {0} items", items.Count()));
            logger.Trace("Done removing items");
        }
    }
}