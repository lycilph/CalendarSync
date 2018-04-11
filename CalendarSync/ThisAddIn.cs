using System;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CalendarSync.Calendars;
using CalendarSync.Data;
using NLog;
using Office = Microsoft.Office.Core;
using Timer = System.Timers.Timer;

namespace CalendarSync
{
    public partial class ThisAddIn
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly GoogleCalendar google_calendar = new GoogleCalendar();
        
        private Timer sync_timer;
        private WindowsFormsSynchronizationContext sync_context;

        private AddinStatus _Status = AddinStatus.Uninitialized;
        public AddinStatus Status
        {
            get { return _Status; }
            set
            {
                if (_Status == value) return;
                _Status = value;
                logger.Trace("Status changed: " + _Status);
            }
        }

        public DateTime? LastSync { get; set; }
        public DateTime NextSync { get; set; }

        public Settings Settings { get; private set; }
        public List<Calendar> GoogleCalendars => google_calendar.Calendars;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            logger.Trace("Starting {0} addin", Settings.ApplicationName);

            // Load settings here
            Settings = Settings.Load();

            // Initialize the google calendar silently after a specified delay
            logger.Trace("Current start up delay is {0} sec", Settings.StartupDelayInSeconds);
            Status = AddinStatus.WaitingToInitialize;
            Task.Delay(TimeSpan.FromSeconds(Settings.StartupDelayInSeconds))
                .ContinueWith(parent => InitializeGoogleCalendar(true));

            // Workaround, so that the sync call can be marshalled back to the main thread
            sync_context = new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(sync_context);

            sync_timer = new Timer(TimeSpan.FromMinutes(Settings.SyncIntervalInMinutes).TotalMilliseconds);
            sync_timer.Elapsed += OnSyncTimerElapsed;
            sync_timer.Start();

            UpdateNextSync();
        }

        private void OnSyncTimerElapsed(object sender, ElapsedEventArgs elapsed_event_args)
        {
            if (!Settings.IsSyncEnabled || Status != AddinStatus.Ready)
            {
                logger.Trace("Automatic syncing disabled");
                return;
            }

            logger.Trace("Automatic syncing " + DateTime.Now);

            var progress = new Progress<string>(s => logger.Trace(s));
            sync_context.Post(_ =>
            {
                SyncCalendars(progress, Settings.Calendar.Id)
                    .ContinueWith(parent => UpdateNextSync());
            }, null);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //       must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CalendarSyncRibbon();
        }

        public Task InitializeGoogleCalendar(bool silent)
        {
            logger.Trace("Initializing google calendar");

            if (Status != AddinStatus.WaitingToInitialize)
            {
                logger.Trace("Google calendar is already initialized");
                return Task.FromResult(0);
            }

            Status = AddinStatus.Initializing;
            return Task.Factory.StartNew(() =>
            {
                if (silent && google_calendar.IsUserAuthorized() || !silent)
                    google_calendar.Initialize();
            })
                .ContinueWith(
                    parent_task =>
                    {
                        Status = (google_calendar.IsUserAuthorized()
                            ? AddinStatus.Ready
                            : AddinStatus.WaitingToInitialize);
                    });
        }

        public Task CheckCalendars(IProgress<string> progress, string calendar_id)
        {
            var start = DateTimeExtensions.ThisMorning();
            var end = start.AddMonths(Settings.SyncWindowInMonths);
            var tcs = new TaskCompletionSource<bool>();

            progress.Report(string.Format("Checking calendars - window is {0} month(s)", Settings.SyncWindowInMonths));

            var outlook_appointments = OutlookCalendar.GetAppointments(start, end);
            progress.Report("----------------------------------");
            progress.Report(string.Format("Found {0} events in the outlook calendar", outlook_appointments.Count));

            google_calendar.GetAppointments(start, end, calendar_id)
                .ContinueWith(parent =>
                {
                    var google_appointments = parent.Result;
                    var items_to_remove =
                        google_appointments.Except(outlook_appointments, AppointmentComparer.Instance).ToList();
                    var items_to_add =
                        outlook_appointments.Except(google_appointments, AppointmentComparer.Instance).ToList();

                    progress.Report(string.Format("Found {0} events in the google calendar", parent.Result.Count));
                    progress.Report("----------------------------------");
                    progress.Report(string.Format("Found {0} items to remove", items_to_remove.Count));
                    items_to_remove.Apply(
                        a =>
                            progress.Report("Remove " + a.Start.ToShortDateString() + " - " + a.End.ToShortDateString() +
                                            ": " + a.Subject));
                    progress.Report("----------------------------------");
                    progress.Report(string.Format("Found {0} items to add", items_to_add.Count));
                    items_to_add.Apply(
                        a =>
                            progress.Report("Add " + a.Start.ToShortDateString() + " - " + a.End.ToShortDateString() +
                                            ": " + a.Subject));

                    tcs.SetResult(true);
                });

            return tcs.Task;
        }

        public Task SyncCalendars(IProgress<string> progress, string calendar_id)
        {
            var start = DateTimeExtensions.ThisMorning();
            var end = start.AddMonths(Settings.SyncWindowInMonths);
            var tcs = new TaskCompletionSource<bool>();

            progress.Report(string.Format("Syncing calendars - window is {0} month(s)", Settings.SyncWindowInMonths));
            LastSync = DateTime.Now;

            var outlook_appointments = OutlookCalendar.GetAppointments(start, end);
            progress.Report("----------------------------------");
            progress.Report(string.Format("Found {0} events in the outlook calendar", outlook_appointments.Count));

            google_calendar.GetAppointments(start, end, calendar_id)
                .ContinueWith(parent =>
                {
                    var google_appointments = parent.Result;
                    var items_to_remove =
                        google_appointments.Except(outlook_appointments, AppointmentComparer.Instance).ToList();
                    var items_to_add =
                        outlook_appointments.Except(google_appointments, AppointmentComparer.Instance).ToList();

                    progress.Report(string.Format("Found {0} events in the google calendar", parent.Result.Count));
                    progress.Report("----------------------------------");
                    progress.Report(string.Format("Found {0} items to remove", items_to_remove.Count));
                    google_calendar.Remove(items_to_remove, progress, google_calendar.Calendars[1].Id);
                    progress.Report("----------------------------------");
                    progress.Report(string.Format("Found {0} items to add", items_to_add.Count));
                    google_calendar.Add(items_to_add, progress, google_calendar.Calendars[1].Id);

                    tcs.SetResult(true);
                });

            return tcs.Task;
        }

        public Task ClearAll(IProgress<string> progress, string calendar_id)
        {
            // Delete the current window of x months (given by settings) and 1 extra before and after this
            var start = DateTimeExtensions.ThisMorning().AddMonths(-1);
            var end = start.AddMonths(Settings.SyncWindowInMonths + 2);

            progress.Report(string.Format("Clearing calendars - window is {0} month(s) +- 1 months", Settings.SyncWindowInMonths));

            return google_calendar.GetAppointments(start, end, calendar_id)
                .ContinueWith(parent => google_calendar.Remove(parent.Result, progress, calendar_id));
        }

        public void UpdateSyncInterval()
        {
            var interval_ms = TimeSpan.FromMinutes(Settings.SyncIntervalInMinutes).TotalMilliseconds;
            if (Math.Abs(interval_ms - sync_timer.Interval) < double.Epsilon)
                return;

            sync_timer.Interval = interval_ms;
            logger.Trace("Updating sync interval to {0} min ({1} ms)", Settings.SyncIntervalInMinutes, interval_ms);

            UpdateNextSync();
        }

        private void UpdateNextSync()
        {
            NextSync = DateTime.Now.AddMinutes(Settings.SyncIntervalInMinutes);
            logger.Trace("Next sync will happen at {0}", NextSync.ToLongTimeString());
        }

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
    }
}