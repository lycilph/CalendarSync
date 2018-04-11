using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Threading.Tasks;
using System.Windows;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Input;
using NLog;
using CalendarSync.Data;

namespace CalendarSync
{
    public partial class CalendarSyncWindow
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public bool IsAddinReady
        {
            get { return (bool)GetValue(IsAddinReadyProperty); }
            set { SetValue(IsAddinReadyProperty, value); }
        }
        public static readonly DependencyProperty IsAddinReadyProperty =
            DependencyProperty.Register("IsAddinReady", typeof(bool), typeof(CalendarSyncWindow), new PropertyMetadata(false));

        public bool CanSync
        {
            get { return (bool)GetValue(CanSyncProperty); }
            set { SetValue(CanSyncProperty, value); }
        }
        public static readonly DependencyProperty CanSyncProperty =
            DependencyProperty.Register("CanSync", typeof(bool), typeof(CalendarSyncWindow), new PropertyMetadata(false));

        public bool IsAutomaticSyncEnabled
        {
            get { return (bool)GetValue(IsAutomaticSyncEnabledProperty); }
            set { SetValue(IsAutomaticSyncEnabledProperty, value); }
        }
        public static readonly DependencyProperty IsAutomaticSyncEnabledProperty =
            DependencyProperty.Register("IsAutomaticSyncEnabled", typeof(bool), typeof(CalendarSyncWindow), new PropertyMetadata(false));

        public int SyncInterval
        {
            get { return (int)GetValue(SyncIntervalProperty); }
            set { SetValue(SyncIntervalProperty, value); }
        }
        public static readonly DependencyProperty SyncIntervalProperty =
            DependencyProperty.Register("SyncInterval", typeof(int), typeof(CalendarSyncWindow), new PropertyMetadata(0, OnSyncIntervalChanged));

        private static void OnSyncIntervalChanged(DependencyObject obj, DependencyPropertyChangedEventArgs args)
        {
            var win = obj as CalendarSyncWindow;
            if (win == null) return;

            var interval = (int) args.NewValue;
            if (interval < Settings.MinSyncIntervalInMinutes)
            {
                interval = Settings.MinSyncIntervalInMinutes;
                win.SyncInterval = interval;
            }
            else
            {
                Globals.ThisAddIn.Settings.SyncIntervalInMinutes = interval;
                Globals.ThisAddIn.Settings.Save();
            }
        }

        public int SyncWindow
        {
            get { return (int)GetValue(SyncWindowProperty); }
            set { SetValue(SyncWindowProperty, value); }
        }
        public static readonly DependencyProperty SyncWindowProperty =
            DependencyProperty.Register("SyncWindow", typeof(int), typeof(CalendarSyncWindow), new PropertyMetadata(0, OnSyncWindowChangedCallback));

        private static void OnSyncWindowChangedCallback(DependencyObject obj, DependencyPropertyChangedEventArgs args)
        {
            var win = obj as CalendarSyncWindow;
            if (win == null) return;

            Globals.ThisAddIn.Settings.SyncWindowInMonths = (int)args.NewValue;
            Globals.ThisAddIn.Settings.Save();
        }

        public ObservableCollection<string> Messages
        {
            get { return (ObservableCollection<string>)GetValue(MessagesProperty); }
            set { SetValue(MessagesProperty, value); }
        }
        public static readonly DependencyProperty MessagesProperty =
            DependencyProperty.Register("Messages", typeof(ObservableCollection<string>), typeof(CalendarSyncWindow), new PropertyMetadata(null));

        public ObservableCollection<Calendar> CalendarList
        {
            get { return (ObservableCollection<Calendar>)GetValue(CalendarListProperty); }
            set { SetValue(CalendarListProperty, value); }
        }
        public static readonly DependencyProperty CalendarListProperty =
            DependencyProperty.Register("CalendarList", typeof(ObservableCollection<Calendar>), typeof(CalendarSyncWindow), new PropertyMetadata(null));

        public Calendar Calendar
        {
            get { return (Calendar)GetValue(CalendarProperty); }
            set { SetValue(CalendarProperty, value); }
        }
        public static readonly DependencyProperty CalendarProperty =
            DependencyProperty.Register("Calendar", typeof(Calendar), typeof(CalendarSyncWindow), new PropertyMetadata(null, OnSelectedCalendarChanged));

        private static void OnSelectedCalendarChanged(DependencyObject obj, DependencyPropertyChangedEventArgs args)
        {
            var win = obj as CalendarSyncWindow;
            if (win == null) return;

            var calendar = args.NewValue as Calendar;
            if (calendar == null) return;

            win.CanSync = true;
            Globals.ThisAddIn.Settings.Calendar = calendar;
            Globals.ThisAddIn.Settings.Save();
        }

        public string LastSync
        {
            get { return (string)GetValue(LastSyncProperty); }
            set { SetValue(LastSyncProperty, value); }
        }
        public static readonly DependencyProperty LastSyncProperty =
            DependencyProperty.Register("LastSync", typeof(string), typeof(CalendarSyncWindow), new PropertyMetadata(string.Empty));

        public string NextSync
        {
            get { return (string)GetValue(NextSyncProperty); }
            set { SetValue(NextSyncProperty, value); }
        }
        public static readonly DependencyProperty NextSyncProperty =
            DependencyProperty.Register("NextSync", typeof(string), typeof(CalendarSyncWindow), new PropertyMetadata(string.Empty));

        public CalendarSyncWindow()
        {
            logger.Trace("CalendarSyncWindow created");

            InitializeComponent();
            DataContext = this;

            Messages = new ObservableCollection<string>();
            Messages.CollectionChanged += MessagesChanged;
        }

        private void MessagesChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            var count = message_list.Items.Count;
            if (count < 2) { return; }

            var item = message_list.Items[count - 2];
            var framework = message_list.ItemContainerGenerator.ContainerFromItem(item) as FrameworkElement;

            framework?.BringIntoView();
        }

        //private void ChangeUser(object sender, RoutedEventArgs e)
        //{
        // Delete current user
        //Globals.ThisAddIn.GoogleCalendar.DeleteUserFile();

        // Reinitialize new user
        //Task.Factory.StartNew(() =>
        //{
        //    Globals.ThisAddIn.GoogleCalendar.Initialize();
        //})
        //    .ContinueWith(parent =>
        //    {
        //        CalendarList = new ObservableCollection<Calendar>(Globals.ThisAddIn.GoogleCalendar.Calendars);
        //        Calendar = CalendarList.FirstOrDefault();
        //    }, CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.FromCurrentSynchronizationContext());
        //}

        private void CalendarSyncWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            logger.Trace("CalendarSyncWindow loaded");

            // Save sync state and disable automatic syncing while window is open
            IsAutomaticSyncEnabled = Globals.ThisAddIn.Settings.IsSyncEnabled;
            Globals.ThisAddIn.Settings.IsSyncEnabled = false;

            if (Globals.ThisAddIn.Status == AddinStatus.Ready)
            {
                logger.Trace("Addin ready");
                Initialize();
            }
            else
            {
                logger.Trace("Addin not ready, initializing");
                Messages.Add("Initializing google calendar");

                Globals.ThisAddIn
                       .InitializeGoogleCalendar(false)
                       .ContinueWith(parent => Initialize(), TaskScheduler.FromCurrentSynchronizationContext());
            }
            
            if (Globals.ThisAddIn.Settings.Calendar == null)
                TabControl.SelectedIndex = 1; // Ie. go to options page if no calendar has been selected
        }

        private void CalendarSyncWindow_OnUnloaded(object sender, RoutedEventArgs e)
        {
            logger.Trace("CalendarSyncWindow unloaded");

            // Update sync interval
            Globals.ThisAddIn.UpdateSyncInterval();

            // Reset automatic syncing (was disabled when window was opened)
            Globals.ThisAddIn.Settings.IsSyncEnabled = IsAutomaticSyncEnabled;
            Globals.ThisAddIn.Settings.Save();
        }

        private void Initialize()
        {
            IsAddinReady = true;
            CalendarList = new ObservableCollection<Calendar>(Globals.ThisAddIn.GoogleCalendars);
            Calendar = (Globals.ThisAddIn.Settings.Calendar != null ? CalendarList.SingleOrDefault(c => Globals.ThisAddIn.Settings.Calendar.Id == c.Id) : null);
            CanSync = Globals.ThisAddIn.Settings.Calendar != null;
            SyncInterval = Globals.ThisAddIn.Settings.SyncIntervalInMinutes;
            SyncWindow = Globals.ThisAddIn.Settings.SyncWindowInMonths;
            LastSync = (Globals.ThisAddIn.LastSync.HasValue ? Globals.ThisAddIn.LastSync.Value.ToLongTimeString() : "NA");
            NextSync = Globals.ThisAddIn.NextSync.ToLongTimeString();
        }

        private void OnSyncIntervalPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private void OnSyncWindowPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private static bool IsTextAllowed(string text)
        {
            var regex = new Regex("[^0-9.-]+"); //regex that matches disallowed text
            return !regex.IsMatch(text);
        }

        private void OnSyncClick(object sender, RoutedEventArgs e)
        {
            var progress = new Progress<string>(Messages.Add);
            Messages.Clear();

            IsAddinReady = false;

            Task.Delay(100)
                .ContinueWith(_ => Globals.ThisAddIn.SyncCalendars(progress, Globals.ThisAddIn.Settings.Calendar.Id)).Unwrap()
                .ContinueWith(_ => IsAddinReady = true, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void OnCheckClick(object sender, RoutedEventArgs e)
        {
            var progress = new Progress<string>(Messages.Add);
            Messages.Clear();

            IsAddinReady = false;

            Task.Delay(100)
                .ContinueWith(_ => Globals.ThisAddIn.CheckCalendars(progress, Globals.ThisAddIn.Settings.Calendar.Id)).Unwrap()
                .ContinueWith(_ => IsAddinReady = true, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void OnClearAllClick(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to clear?", "Warning", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                return;

            var progress = new Progress<string>(Messages.Add);
            Messages.Clear();

            TabControl.SelectedIndex = 0;
            IsAddinReady = false;

            Task.Delay(100)
                .ContinueWith(_ => Globals.ThisAddIn.ClearAll(progress, Globals.ThisAddIn.Settings.Calendar.Id)).Unwrap()
                .ContinueWith(_ => IsAddinReady = true, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
}
