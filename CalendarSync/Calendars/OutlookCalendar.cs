using System;
using System.Collections.Generic;
using System.Linq;
using CalendarSync.Data;
using Microsoft.Office.Interop.Outlook;

namespace CalendarSync.Calendars
{
    internal class OutlookCalendar
    {
        public static List<Appointment> GetAppointments(DateTime start, DateTime end)
        {
            var calendar = Globals.ThisAddIn.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            var items = calendar.Items;

            items.IncludeRecurrences = true;
            items.Sort("[Start]", Type.Missing);

            var filter = "[Start] >= '" + start.ToString("g") + "' and [Start] < '" + end.ToString("g") + "'";
            var appointments = items.Restrict(filter)
                .Cast<object>()
                .OfType<AppointmentItem>()
                .Select(i => new Appointment(i))
                .ToList();
            return appointments;
        }
    }
}