using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Calendar.v3.Data;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CalendarSync.Data
{
    public class Appointment
    {
        private const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        private const string OutlookIdKey = "OutlookId";

        public string Id { get; set; }
        public string OriginalId { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Location { get; set; }
        public string Organizer { get; set; }
        public List<Contact> Recipients { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public DateTime LastModificationTime { get; set; }

        public Appointment() { }
        public Appointment(Outlook.AppointmentItem item)
        {
            Subject = item.Subject;
            Body = item.Body;
            Location = item.Location;
            Organizer = item.Organizer;
            Recipients = item.Recipients.OfType<Outlook.Recipient>().Select(r => new Contact {Name = r.Name, Email = r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS).ToString()}).ToList();
            Start = item.Start;
            End = item.End;
            LastModificationTime = item.LastModificationTime;
            Id = item.GlobalAppointmentID + "-" + Start.ToShortDateString() + "-" + End.ToShortDateString();
            OriginalId = Id;
        }
        public Appointment(Event item)
        {
            Id = item.ExtendedProperties.Shared[OutlookIdKey];
            OriginalId = item.Id;
            Subject = item.Summary;
            Body = item.Description;
            Location = item.Location;
            Organizer = item.Organizer.DisplayName;

            if (item.Attendees != null)
                Recipients = item.Attendees.Select(a => new Contact {Name = a.DisplayName, Email = a.Email}).ToList();

            if (item.Start.DateTime.HasValue)
                Start = item.Start.DateTime.Value;

            if (item.End.DateTime.HasValue)
                End = item.End.DateTime.Value;

            if (item.Updated != null)
                LastModificationTime = item.Updated.Value;
        }

        public Event ToGoogleEvent()
        {
            var e = new Event
            {
                Summary = Subject,
                Location = Location,
                Description = Body,
                Organizer = new Event.OrganizerData
                {
                    DisplayName = Organizer
                },
                Attendees = Recipients.Select(r => new EventAttendee {DisplayName = r.Name, Email = r.Email}).ToArray(),
                Start = new EventDateTime { DateTime = Start },
                End = new EventDateTime { DateTime = End },
                Updated = LastModificationTime,
                ExtendedProperties = new Event.ExtendedPropertiesData
                {
                    Shared = new Dictionary<string, string>
                    {
                        {OutlookIdKey, Id}
                    }
                }
            };

            return e;
        }
    }
}
