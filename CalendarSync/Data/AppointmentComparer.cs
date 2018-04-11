using System.Collections.Generic;

namespace CalendarSync.Data
{
    public class AppointmentComparer : EqualityComparer<Appointment>
    {
        public static AppointmentComparer Instance = new AppointmentComparer();

        public override bool Equals(Appointment a1, Appointment a2)
        {
            return a1.Id == a2.Id && a1.Start == a2.Start && a1.End == a2.End;
        }

        public override int GetHashCode(Appointment obj)
        {
            return obj.Id.GetHashCode() ^ obj.Start.GetHashCode() ^ obj.End.GetHashCode();
        }
    }
}