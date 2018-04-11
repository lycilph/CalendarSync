using System;

namespace CalendarSync
{
    public static class DateTimeExtensions
    {
        public static DateTime NowAtTime(int hour, int minute, int second)
        {
            var now = DateTime.Now;
            return new DateTime(now.Year, now.Month, now.Day, hour, minute, second);
        }

        public static DateTime ThisMorning()
        {
            return NowAtTime(8, 0, 0);
        }
    }
}
