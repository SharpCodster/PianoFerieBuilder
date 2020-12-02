using System;

namespace PianoFerieBuilder.Models
{
    public class CalendarDay
    {
        public DateTime Date { get; set; }
        public bool IsHoliday { get; set; }
        public bool IsWeekend { get; set; }
    }
}
