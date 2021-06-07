using PianoFerieBuilder.Helpers;
using System;
using System.Linq;
using System.Collections.Generic;

namespace PianoFerieBuilder.Models
{
    public class Calendar
    {
        public List<CalendarDay> Days { get; }
 
        public Calendar(int year)
        {
            Days = new List<CalendarDay>();
            List<DateTime>  holidays = GetItalianHolidays(year);

            DateTime currentDate = new DateTime(year, 1, 1);
            DateTime stopDate = new DateTime(year, 12, 31);

            while (currentDate <= stopDate)
            {
                CalendarDay entity = new CalendarDay();
                entity.Date = currentDate;
                entity.IsWeekend = currentDate.DayOfWeek == DayOfWeek.Sunday || currentDate.DayOfWeek == DayOfWeek.Saturday;
                entity.IsHoliday = holidays.Contains(currentDate);

                Days.Add(entity);

                currentDate = currentDate.AddDays(1.0);
            }
        }

        public int PayeableDays
        {
            get
            {
                return Days.Where(v => (!v.IsWeekend && !v.IsHoliday) || (v.IsWeekend && v.IsHoliday)).Count();
            }
        }

        private List<DateTime> GetItalianHolidays(int year)
        {
            List<DateTime> dates = new List<DateTime>();

            // Capodanno
            dates.Add(new DateTime(year, 1, 1));
            // Epifania
            dates.Add(new DateTime(year, 1, 6));
            // Liberazione
            dates.Add(new DateTime(year, 4, 25));
            // Primo Maggio
            dates.Add(new DateTime(year, 5, 1));
            // Festa Repubblica
            dates.Add(new DateTime(year, 6, 2));
            // Ferragosto
            dates.Add(new DateTime(year, 8, 15));
            // Morti
            dates.Add(new DateTime(year, 11, 1));
            // Immacolata
            dates.Add(new DateTime(year, 12, 8));
            // Natale
            dates.Add(new DateTime(year, 12, 25));
            // Santo Stefano
            dates.Add(new DateTime(year, 12, 26));

            dates.AddRange(GetEasterSundayAndMonday(year));
            dates.AddRange(GetLocalHolidays(year));

            return dates;
        }

        private List<DateTime> GetLocalHolidays(int year)
        {
            List<DateTime> dates = new List<DateTime>();
            // San Giusto
            dates.Add(new DateTime(year, 11, 3));
            return dates;
        }

        private List<DateTime> GetEasterSundayAndMonday(int year)
        {
            List<DateTime> dates = new List<DateTime>();
            DateTime easter = EasterCalculator.GetEasterSunday(year);
            dates.Add(easter);
            dates.Add(easter.AddDays(1));
            return dates;
        }
    }
}
