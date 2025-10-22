using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace QuanLy.Helpers
{
    public static class TuanHelper
    {
        /// <summary>
        /// Sinh dropdown tuần ISO cho nhiều năm.
        /// </summary>
        public static List<SelectListItem> GenerateWeekDropdown(int startYear, int endYear)
        {
            var items = new List<SelectListItem>();
            var cal = CultureInfo.CurrentCulture.Calendar;
            var weekRule = CalendarWeekRule.FirstFourDayWeek;
            var firstDay = DayOfWeek.Monday;

            for (int year = startYear; year <= endYear; year++)
            {
                int weeksInYear = cal.GetWeekOfYear(
                    new DateTime(year, 12, 31),
                    weekRule,
                    firstDay
                );

                for (int week = 1; week <= weeksInYear; week++)
                {
                    // Tìm ngày thứ năm của tuần 1
                    var jan1 = new DateTime(year, 1, 1);
                    int offset = DayOfWeek.Thursday - jan1.DayOfWeek;
                    var firstThursday = jan1.AddDays(offset);
                    int firstIsoWeek = cal.GetWeekOfYear(firstThursday, weekRule, firstDay);
                    int delta = week - (firstIsoWeek <= 1 ? 1 : 0);

                    // Tính ngày Monday của tuần hiện tại
                    var weekStart = firstThursday.AddDays(delta * 7).AddDays(-3);
                    if (weekStart.DayOfWeek != DayOfWeek.Monday)
                        weekStart = weekStart.AddDays((int)DayOfWeek.Monday - (int)weekStart.DayOfWeek);

                    string code = $"Y{year % 100:D2}W{week:D2}";
                    string label = $"{code} ({weekStart:dd/MM/yyyy} – {weekStart.AddDays(6):dd/MM/yyyy})";
                    items.Add(new SelectListItem(label, code));
                }
            }

            return items;
        }

        /// <summary>
        /// Từ code tuần ISO ("Y25W29") tính ra ngày bắt đầu (Monday) và kết thúc (Sunday).
        /// </summary>
        public static (DateTime Start, DateTime End) GetDateRangeFromCode(string code)
        {
            // code dạng "YyyWww", ví dụ "Y25W29"
            int year = 2000 + int.Parse(code.Substring(1, 2));
            int week = int.Parse(code.Substring(4, 2));

            var cal = CultureInfo.CurrentCulture.Calendar;
            var weekRule = CalendarWeekRule.FirstFourDayWeek;
            var firstDay = DayOfWeek.Monday;

            // Tìm Monday tuần 1
            var jan1 = new DateTime(year, 1, 1);
            int offset = DayOfWeek.Thursday - jan1.DayOfWeek;
            var firstThursday = jan1.AddDays(offset);
            int firstIsoWeek = cal.GetWeekOfYear(firstThursday, weekRule, firstDay);
            int delta = week - (firstIsoWeek <= 1 ? 1 : 0);

            var weekStart = firstThursday.AddDays(delta * 7).AddDays(-3);
            if (weekStart.DayOfWeek != DayOfWeek.Monday)
                weekStart = weekStart.AddDays((int)DayOfWeek.Monday - (int)weekStart.DayOfWeek);

            return (weekStart, weekStart.AddDays(6));
        }
    }
}
