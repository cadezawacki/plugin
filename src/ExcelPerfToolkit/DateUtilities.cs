using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace ExcelPerfToolkit;

/// <summary>
/// Round 3 date utilities. <see cref="WorkdayAdd"/> advances date serials by a number of
/// working days with a configurable weekend mask and an optional holiday list - the flexible
/// form of Excel's <c>WORKDAY.INTL</c>, vectorized over whole ranges in one pass and
/// registered <c>IsThreadSafe = true</c>.
///
/// <para>Bottlenecks addressed: #1 and #2 (one bulk in, one bulk out, no COM re-entry).</para>
/// </summary>
public static class DateUtilities
{
    private const string DefaultWeekendMask = "0000011"; // Mon..Sun: Saturday and Sunday are weekend.

    /// <summary>
    /// Returns the date serial reached by advancing each start date by the given number of
    /// working days, skipping weekend days and holidays. <paramref name="days"/> may be
    /// negative to go backward; zero returns the start date's day. The weekend mask is seven
    /// characters Monday..Sunday where <c>'1'</c> marks a non-working day (default
    /// <c>"0000011"</c>). <paramref name="holidays"/> is an optional range of date serials to
    /// skip.
    ///
    /// <para><paramref name="startDates"/> and <paramref name="days"/> may each be a single
    /// value or a block; a single value is broadcast against a block, and two blocks must
    /// share the same shape. Non-numeric start dates yield <c>#VALUE!</c> in that cell.</para>
    /// Marshaling cost: O(1). Thread-safety: SAFE for MTR.
    /// </summary>
    [ExcelFunction(Name = "EPT.WORKDAYADD", Description = "Add working days to dates with a custom weekend mask and holiday list (vectorized WORKDAY.INTL).", Category = "EPT.Date", IsThreadSafe = true, IsVolatile = false)]
    public static object WorkdayAdd(
        [ExcelArgument(Name = "start_dates", Description = "A date serial or a block of date serials.")] object startDates,
        [ExcelArgument(Name = "days", Description = "Working days to add (negative goes backward); scalar or block.")] object days,
        [ExcelArgument(Name = "weekend_mask", Description = "7 chars Mon..Sun; '1' = non-working. Default \"0000011\".")] string weekendMask = DefaultWeekendMask,
        [ExcelArgument(Name = "holidays", Description = "Optional range of date serials to skip.")] object? holidays = null)
    {
        var weekend = ParseWeekendMask(weekendMask);
        var holidaySet = ParseHolidays(holidays);

        var starts = Marshaling.AsArray2D(startDates);
        var dayCounts = Marshaling.AsArray2D(days);
        var startScalar = starts.GetLength(0) == 1 && starts.GetLength(1) == 1;
        var daysScalar = dayCounts.GetLength(0) == 1 && dayCounts.GetLength(1) == 1;

        int rows, cols;
        if (startScalar && daysScalar)
        {
            return ComputeOne(starts[0, 0], dayCounts[0, 0], weekend, holidaySet);
        }
        if (startScalar)
        {
            rows = dayCounts.GetLength(0);
            cols = dayCounts.GetLength(1);
        }
        else if (daysScalar)
        {
            rows = starts.GetLength(0);
            cols = starts.GetLength(1);
        }
        else
        {
            rows = starts.GetLength(0);
            cols = starts.GetLength(1);
            if (dayCounts.GetLength(0) != rows || dayCounts.GetLength(1) != cols)
            {
                throw new ArgumentException("start_dates and days must have the same shape when both are ranges.");
            }
        }

        var result = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                var start = startScalar ? starts[0, 0] : starts[r, c];
                var dayCount = daysScalar ? dayCounts[0, 0] : dayCounts[r, c];
                result[r, c] = ComputeOne(start, dayCount, weekend, holidaySet);
            }
        }
        return result;
    }

    /// <summary>Excel's last representable date serial: 9999-12-31.</summary>
    private const int MaxSerial = 2_958_465;

    private static object ComputeOne(object? start, object? days, bool[] weekend, HashSet<int> holidays)
    {
        if (!Marshaling.TryToDouble(start, out var startSerial) || startSerial < 0d)
        {
            return ExcelError.ExcelErrorValue;
        }
        if (!Marshaling.TryToDouble(days, out var dayCountRaw))
        {
            return ExcelError.ExcelErrorValue;
        }

        DateTime date;
        try
        {
            date = DateTime.FromOADate(startSerial).Date;
        }
        catch (ArgumentException)
        {
            return ExcelError.ExcelErrorValue;
        }

        // Reject before the int cast: an out-of-int-range double converts to
        // int.MinValue, which would flip the walk direction and overflow Math.Abs.
        // No working-day count larger than the whole serial span can ever complete.
        var truncated = Math.Truncate(dayCountRaw);
        if (Math.Abs(truncated) > MaxSerial)
        {
            return ExcelError.ExcelErrorNum;
        }
        var remaining = (int)truncated;
        if (remaining == 0)
        {
            return date.ToOADate();
        }

        // Walk in pure int-serial space: DateTime.AddDays + DayOfWeek + ToOADate per
        // step cost ~3x more (measured 16 ns/step vs 5 ns) and the serial bounds double
        // as the year-1/9999 guards - a per-cell #NUM! instead of an exception that
        // would fail the whole block.
        var serial = (int)date.ToOADate();
        // Map DayOfWeek (Sunday=0..Saturday=6) onto the Monday-first mask index.
        var maskIndex = ((int)date.DayOfWeek + 6) % 7;
        var step = remaining > 0 ? 1 : -1;
        remaining = Math.Abs(remaining);
        var checkHolidays = holidays.Count > 0;
        while (remaining > 0)
        {
            serial += step;
            if (serial < 0 || serial > MaxSerial)
            {
                return ExcelError.ExcelErrorNum;
            }
            maskIndex += step;
            if (maskIndex == 7)
            {
                maskIndex = 0;
            }
            else if (maskIndex < 0)
            {
                maskIndex = 6;
            }
            if (!weekend[maskIndex] && !(checkHolidays && holidays.Contains(serial)))
            {
                remaining--;
            }
        }
        return (double)serial;
    }

    private static bool[] ParseWeekendMask(string mask)
    {
        if (string.IsNullOrEmpty(mask))
        {
            mask = DefaultWeekendMask;
        }
        if (mask.Length != 7)
        {
            throw new ArgumentException("weekend_mask must be exactly 7 characters (Mon..Sun).", nameof(mask));
        }
        var weekend = new bool[7];
        var working = 0;
        for (var i = 0; i < 7; i++)
        {
            switch (mask[i])
            {
                case '1':
                    weekend[i] = true;
                    break;
                case '0':
                    weekend[i] = false;
                    working++;
                    break;
                default:
                    throw new ArgumentException("weekend_mask may only contain '0' and '1'.", nameof(mask));
            }
        }
        if (working == 0)
        {
            throw new ArgumentException("weekend_mask marks every day as non-working.", nameof(mask));
        }
        return weekend;
    }

    private static HashSet<int> ParseHolidays(object? holidays)
    {
        var set = new HashSet<int>();
        if (Marshaling.IsBlankOrError(holidays))
        {
            return set;
        }
        var block = Marshaling.AsArray2D(holidays!);
        var rows = block.GetLength(0);
        var cols = block.GetLength(1);
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                if (Marshaling.TryToDouble(block[r, c], out var serial) && serial >= 0d)
                {
                    set.Add((int)Math.Floor(serial));
                }
            }
        }
        return set;
    }
}
