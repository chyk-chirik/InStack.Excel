namespace InStack.Excel.Builder.Extensions;

/// <summary>
/// Taken from the framework source code
/// </summary>
internal class DateExtensions
{
    // Number of days in a non-leap year
    private const int DaysPerYear = 365;
    // Number of days in 4 years
    private const int DaysPer4Years = DaysPerYear * 4 + 1;       // 1461
    // Number of days in 100 years
    private const int DaysPer100Years = DaysPer4Years * 25 - 1;  // 36524
    // Number of days in 400 years
    private const int DaysPer400Years = DaysPer100Years * 4 + 1; // 146097

    // Number of days from 1/1/0001 to 12/30/1899
    private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

    private const long DoubleDateOffset = DaysTo1899 * TimeSpan.TicksPerDay;
    // The minimum OA date is 0100/01/01 (Note its year 100).
    // The maximum OA date is 9999/12/31
    private const long OaDateMinAsTicks = (DaysPer100Years - DaysPerYear) * TimeSpan.TicksPerDay;

    internal static double TicksToOaDate(long value)
    {
        switch (value)
        {
            case 0:
                return 0.0;  // Returns OleAut's zeroed date value.
            // This is a fix for VB. They want the default day to be 1/1/0001 rather than 12/30/1899.
            case < TimeSpan.TicksPerDay:
                value += DoubleDateOffset; // We could have moved this fix down, but we would like to keep the bounds check.
                break;
        }

        if (value < OaDateMinAsTicks)
            throw new OverflowException("Invalid OA date");
        // Currently, our max date == OA's max date (12/31/9999), so we don't
        // need an overflow check in that direction.
        var millis = (value - DoubleDateOffset) / TimeSpan.TicksPerMillisecond;
            
        if (millis < 0)
        {
            var frac = millis % TimeSpan.MillisecondsPerDay;
            if (frac != 0) millis -= (TimeSpan.MillisecondsPerDay + frac) * 2;
        }
        return (double)millis / TimeSpan.MillisecondsPerDay;
    }
}