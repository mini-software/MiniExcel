using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace MiniExcelNumberFormat
{
    static internal class Formatter
    {
        static public string Format(object value, string formatString, CultureInfo culture, bool isDate1904)
        {
            var format = new NumberFormat(formatString);
            if (!format.IsValid)
                return CompatibleConvert.ToString(value, culture);

            var section = Evaluator.GetSection(format.Sections, value);
            if (section == null)
                return CompatibleConvert.ToString(value, culture);

            return Format(value, section, culture, isDate1904);
        }

        static public string Format(object value, Section node, CultureInfo culture, bool isDate1904)
        {
            switch (node.Type)
            {
                case SectionType.Number:
                    // Hide sign under certain conditions and section index
                    var number = Convert.ToDouble(value, culture);
                    if ((node.SectionIndex == 0 && node.Condition != null) || node.SectionIndex == 1)
                        number = Math.Abs(number);

                    return FormatNumber(number, node.Number, culture);

                case SectionType.Date:
                    if (ExcelDateTime.TryConvert(value, isDate1904, culture, out var excelDateTime))
                    {
                        return FormatDate(excelDateTime, node.GeneralTextDateDurationParts, culture);
                    }
                    else
                    {
                        throw new FormatException("Unexpected date value");
                    }

                case SectionType.Duration:
                    if (value is TimeSpan ts)
                    {
                        return FormatTimeSpan(ts, node.GeneralTextDateDurationParts, culture);
                    }
                    else
                    {
                        var d = Convert.ToDouble(value);
                        return FormatTimeSpan(TimeSpan.FromDays(d), node.GeneralTextDateDurationParts, culture);
                    }

                case SectionType.General:
                case SectionType.Text:
                    return FormatGeneralText(CompatibleConvert.ToString(value, culture), node.GeneralTextDateDurationParts);

                case SectionType.Exponential:
                    return FormatExponential(Convert.ToDouble(value, culture), node, culture);

                case SectionType.Fraction:
                    return FormatFraction(Convert.ToDouble(value, culture), node, culture);

                default:
                    throw new InvalidOperationException("Unknown number format section");
            }
        }

        static string FormatGeneralText(string text, List<string> tokens)
        {
            var result = new StringBuilder();
            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (Token.IsGeneral(token) || token == "@")
                {
                    result.Append(text);
                }
                else
                {
                    FormatLiteral(token, result);
                }
            }
            return result.ToString();
        }

        private static string FormatTimeSpan(TimeSpan timeSpan, List<string> tokens, CultureInfo culture)
        {
            // NOTE/TODO: assumes there is exactly one [hh], [mm] or [ss] using the integer part of TimeSpan.TotalXXX when formatting.
            // The timeSpan input is then truncated to the remainder fraction, which is used to format mm and/or ss.
            var result = new StringBuilder();
            var containsMilliseconds = false;
            for (var i = tokens.Count - 1; i >= 0; i--)
            {
                if (tokens[i].StartsWith(".0"))
                {
                    containsMilliseconds = true;
                    break;
                }
            }

            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];

                if (token.StartsWith("m", StringComparison.OrdinalIgnoreCase))
                {
                    var value = timeSpan.Minutes;
                    var digits = token.Length;
                    result.Append(value.ToString("D" + digits));
                }
                else if (token.StartsWith("s", StringComparison.OrdinalIgnoreCase))
                {
                    // If format does not include ms, then include ms in seconds and round before printing
                    var formatMs = containsMilliseconds ? 0 : timeSpan.Milliseconds / 1000D;
                    var value = (int)Math.Round(timeSpan.Seconds + formatMs, 0, MidpointRounding.AwayFromZero);
                    var digits = token.Length;
                    result.Append(value.ToString("D" + digits));
                }
                else if (token.StartsWith("[h", StringComparison.OrdinalIgnoreCase))
                {
                    var value = (int)timeSpan.TotalHours;
                    var digits = token.Length - 2;
                    result.Append(value.ToString("D" + digits));
                    timeSpan = new TimeSpan(0, 0, Math.Abs(timeSpan.Minutes), Math.Abs(timeSpan.Seconds), Math.Abs(timeSpan.Milliseconds));
                }
                else if (token.StartsWith("[m", StringComparison.OrdinalIgnoreCase))
                {
                    var value = (int)timeSpan.TotalMinutes;
                    var digits = token.Length - 2;
                    result.Append(value.ToString("D" + digits));
                    timeSpan = new TimeSpan(0, 0, 0, Math.Abs(timeSpan.Seconds), Math.Abs(timeSpan.Milliseconds));
                }
                else if (token.StartsWith("[s", StringComparison.OrdinalIgnoreCase))
                {
                    var value = (int)timeSpan.TotalSeconds;
                    var digits = token.Length - 2;
                    result.Append(value.ToString("D" + digits));
                    timeSpan = new TimeSpan(0, 0, 0, 0, Math.Abs(timeSpan.Milliseconds));
                }
                else if (token.StartsWith(".0")) {
                    var value = timeSpan.Milliseconds;
                    var digits = token.Length - 1;
                    result.Append("." + value.ToString("D" + digits));
                }
                else
                {
                    FormatLiteral(token, result);
                }
            }

            return result.ToString();
        }

        private static string FormatDate(ExcelDateTime date, List<string> tokens, CultureInfo culture)
        {
            var containsAmPm = ContainsAmPm(tokens);

            var result = new StringBuilder();
            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];

                if (token.StartsWith("y", StringComparison.OrdinalIgnoreCase))
                {
                    // year
                    var digits = token.Length;
                    if (digits < 2)
                        digits = 2;
                    if (digits == 3)
                        digits = 4;

                    var year = date.Year;
                    if (digits == 2)
                        year = year % 100;

                    result.Append(year.ToString("D" + digits));
                }
                else if (token.StartsWith("m", StringComparison.OrdinalIgnoreCase))
                {
                    // If  "m" or "mm" code is used immediately after the "h" or "hh" code (for hours) or immediately before 
                    // the "ss" code (for seconds), the application shall display minutes instead of the month. 
                    if (LookBackDatePart(tokens, i - 1, "h") || LookAheadDatePart(tokens, i + 1, "s"))
                    {
                        var digits = token.Length;
                        result.Append(date.Minute.ToString("D" + digits));
                    }
                    else
                    {
                        var digits = token.Length;
                        if (digits == 3)
                        {
                            result.Append(culture.DateTimeFormat.AbbreviatedMonthNames[date.Month - 1]);
                        }
                        else if (digits == 4)
                        {
                            result.Append(culture.DateTimeFormat.MonthNames[date.Month - 1]);
                        }
                        else if (digits == 5)
                        {
                            result.Append(culture.DateTimeFormat.MonthNames[date.Month - 1][0]);
                        }
                        else
                        {
                            result.Append(date.Month.ToString("D" + digits));
                        }
                    }
                }
                else if (token.StartsWith("d", StringComparison.OrdinalIgnoreCase))
                {
                    var digits = token.Length;
                    if (digits == 3)
                    {
                        // Sun-Sat
                        result.Append(culture.DateTimeFormat.AbbreviatedDayNames[(int)date.DayOfWeek]);
                    }
                    else if (digits == 4)
                    {
                        // Sunday-Saturday
                        result.Append(culture.DateTimeFormat.DayNames[(int)date.DayOfWeek]);
                    }
                    else
                    {
                        result.Append(date.Day.ToString("D" + digits));
                    }
                }
                else if (token.StartsWith("h", StringComparison.OrdinalIgnoreCase))
                {
                    var digits = token.Length;
                    if (containsAmPm)
                        result.Append(((date.Hour + 11) % 12 + 1).ToString("D" + digits));
                    else
                        result.Append(date.Hour.ToString("D" + digits));
                }
                else if (token.StartsWith("s", StringComparison.OrdinalIgnoreCase))
                {
                    var digits = token.Length;
                    result.Append(date.Second.ToString("D" + digits));
                }
                else if (token.StartsWith("g", StringComparison.OrdinalIgnoreCase))
                {
                    var era = culture.DateTimeFormat.Calendar.GetEra(date.AdjustedDateTime);
                    var digits = token.Length;
                    if (digits < 3)
                    {
                        result.Append(culture.DateTimeFormat.GetAbbreviatedEraName(era));
                    }
                    else
                    {
                        result.Append(culture.DateTimeFormat.GetEraName(era));
                    }
                }
                else if (string.Compare(token, "am/pm", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    var ampm = date.ToString("tt", CultureInfo.InvariantCulture);
                    result.Append(ampm.ToUpperInvariant());
                }
                else if (string.Compare(token, "a/p", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    var ampm = date.ToString("%t", CultureInfo.InvariantCulture);
                    if (char.IsUpper(token[0]))
                    {
                        result.Append(ampm.ToUpperInvariant());
                    }
                    else
                    {
                        result.Append(ampm.ToLowerInvariant());
                    }
                }
                else if (token.StartsWith(".0"))
                {
                    var value = date.Millisecond;
                    var digits = token.Length - 1;
                    result.Append("." + value.ToString("D" + digits));
                }
                else if (token == "/")
                {
#if NETSTANDARD1_0
                    result.Append(DateTime.MaxValue.ToString("/d", culture)[0]);
#else
                    result.Append(culture.DateTimeFormat.DateSeparator);
#endif
                }
                else if (token == ",")
                {
                    while (i < tokens.Count - 1 && tokens[i + 1] == ",")
                    {
                        i++;
                    }

                    result.Append(",");
                }
                else
                {
                    FormatLiteral(token, result);
                }
            }

            return result.ToString();
        }

        private static bool LookAheadDatePart(List<string> tokens, int fromIndex, string startsWith)
        {
            for (var i = fromIndex; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (token.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase))
                    return true;
                if (Token.IsDatePart(token))
                    return false;
            }

            return false;
        }

        private static bool LookBackDatePart(List<string> tokens, int fromIndex, string startsWith)
        {
            for (var i = fromIndex; i >= 0; i--)
            {
                var token = tokens[i];
                if (token.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase))
                    return true;
                if (Token.IsDatePart(token))
                    return false;
            }

            return false;
        }

        private static bool ContainsAmPm(List<string> tokens)
        {
            foreach (var token in tokens)
            {
                if (string.Compare(token, "am/pm", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return true;
                }

                if (string.Compare(token, "a/p", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return true;
                }
            }

            return false;
        }

        static string FormatNumber(double value, DecimalSection format, CultureInfo culture)
        {
            bool thousandSeparator = format.ThousandSeparator;
            value = value / format.ThousandDivisor;
            value = value * format.PercentMultiplier;

            var result = new StringBuilder();
            FormatNumber(value, format.BeforeDecimal, format.DecimalSeparator, format.AfterDecimal, thousandSeparator, culture, result);
            return result.ToString();
        }

        static void FormatNumber(double value, List<string> beforeDecimal, bool decimalSeparator, List<string> afterDecimal, bool thousandSeparator, CultureInfo culture, StringBuilder result)
        {
            int signitificantDigits = 0;
            if (afterDecimal != null)
                signitificantDigits = GetDigitCount(afterDecimal);

            var valueString = Math.Abs(value).ToString("F" + signitificantDigits, CultureInfo.InvariantCulture);
            var valueStrings = valueString.Split('.');
            var thousandsString = valueStrings[0];
            var decimalString = valueStrings.Length > 1 ? valueStrings[1].TrimEnd('0') : "";

            if (value < 0)
            {
                result.Append("-");
            }

            if (beforeDecimal != null)
            {
                FormatThousands(thousandsString, thousandSeparator, false, beforeDecimal, culture, result);
            }

            if (decimalSeparator) { 
                result.Append(culture.NumberFormat.NumberDecimalSeparator);
            }

            if (afterDecimal != null)
            {
                FormatDecimals(decimalString, afterDecimal, result);
            }
        }

        /// <summary>
        /// Prints right-aligned, left-padded integer before the decimal separator. With optional most-significant zero.
        /// </summary>
        public static void FormatThousands(string valueString, bool thousandSeparator, bool significantZero, List<string> tokens, CultureInfo culture, StringBuilder result)
        {
            var significant = false;
            var formatDigits = GetDigitCount(tokens);
            valueString = valueString.PadLeft(formatDigits, '0');

            // Print literals occurring before any placeholders
            var tokenIndex = 0;
            for (; tokenIndex < tokens.Count; tokenIndex++)
            {
                var token = tokens[tokenIndex];
                if (Token.IsPlaceholder(token))
                    break;
                else
                    FormatLiteral(token, result);
            }

            // Print value digits until there are as many digits remaining as there are placeholders
            var digitIndex = 0;
            for (; digitIndex < (valueString.Length - formatDigits); digitIndex++)
            {
                significant = true;
                result.Append(valueString[digitIndex]);

                if (thousandSeparator)
                    FormatThousandSeparator(valueString, digitIndex, culture, result);
            }

            // Print remaining value digits and format literals
            for (; tokenIndex < tokens.Count; ++tokenIndex)
            {
                var token = tokens[tokenIndex];
                if (Token.IsPlaceholder(token))
                {
                    var c = valueString[digitIndex];
                    if (c != '0' || (significantZero && digitIndex == valueString.Length - 1)) significant = true;

                    FormatPlaceholder(token, c, significant, result);

                    if (thousandSeparator && (significant || token.Equals("0")))
                        FormatThousandSeparator(valueString, digitIndex, culture, result);

                    digitIndex++;
                }
                else
                {
                    FormatLiteral(token, result);
                }
            }
        }

        static void FormatThousandSeparator(string valueString, int digit, CultureInfo culture, StringBuilder result)
        {
            var positionInTens = valueString.Length - 1 - digit;
            if (positionInTens > 0 && (positionInTens % 3) == 0)
            {
                result.Append(culture.NumberFormat.NumberGroupSeparator);
            }
        }

        /// <summary>
        /// Prints left-aligned, right-padded integer after the decimal separator. Does not print significant zero.
        /// </summary>
        public static void FormatDecimals(string valueString, List<string> tokens, StringBuilder result)
        {
            var significant = true;
            var unpaddedDigits = valueString.Length;
            var formatDigits = GetDigitCount(tokens);

            valueString = valueString.PadRight(formatDigits, '0');

            // Print all format digits
            var valueIndex = 0;
            for (var tokenIndex = 0; tokenIndex < tokens.Count; ++tokenIndex)
            {
                var token = tokens[tokenIndex];
                if (Token.IsPlaceholder(token))
                {
                    var c = valueString[valueIndex];
                    significant = valueIndex < unpaddedDigits;

                    FormatPlaceholder(token, c, significant, result);
                    valueIndex++;
                }
                else
                {
                    FormatLiteral(token, result);
                }
            }
        }

        static string FormatExponential(double value, Section format, CultureInfo culture)
        {
            // The application shall display a number to the right of 
            // the "E" symbol that corresponds to the number of places that 
            // the decimal point was moved. 

            var baseDigits = 0;
            if (format.Exponential.BeforeDecimal != null)
            {
                baseDigits = GetDigitCount(format.Exponential.BeforeDecimal);
            }

            var exponent = (int)Math.Floor(Math.Log10(Math.Abs(value)));
            var mantissa = value / Math.Pow(10, exponent);

            var shift = Math.Abs(exponent) % baseDigits;
            if (shift > 0)
            {
                if (exponent < 0)
                    shift = (baseDigits - shift);

                mantissa *= Math.Pow(10, shift);
                exponent -= shift;
            }

            var result = new StringBuilder();
            FormatNumber(mantissa, format.Exponential.BeforeDecimal, format.Exponential.DecimalSeparator, format.Exponential.AfterDecimal, false, culture, result);

            result.Append(format.Exponential.ExponentialToken[0]);

            if (format.Exponential.ExponentialToken[1] == '+' && exponent >= 0)
            {
                result.Append("+");
            }
            else if (exponent < 0)
            {
                result.Append("-");
            }

            FormatThousands(Math.Abs(exponent).ToString(CultureInfo.InvariantCulture), false, false, format.Exponential.Power, culture, result);
            return result.ToString();
        }

        static string FormatFraction(double value, Section format, CultureInfo culture)
        {
            int integral = 0;
            int numerator, denominator;

            bool sign = value < 0;

            if (format.Fraction.IntegerPart != null)
            {
                integral = (int)Math.Truncate(value);
                value = Math.Abs(value - integral);
            }

            if (format.Fraction.DenominatorConstant != 0)
            {
                denominator = format.Fraction.DenominatorConstant;
                var rr = Math.Round(value * denominator);
                var b = Math.Floor(rr / denominator);
                numerator = (int)(rr - b * denominator);
            }
            else
            {
                var denominatorDigits = Math.Min(GetDigitCount(format.Fraction.Denominator), 7);
                GetFraction(value, (int)Math.Pow(10, denominatorDigits) - 1, out numerator, out denominator);
            }

            // Don't hide fraction if at least one zero in the numerator format
            var numeratorZeros = GetZeroCount(format.Fraction.Numerator);
            var hideFraction = (format.Fraction.IntegerPart != null && numerator == 0 && numeratorZeros == 0);

            var result = new StringBuilder();

            if (sign)
                result.Append("-");

            // Print integer part with significant zero if fraction part is hidden
            if (format.Fraction.IntegerPart != null)
                FormatThousands(Math.Abs(integral).ToString("F0", CultureInfo.InvariantCulture), false, hideFraction, format.Fraction.IntegerPart, culture, result);

            var numeratorString = Math.Abs(numerator).ToString("F0", CultureInfo.InvariantCulture);
            var denominatorString = denominator.ToString("F0", CultureInfo.InvariantCulture);

            var fraction = new StringBuilder();

            FormatThousands(numeratorString, false, true, format.Fraction.Numerator, culture, fraction);

            fraction.Append("/");

            if (format.Fraction.DenominatorPrefix != null)
                FormatThousands("", false, false, format.Fraction.DenominatorPrefix, culture, fraction);

            if (format.Fraction.DenominatorConstant != 0)
                fraction.Append(format.Fraction.DenominatorConstant.ToString());
            else
                FormatDenominator(denominatorString, format.Fraction.Denominator, fraction);

            if (format.Fraction.DenominatorSuffix != null)
                FormatThousands("", false, false, format.Fraction.DenominatorSuffix, culture, fraction);

            if (hideFraction)
                result.Append(new string(' ', fraction.ToString().Length));
            else
                result.Append(fraction.ToString());

            if (format.Fraction.FractionSuffix != null)
                FormatThousands("", false, false, format.Fraction.FractionSuffix, culture, result);

            return result.ToString();
        }

        // Adapted from ssf.js 'frac()' helper
        static void GetFraction(double x, int D, out int nom, out int den)
        {
            var sgn = x < 0 ? -1 : 1;
            var B = x * sgn;
            var P_2 = 0.0;
            var P_1 = 1.0;
            var P = 0.0;
            var Q_2 = 1.0;
            var Q_1 = 0.0;
            var Q = 0.0;
            var A = Math.Floor(B);
            while (Q_1 < D)
            {
                A = Math.Floor(B);
                P = A * P_1 + P_2;
                Q = A * Q_1 + Q_2;
                if ((B - A) < 0.00000005) break;
                B = 1 / (B - A);
                P_2 = P_1; P_1 = P;
                Q_2 = Q_1; Q_1 = Q;
            }
            if (Q > D) { if (Q_1 > D) { Q = Q_2; P = P_2; } else { Q = Q_1; P = P_1; } }
            nom = (int)(sgn * P);
            den = (int)Q;
        }

        /// <summary>
        /// Prints left-aligned, left-padded fraction integer denominator.
        /// Assumes tokens contain only placeholders, valueString has fewer or equal number of digits as tokens.
        /// </summary>
        public static void FormatDenominator(string valueString, List<string> tokens, StringBuilder result)
        {
            var formatDigits = GetDigitCount(tokens);
            valueString = valueString.PadLeft(formatDigits, '0');

            bool significant = false;
            var valueIndex = 0;
            for (var tokenIndex = 0; tokenIndex < tokens.Count; ++tokenIndex)
            {
                var token = tokens[tokenIndex];
                char c;
                if (valueIndex < valueString.Length) {
                    c = GetLeftAlignedValueDigit(token, valueString, valueIndex, significant, out valueIndex);

                    if (c != '0')
                        significant = true;
                } else { 
                    c = '0';
                    significant = false;
                }

                FormatPlaceholder(token, c, significant, result);
            }
        }

        /// <summary>
        /// Returns the first digit from valueString. If the token is '?' 
        /// returns the first significant digit from valueString, or '0' if there are no significant digits.
        /// The out valueIndex parameter contains the offset to the next digit in valueString.
        /// </summary>
        static char GetLeftAlignedValueDigit(string token, string valueString, int startIndex, bool significant, out int valueIndex)
        {
            char c;
            valueIndex = startIndex;
            if (valueIndex < valueString.Length)
            {
                c = valueString[valueIndex];
                valueIndex++;

                if (c != '0')
                    significant = true;

                if (token == "?" && !significant)
                {
                    // Eat insignificant zeros to left align denominator
                    while (valueIndex < valueString.Length)
                    {
                        c = valueString[valueIndex];
                        valueIndex++;

                        if (c != '0')
                        {
                            significant = true;
                            break;
                        }
                    }
                }
            }
            else
            {
                c = '0';
                significant = false;
            }

            return c;
        }

        static void FormatPlaceholder(string token, char c, bool significant, StringBuilder result)
        {
            if (token == "0")
            {
                if (significant)
                    result.Append(c);
                else
                    result.Append("0");
            }
            else if (token == "#")
            {
                if (significant)
                    result.Append(c);
            }
            else if (token == "?")
            {
                if (significant)
                    result.Append(c);
                else
                    result.Append(" ");
            }
        }

        static int GetDigitCount(List<string> tokens)
        {
            var counter = 0;
            foreach (var token in tokens)
            {
                if (Token.IsPlaceholder(token))
                {
                    counter++;
                }
            }
            return counter;
        }

        static int GetZeroCount(List<string> tokens)
        {
            var counter = 0;
            foreach (var token in tokens)
            {
                if (token == "0")
                {
                    counter++;
                }
            }
            return counter;
        }

        static void FormatLiteral(string token, StringBuilder result)
        {
            string literal = string.Empty;
            if (token == ",")
            {
                ; // skip commas
            }
            else if (token.Length == 2 && (token[0] == '*' || token[0] == '\\'))
            {
                // TODO: * = repeat to fill cell
                literal = token[1].ToString();
            }
            else if (token.Length == 2 && token[0] == '_')
            {
                literal = " ";
            }
            else if (token.StartsWith("\""))
            {
                literal = token.Substring(1, token.Length - 2);
            }
            else
            {
                literal = token;
            }
            result.Append(literal);
        }
    }
}
