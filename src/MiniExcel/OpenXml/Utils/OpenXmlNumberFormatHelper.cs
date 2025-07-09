namespace MiniExcelLib.OpenXml.Utils;

/// <summary>
/// This code edit from https://github.com/andersnm/ExcelNumberFormat
/// </summary>
internal class OpenXmlNumberFormatHelper
{
    /// <summary>
    /// Initializes a new instance of the <see cref="OpenXmlNumberFormatHelper"/> class.
    /// </summary>
    /// <param name="formatString">The number format string.</param>
    public OpenXmlNumberFormatHelper(string formatString)
    {
        var sections = Parser.ParseSections(formatString, out bool syntaxError);

        IsValid = !syntaxError;
        FormatString = formatString;

        if (IsValid)
        {
            Sections = sections;
            IsDateTimeFormat = Evaluator.GetFirstSection(Sections, SectionType.Date) is not null;
            IsTimeSpanFormat = Evaluator.GetFirstSection(Sections, SectionType.Duration) is not null;
        }
        else
        {
            Sections = [];
        }
    }

    /// <summary>
    /// Gets a value indicating whether the number format string is valid.
    /// </summary>
    public bool IsValid { get; }

    /// <summary>
    /// Gets the number format string.
    /// </summary>
    public string FormatString { get; }

    /// <summary>
    /// Gets a value indicating whether the format represents a DateTime
    /// </summary>
    public bool IsDateTimeFormat { get; }

    /// <summary>
    /// Gets a value indicating whether the format represents a TimeSpan
    /// </summary>
    public bool IsTimeSpanFormat { get; }

    internal List<Section> Sections { get; }
}

internal static class Evaluator
{
    public static Section? GetSection(List<Section> sections, object value)
    {
        // Standard format has up to 4 sections:
        // Positive;Negative;Zero;Text
        return value switch
        {
            string s => sections.Count >= 4 ? sections[3] : null,
            DateTime dt => GetFirstSection(sections, SectionType.Date), // TODO: Check date conditions need date helpers and Date1904 knowledge
            TimeSpan ts => GetNumericSection(sections),
            double d => GetNumericSection(sections),
            int i => GetNumericSection(sections),
            short s => GetNumericSection(sections),
            _ => null
        };
    }

    public static Section? GetFirstSection(List<Section> sections, SectionType type)
    {
        return sections.FirstOrDefault(section => section.Type == type);
    }

    private static Section? GetNumericSection(List<Section> sections)
    {
        return sections.Count < 3 ? null : sections[2];
    }
}

internal enum SectionType
{
    General,
    Number,
    Fraction,
    Exponential,
    Date,
    Duration,
    Text
}

internal class Section
{
    public int SectionIndex { get; set; }
    public SectionType Type { get; set; }
    public List<string>? GeneralTextDateDurationParts { get; set; }
}

internal class FractionSection
{
    public List<string> IntegerPart { get; set; }
    public List<string> Numerator { get; set; }
    public List<string> DenominatorPrefix { get; set; }
    public List<string> Denominator { get; set; }
    public int DenominatorConstant { get; set; }
    public List<string> DenominatorSuffix { get; set; }
    public List<string> FractionSuffix { get; set; }

    public static bool TryParse(List<string> tokens, out FractionSection? format)
    {
        List<string>? numeratorParts = null;
        List<string>? denominatorParts = null;

        for (var i = 0; i < tokens.Count; i++)
        {
            var part = tokens[i];
            if (part == "/")
            {
                numeratorParts = tokens.GetRange(0, i);
                i++;
                denominatorParts = tokens.GetRange(i, tokens.Count - i);
                break;
            }
        }

        if (numeratorParts is null)
        {
            format = null;
            return false;
        }

        GetNumerator(numeratorParts, out var integerPart, out var numeratorPart);

        if (!TryGetDenominator(denominatorParts, out var denominatorPrefix, out var denominatorPart, out var denominatorConstant, out var denominatorSuffix, out var fractionSuffix))
        {
            format = null;
            return false;
        }

        format = new FractionSection
        {
            IntegerPart = integerPart,
            Numerator = numeratorPart,
            DenominatorPrefix = denominatorPrefix,
            Denominator = denominatorPart,
            DenominatorConstant = denominatorConstant,
            DenominatorSuffix = denominatorSuffix,
            FractionSuffix = fractionSuffix
        };

        return true;
    }

    private static void GetNumerator(List<string> tokens, out List<string>? integerPart, out List<string>? numeratorPart)
    {
        var hasPlaceholder = false;
        var hasSpace = false;
        var hasIntegerPart = false;
        var numeratorIndex = -1;
        var index = tokens.Count - 1;
        while (index >= 0)
        {
            var token = tokens[index];
            if (Token.IsPlaceholder(token))
            {
                hasPlaceholder = true;

                if (hasSpace)
                {
                    hasIntegerPart = true;
                    break;
                }
            }
            else
            {
                if (hasPlaceholder && !hasSpace)
                {
                    // First time we get here marks the end of the integer part
                    hasSpace = true;
                    numeratorIndex = index + 1;
                }
            }
            index--;
        }

        if (hasIntegerPart)
        {
            integerPart = tokens.GetRange(0, numeratorIndex);
            numeratorPart = tokens.GetRange(numeratorIndex, tokens.Count - numeratorIndex);
        }
        else
        {
            integerPart = null;
            numeratorPart = tokens;
        }
    }

    private static bool TryGetDenominator(List<string>? tokens, out List<string>? denominatorPrefix, out List<string>? denominatorPart, out int denominatorConstant, out List<string>? denominatorSuffix, out List<string>? fractionSuffix)
    {
        var index = 0;
        var hasPlaceholder = false;
        var hasConstant = false;

        var constant = new StringBuilder();

        // Read literals until the first number placeholder or digit
        while (index < tokens?.Count)
        {
            var token = tokens[index];
            if (Token.IsPlaceholder(token))
            {
                hasPlaceholder = true;
                break;
            }
            else
            if (Token.IsDigit19(token))
            {
                hasConstant = true;
                break;
            }
            index++;
        }

        if (!hasPlaceholder && !hasConstant)
        {
            denominatorPrefix = null;
            denominatorPart = null;
            denominatorConstant = 0;
            denominatorSuffix = null;
            fractionSuffix = null;
            return false;
        }

        // The denominator starts here, keep the index
        var denominatorIndex = index;

        // Read placeholders or digits in sequence
        while (index < tokens.Count)
        {
            var token = tokens[index];
            if (hasPlaceholder && Token.IsPlaceholder(token))
            {
                // OK
            }
            else
            if (hasConstant && (Token.IsDigit09(token)))
            {
                constant.Append(token);
            }
            else
            {
                break;
            }
            index++;
        }

        // 'index' is now at the first token after the denominator placeholders.
        // The remaining, if anything, is to be treated in one or two parts:
        // Any ultimately terminating literals are considered the "Fraction suffix".
        // Anything between the denominator and the fraction suffix is the "Denominator suffix".
        // Placeholders in the denominator suffix are treated as insignificant zeros.

        // Scan backwards to determine the fraction suffix
        int fractionSuffixIndex = tokens.Count;
        while (fractionSuffixIndex > index)
        {
            var token = tokens[fractionSuffixIndex - 1];
            if (Token.IsPlaceholder(token))
            {
                break;
            }

            fractionSuffixIndex--;
        }

        // Finally extract the detected token ranges
        denominatorPrefix = denominatorIndex > 0 ? tokens.GetRange(0, denominatorIndex) : null;
        denominatorConstant = hasConstant ? int.Parse(constant.ToString()) : 0;
        denominatorPart = tokens.GetRange(denominatorIndex, index - denominatorIndex);
        denominatorSuffix = index < fractionSuffixIndex ? tokens.GetRange(index, fractionSuffixIndex - index) : null;
        fractionSuffix = fractionSuffixIndex < tokens.Count ? tokens.GetRange(fractionSuffixIndex, tokens.Count - fractionSuffixIndex) : null;

        return true;
    }
}

internal class ExponentialSection
{
    public List<string> BeforeDecimal { get; set; }
    public bool DecimalSeparator { get; set; }
    public List<string> AfterDecimal { get; set; }
    public string ExponentialToken { get; set; }
    public List<string> Power { get; set; }

    public static bool TryParse(List<string> tokens, out ExponentialSection? format)
    {
        format = null;

        string exponentialToken;

        int partCount = Parser.ParseNumberTokens(tokens, 0, out var beforeDecimal, out var decimalSeparator, out var afterDecimal);

        if (partCount == 0)
            return false;

        int position = partCount;
        if (position < tokens.Count && Token.IsExponent(tokens[position]))
        {
            exponentialToken = tokens[position];
            position++;
        }
        else
        {
            return false;
        }

        format = new ExponentialSection
        {
            BeforeDecimal = beforeDecimal,
            DecimalSeparator = decimalSeparator,
            AfterDecimal = afterDecimal,
            ExponentialToken = exponentialToken,
            Power = tokens.GetRange(position, tokens.Count - position)
        };

        return true;
    }
}

internal class DecimalSection
{
    public bool ThousandSeparator { get; set; }
    public double ThousandDivisor { get; set; }
    public double PercentMultiplier { get; set; }
    public List<string> BeforeDecimal { get; set; }
    public bool DecimalSeparator { get; set; }
    public List<string> AfterDecimal { get; set; }

    public static bool TryParse(List<string> tokens, out DecimalSection? format)
    {
        if (Parser.ParseNumberTokens(tokens, 0, out var beforeDecimal, out var decimalSeparator, out var afterDecimal) == tokens.Count)
        {
            var divisor = GetTrailingCommasDivisor(tokens, out var thousandSeparator);
            var multiplier = GetPercentMultiplier(tokens);

            format = new DecimalSection
            {
                BeforeDecimal = beforeDecimal,
                DecimalSeparator = decimalSeparator,
                AfterDecimal = afterDecimal,
                PercentMultiplier = multiplier,
                ThousandDivisor = divisor,
                ThousandSeparator = thousandSeparator
            };

            return true;
        }

        format = null;
        return false;
    }

    private static double GetPercentMultiplier(List<string> tokens)
    {
        // If there is a percentage literal in the part list, multiply the result by 100
        return tokens.Any(token => token == "%") ? 100 : 1;
    }

    private static double GetTrailingCommasDivisor(List<string> tokens, out bool thousandSeparator)
    {
        // This parses all comma literals in the part list:
        // Each comma after the last digit placeholder divides the result by 1000.
        // If there are any other commas, display the result with thousand separators.
        bool hasLastPlaceholder = false;
        var divisor = 1.0;

        for (var j = 0; j < tokens.Count; j++)
        {
            var tokenIndex = tokens.Count - 1 - j;
            var token = tokens[tokenIndex];

            if (!hasLastPlaceholder)
            {
                if (Token.IsPlaceholder(token))
                {
                    // Each trailing comma multiplies the divisor by 1000
                    for (var k = tokenIndex + 1; k < tokens.Count; k++)
                    {
                        token = tokens[k];
                        if (token == ",")
                            divisor *= 1000.0;
                        else
                            break;
                    }

                    // Continue scanning backwards from the last digit placeholder, 
                    // but now look for a thousand separator comma
                    hasLastPlaceholder = true;
                }
            }
            else
            {
                if (token == ",")
                {
                    thousandSeparator = true;
                    return divisor;
                }
            }
        }

        thousandSeparator = false;
        return divisor;
    }
}

/// <summary>
/// Similar to regular .NET DateTime, but also supports 0/1 1900 and 29/2 1900.
/// </summary>
internal class ExcelDateTime
{
    /// <summary>
    /// The closest .NET DateTime to the specified excel date.
    /// </summary>
    public DateTime AdjustedDateTime { get; }

    /// <summary>
    /// Number of days to adjust by in post.
    /// </summary>
    public int AdjustDaysPost { get; }

    /// <summary>
    /// Constructs a new ExcelDateTime from a numeric value.
    /// </summary>
    public ExcelDateTime(double numericDate, bool isDate1904)
    {
        if (isDate1904)
        {
            numericDate += 1462.0;
            AdjustedDateTime = new DateTime(DoubleDateToTicks(numericDate), DateTimeKind.Unspecified);
        }
        else
        {
            // internal dates before 30/12/1899 should add two days to get the real date
            // internal dates on 30/12 19899 should add two days, but subtract a day post to get the real date
            // internal dates before 28/2/1900 should add one day to get the real date
            // internal dates on 28/2 1900 should use the same date, but add a day post to get the real date

            var internalDateTime = new DateTime(DoubleDateToTicks(numericDate), DateTimeKind.Unspecified);
            if (internalDateTime < Excel1900ZeroethMinDate)
            {
                AdjustDaysPost = 0;
                AdjustedDateTime = internalDateTime.AddDays(2);
            }

            else if (internalDateTime < Excel1900ZeroethMaxDate)
            {
                AdjustDaysPost = -1;
                AdjustedDateTime = internalDateTime.AddDays(2);
            }

            else if (internalDateTime < Excel1900LeapMinDate)
            {
                AdjustDaysPost = 0;
                AdjustedDateTime = internalDateTime.AddDays(1);
            }

            else if (internalDateTime < Excel1900LeapMaxDate)
            {
                AdjustDaysPost = 1;
                AdjustedDateTime = internalDateTime;
            }
            else
            {
                AdjustDaysPost = 0;
                AdjustedDateTime = internalDateTime;
            }
        }
    }

    private static readonly DateTime Excel1900LeapMinDate = new(1900, 2, 28);
    private static readonly DateTime Excel1900LeapMaxDate = new(1900, 3, 1);
    private static readonly DateTime Excel1900ZeroethMinDate = new(1899, 12, 30);
    private static readonly DateTime Excel1900ZeroethMaxDate = new(1899, 12, 31);

    /// <summary>
    /// Wraps a regular .NET datetime.
    /// </summary>
    /// <param name="value"></param>
    public ExcelDateTime(DateTime value)
    {
        AdjustedDateTime = value;
        AdjustDaysPost = 0;
    }

    public int Year => AdjustedDateTime.Year;
    public int Month => AdjustedDateTime.Month;
    public int Day => AdjustedDateTime.Day + AdjustDaysPost;
    public int Hour => AdjustedDateTime.Hour;
    public int Minute => AdjustedDateTime.Minute;
    public int Second => AdjustedDateTime.Second;
    public int Millisecond => AdjustedDateTime.Millisecond;
    public DayOfWeek DayOfWeek => AdjustedDateTime.DayOfWeek;

    public string ToString(string numberFormat, CultureInfo culture)
    {
        return AdjustedDateTime.ToString(numberFormat, culture);
    }

    public static bool TryConvert(object value, bool isDate1904, CultureInfo culture, out ExcelDateTime? result)
    {
        switch (value)
        {
            case double doubleValue:
                result = new ExcelDateTime(doubleValue, isDate1904);
                return true;
            case int intValue:
                result = new ExcelDateTime(intValue, isDate1904);
                return true;
            case short shortValue:
                result = new ExcelDateTime(shortValue, isDate1904);
                return true;
            case DateTime dateTimeValue:
                result = new ExcelDateTime(dateTimeValue);
                return true;
            default:
                result = null;
                return false;
        }
    }

    // From DateTime class to enable OADate in PCL
    // Number of 100ns ticks per time unit
    private const long TicksPerMillisecond = 10000;
    private const long TicksPerSecond = TicksPerMillisecond * 1000;
    private const long TicksPerMinute = TicksPerSecond * 60;
    private const long TicksPerHour = TicksPerMinute * 60;
    private const long TicksPerDay = TicksPerHour * 24;

    private const int MillisPerSecond = 1000;
    private const int MillisPerMinute = MillisPerSecond * 60;
    private const int MillisPerHour = MillisPerMinute * 60;
    private const int MillisPerDay = MillisPerHour * 24;

    // Number of days in a non-leap year
    private const int DaysPerYear = 365;

    // Number of days in 4 years
    private const int DaysPer4Years = DaysPerYear * 4 + 1;

    // Number of days in 100 years
    private const int DaysPer100Years = DaysPer4Years * 25 - 1;

    // Number of days in 400 years
    private const int DaysPer400Years = DaysPer100Years * 4 + 1;

    // Number of days from 1/1/0001 to 12/30/1899
    private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

    private const long DoubleDateOffset = DaysTo1899 * TicksPerDay;

    internal static long DoubleDateToTicks(double value)
    {
        long millis = (long)(value * MillisPerDay + (value >= 0 ? 0.5 : -0.5));

        // The interesting thing here is when you have a value like 12.5 it all positive 12 days and 12 hours from 01/01/1899
        // However if you a value of -12.25 it is minus 12 days but still positive 6 hours, almost as though you meant -11.75 all negative
        // This line below fixes up the millis in the negative case
        if (millis < 0)
        {
            millis -= millis % MillisPerDay * 2;
        }

        millis += DoubleDateOffset / TicksPerMillisecond;
        return millis * TicksPerMillisecond;
    }
}

internal static class Parser
{
    public static List<Section> ParseSections(string formatString, out bool syntaxError)
    {
        var tokenizer = new Tokenizer(formatString);
        var sections = new List<Section>();
        syntaxError = false;
        while (true)
        {
            var section = ParseSection(tokenizer, sections.Count, out var sectionSyntaxError);

            if (sectionSyntaxError)
                syntaxError = true;

            if (section is null)
                break;

            sections.Add(section);
        }

        return sections;
    }

    private static Section? ParseSection(Tokenizer reader, int index, out bool syntaxError)
    {
        bool hasDateParts = false;
        bool hasDurationParts = false;
        bool hasGeneralPart = false;
        bool hasTextPart = false;
        bool hasPlaceholders = false;
        List<string> tokens = [];

        syntaxError = false;
        while (ReadToken(reader, out syntaxError) is { } token)
        {
            if (token == ";")
                break;

            hasPlaceholders |= Token.IsPlaceholder(token);

            if (Token.IsDatePart(token))
            {
                hasDateParts = true;
                hasDurationParts |= Token.IsDurationPart(token);
                tokens.Add(token);
            }
            else
            {
                tokens.Add(token);
            }
        }

        if (syntaxError || tokens.Count == 0)
        {
            return null;
        }

        if (
            (hasDateParts && (hasGeneralPart || hasTextPart)) ||
            (hasGeneralPart && (hasDateParts || hasTextPart)) ||
            (hasTextPart && (hasGeneralPart || hasDateParts)))
        {
            // Cannot mix date, general and/or text parts
            syntaxError = true;
            return null;
        }

        SectionType type; 
        FractionSection? fraction = null;
        ExponentialSection? exponential = null;
        DecimalSection? number = null;
        List<string>? generalTextDateDuration = null;

        if (hasDateParts)
        {
            type = hasDurationParts ? SectionType.Duration : SectionType.Date;
            ParseMilliseconds(tokens, out generalTextDateDuration);
        }
        else if (hasGeneralPart)
        {
            type = SectionType.General;
            generalTextDateDuration = tokens;
        }
        else if (hasTextPart || !hasPlaceholders)
        {
            type = SectionType.Text;
            generalTextDateDuration = tokens;
        }
        else if (FractionSection.TryParse(tokens, out fraction))
        {
            type = SectionType.Fraction;
        }
        else if (ExponentialSection.TryParse(tokens, out exponential))
        {
            type = SectionType.Exponential;
        }
        else if (DecimalSection.TryParse(tokens, out number))
        {
            type = SectionType.Number;
        }
        else
        {
            // Unable to parse format string
            syntaxError = true;
            return null;
        }

        return new Section
        {
            Type = type,
            SectionIndex = index,
            GeneralTextDateDurationParts = generalTextDateDuration
        };
    }

    internal static int ParseNumberTokens(List<string> tokens, int startPosition, out List<string>? beforeDecimal, out bool decimalSeparator, out List<string>? afterDecimal)
    {
        beforeDecimal = null;
        afterDecimal = null;
        decimalSeparator = false;

        List<string> remainder = [];
        var index = 0;
        for (index = 0; index < tokens.Count; ++index)
        {
            var token = tokens[index];
            if (token == "." && beforeDecimal is null)
            {
                decimalSeparator = true;
                beforeDecimal = tokens.GetRange(0, index); // TODO: why not remainder? has only valid tokens...

                remainder = [];
            }
            else if (Token.IsNumberLiteral(token))
            {
                remainder.Add(token);
            }
            else if (token.StartsWith("["))
            {
                // ignore
            }
            else
            {
                break;
            }
        }

        if (remainder.Count > 0)
        {
            if (beforeDecimal is not null)
            {
                afterDecimal = remainder;
            }
            else
            {
                beforeDecimal = remainder;
            }
        }

        return index;
    }

    private static void ParseMilliseconds(List<string> tokens, out List<string> result)
    {
        // if tokens form .0 through .000.., combine to single subsecond token
        result = [];
        for (var i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (token == ".")
            {
                var zeros = 0;
                while (i + 1 < tokens.Count && tokens[i + 1] == "0")
                {
                    i++;
                    zeros++;
                }

                if (zeros > 0)
                    result.Add("." + new string('0', zeros));
                else
                    result.Add(".");
            }
            else
            {
                result.Add(token);
            }
        }
    }

    private static string? ReadToken(Tokenizer reader, out bool syntaxError)
    {
        var offset = reader.Position;
        if (
            ReadLiteral(reader) ||
            reader.ReadEnclosed('[', ']') ||

            // Symbols
            reader.ReadOneOf("#?,!&%+-$€£0123456789{}():;/.@ ") ||
            reader.ReadString("e+", true) ||
            reader.ReadString("e-", true) ||
            reader.ReadString("General", true) ||

            // Date
            reader.ReadString("am/pm", true) ||
            reader.ReadString("a/p", true) ||
            reader.ReadOneOrMore('y') ||
            reader.ReadOneOrMore('Y') ||
            reader.ReadOneOrMore('m') ||
            reader.ReadOneOrMore('M') ||
            reader.ReadOneOrMore('d') ||
            reader.ReadOneOrMore('D') ||
            reader.ReadOneOrMore('h') ||
            reader.ReadOneOrMore('H') ||
            reader.ReadOneOrMore('s') ||
            reader.ReadOneOrMore('S') ||
            reader.ReadOneOrMore('g') ||
            reader.ReadOneOrMore('G'))
        {
            syntaxError = false;
            var length = reader.Position - offset;
            return reader.Substring(offset, length);
        }

        syntaxError = reader.Position < reader.Length;
        return null;
    }

    private static bool ReadLiteral(Tokenizer reader)
    {
        if (reader.Peek() == '\\' || reader.Peek() == '*' || reader.Peek() == '_')
        {
            reader.Advance(2);
            return true;
        }

        if (reader.ReadEnclosed('"', '"'))
        {
            return true;
        }

        return false;
    }
}

internal class Tokenizer(string? fmt)
{
    private readonly string? _formatString = fmt;

    public int Position { get; private set; } = 0;
    public int Length => _formatString?.Length ?? 0;

    public string? Substring(int startIndex, int length)
    {
        return _formatString?.Substring(startIndex, length);
    }

    public int Peek(int offset = 0)
    {
        if (Position + offset >= Length)
            return -1;

        return _formatString?[Position + offset] ?? 0;
    }

    public int PeekUntil(int startOffset, int until)
    {
        int offset = startOffset;
        while (true)
        {
            var c = Peek(offset++);
            if (c == -1)
                break;
            if (c == until)
                return offset - startOffset;
        }
        return 0;
    }

    public bool PeekOneOf(int offset, string s)
    {
        foreach (var c in s)
        {
            if (Peek(offset) == c)
            {
                return true;
            }
        }
        return false;
    }

    public void Advance(int characters = 1)
    {
        Position = Math.Min(Position + characters, _formatString.Length);
    }

    public bool ReadOneOrMore(int c)
    {
        if (Peek() != c)
            return false;

        while (Peek() == c)
            Advance();

        return true;
    }

    public bool ReadOneOf(string s)
    {
        if (PeekOneOf(0, s))
        {
            Advance();
            return true;
        }
        return false;
    }

    public bool ReadString(string s, bool ignoreCase = false)
    {
        if (Position + s.Length > Length)
            return false;

        for (var i = 0; i < s.Length; i++)
        {
            var c1 = s[i];
            var c2 = (char)Peek(i);
            if (ignoreCase)
            {
                if (char.ToLower(c1) != char.ToLower(c2)) return false;
            }
            else
            {
                if (c1 != c2) return false;
            }
        }

        Advance(s.Length);
        return true;
    }

    public bool ReadEnclosed(char open, char close)
    {
        if (Peek() == open)
        {
            int length = PeekUntil(1, close);
            if (length > 0)
            {
                Advance(1 + length);
                return true;
            }
        }

        return false;
    }
}

internal static class Token
{
    public static bool IsExponent(string token) =>
        string.Compare(token, "e+", StringComparison.OrdinalIgnoreCase) == 0 ||
        string.Compare(token, "e-", StringComparison.OrdinalIgnoreCase) == 0;

    public static bool IsLiteral(string token) =>
        token.StartsWith("_", StringComparison.Ordinal) ||
        token.StartsWith("\\", StringComparison.Ordinal) ||
        token.StartsWith("\"", StringComparison.Ordinal) ||
        token.StartsWith("*", StringComparison.Ordinal) ||
        token == "," ||
        token == "!" ||
        token == "&" ||
        token == "%" ||
        token == "+" ||
        token == "-" ||
        token == "$" ||
        token == "€" ||
        token == "£" ||
        token == "1" ||
        token == "2" ||
        token == "3" ||
        token == "4" ||
        token == "5" ||
        token == "6" ||
        token == "7" ||
        token == "8" ||
        token == "9" ||
        token == "{" ||
        token == "}" ||
        token == "(" ||
        token == ")" ||
        token == " ";

    public static bool IsNumberLiteral(string token) =>
        IsPlaceholder(token) ||
        IsLiteral(token) ||
        token == ".";

    public static bool IsPlaceholder(string token) => token is "0" or "#" or "?";

    public static bool IsGeneral(string token)
    {
        return string.Compare(token, "general", StringComparison.OrdinalIgnoreCase) == 0;
    }

    public static bool IsDatePart(string token) =>
        token.StartsWith("y", StringComparison.OrdinalIgnoreCase) ||
        token.StartsWith("m", StringComparison.OrdinalIgnoreCase) ||
        token.StartsWith("d", StringComparison.OrdinalIgnoreCase) ||
        token.StartsWith("s", StringComparison.OrdinalIgnoreCase) ||
        token.StartsWith("h", StringComparison.OrdinalIgnoreCase) ||
        (token.StartsWith("g", StringComparison.OrdinalIgnoreCase) && !IsGeneral(token)) ||
        string.Compare(token, "am/pm", StringComparison.OrdinalIgnoreCase) == 0 ||
        string.Compare(token, "a/p", StringComparison.OrdinalIgnoreCase) == 0 ||
        IsDurationPart(token);

    public static bool IsDurationPart(string token) =>
        token.StartsWith("[h", StringComparison.OrdinalIgnoreCase) ||
        token.StartsWith("[m", StringComparison.OrdinalIgnoreCase) ||
        token.StartsWith("[s", StringComparison.OrdinalIgnoreCase);

    public static bool IsDigit09(string token)
    {
        return token == "0" || IsDigit19(token);
    }

    public static bool IsDigit19(string token) => token switch
    {
        "1" or "2" or "3" or "4" or "5" or "6" or "7" or "8" or "9" => true,
        _ => false
    };
}