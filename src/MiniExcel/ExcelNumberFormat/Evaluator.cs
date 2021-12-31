using System;
using System.Collections.Generic;
using System.Text;

namespace MiniExcelNumberFormat
{
    internal static class Evaluator
    {
        public static Section GetSection(List<Section> sections, object value)
        {
            // Standard format has up to 4 sections:
            // Positive;Negative;Zero;Text
            switch (value)
            {
                case string s:
                    if (sections.Count >= 4)
                        return sections[3];

                    return null;

                case DateTime dt:
                    // TODO: Check date conditions need date helpers and Date1904 knowledge
                    return GetFirstSection(sections, SectionType.Date);

                case TimeSpan ts:
                    return GetNumericSection(sections, ts.TotalDays);

                case double d:
                    return GetNumericSection(sections, d);

                case int i:
                    return GetNumericSection(sections, i);

                case short s:
                    return GetNumericSection(sections, s);

                default:
                    return null;
            }
        }

        public static Section GetFirstSection(List<Section> sections, SectionType type)
        {
            foreach (var section in sections)
                if (section.Type == type)
                    return section;
            return null;
        }

        private static Section GetNumericSection(List<Section> sections, double value)
        {
            // First section applies if 
            // - Has a condition:
            // - There is 1 section, or
            // - There are 2 sections, and the value is 0 or positive, or
            // - There are >2 sections, and the value is positive
            if (sections.Count < 1)
            {
                return null;
            }

            var section0 = sections[0];

            if (section0.Condition != null)
            {
                if (section0.Condition.Evaluate(value))
                {
                    return section0;
                }
            }
            else if (sections.Count == 1 || (sections.Count == 2 && value >= 0) || (sections.Count >= 2 && value > 0))
            {
                return section0;
            }

            if (sections.Count < 2)
            {
                return null;
            }

            var section1 = sections[1];

            // First condition didnt match, or was a negative number. Second condition applies if:
            // - Has a condition, or
            // - Value is negative, or
            // - There are two sections, and the first section had a non-matching condition
            if (section1.Condition != null)
            {
                if (section1.Condition.Evaluate(value))
                {
                    return section1;
                }
            }
            else if (value < 0 || (sections.Count == 2 && section0.Condition != null))
            {
                return section1;
            }

            // Second condition didnt match, or was positive. The following 
            // sections cannot have conditions, always fall back to the third 
            // section (for zero formatting) if specified.
            if (sections.Count < 3)
            {
                return null;
            }

            return sections[2];
        }
    }
}
