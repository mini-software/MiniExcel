using System.Collections.Generic;

namespace MiniExcelNumberFormat{
    internal class Section
    {
        public int SectionIndex { get; set; }

        public SectionType Type { get; set; }

        public Color Color { get; set; }

        public Condition Condition { get; set; }

        public ExponentialSection Exponential { get; set; }

        public FractionSection Fraction { get; set; }

        public DecimalSection Number { get; set; }

        public List<string> GeneralTextDateDurationParts { get; set; }
    }
}