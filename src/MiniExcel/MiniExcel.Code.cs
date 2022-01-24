namespace MiniExcelLibs
{
    using System;
    using System.Linq;

    public static partial class MiniExcel
    {
        public static string LISENCE_CODE = null;
        private static bool _HAS_LICENSE = false;
        private static bool _FirstTime = true;
        /// <summary>
        /// Please feel free to take license code from here ^___^
        /// </summary>
        private static string[] LISENCES = new string[10] {
            "096b9eff-c0e7-4813-8352-211811063214",
            "e5db2221-aef1-4f84-8121-82182c179863",
            "71430e9c-c8e0-4a7d-bab6-261535b2ac9d",
            "0cdbd11c-60c5-41cd-9ba3-3d0de67f72df",
            "33498913-6a61-4940-929b-6bab04c02bba",
            "9557941a-acf7-4dbd-a81b-bc46aafa15b7",
            "e470d4df-f4ec-4895-82cd-c3ba76d94ca8",
            "18c7fed4-3688-42aa-9a0f-bd57021f81d3",
            "8ce7ea6f-aa15-4ea7-81a0-334308313508",
            "39982af1-3452-4489-9fb3-5f86a51835c1",
        };

        private static void CheckLicense()
        {
            if (!_FirstTime)
                return;
            _FirstTime = false;
            if (_HAS_LICENSE)
                return;
            if (LISENCES.Contains(LISENCE_CODE))
            {
                _HAS_LICENSE = true;
                return;
            }
            try
            {
                Console.WriteLine(@"Dear user, you can donate/get code to hide this console output, web link : https://shps951023.github.io/MiniExcel and set MiniExcel.LISENCE_CODE = ""YourCode""");
            }
            catch (Exception) { }
        }
    }
}
