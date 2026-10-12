using System.Text;

namespace SW.ExportToExcel
{
    /// <summary>Options for <see cref="ExcelExportExtensions"/>. Not used by the original <c>WriteExcel</c>/<c>ExportToExcel</c> methods.</summary>
    public sealed class ExcelExportOptions
    {
        internal static readonly ExcelExportOptions Default = new ExcelExportOptions();

        /// <summary>Worksheet name. Characters Excel forbids (<c>[ ] : * ? / \</c>) become <c>_</c>; longer than 31 characters is truncated.</summary>
        public string SheetName { get; set; } = "Exported";

        /// <summary>Keep the header row visible while scrolling.</summary>
        public bool FreezeHeaderRow { get; set; } = true;

        public bool BoldHeader { get; set; } = true;

        /// <summary>Add Excel's filter drop-downs to the header row.</summary>
        public bool AutoFilter { get; set; }

        /// <summary>Size each column to its header and the first <see cref="AutoFitSampleRows"/> rows.</summary>
        public bool AutoFitColumns { get; set; } = true;

        /// <summary>
        /// Rows read ahead before anything is written, for column sizing and for discovering the keys of
        /// dictionary rows. Keys that first appear after these rows are not exported unless columns are given.
        /// </summary>
        public int AutoFitSampleRows { get; set; } = 100;

        public double MaxColumnWidth { get; set; } = 60;

        /// <summary>Excel number format for <c>DateTime</c>/<c>DateTimeOffset</c> cells, e.g. <c>yyyy-mm-dd hh:mm</c>. Null uses Excel's built-in short date and time.</summary>
        public string DateTimeFormat { get; set; }

        /// <summary>Excel number format for <c>DateOnly</c> cells. Null uses Excel's built-in short date.</summary>
        public string DateFormat { get; set; }

        internal string SafeSheetName()
        {
            var name = new StringBuilder(SheetName ?? string.Empty);
            foreach (var c in "[]:*?/\\")
                name.Replace(c, '_');
            var result = name.ToString().Trim('\'');
            if (result.Length > 31)
                result = result.Substring(0, 31);
            return result.Length == 0 ? "Sheet1" : result;
        }
    }
}
