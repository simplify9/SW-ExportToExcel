using System;

namespace SW.ExportToExcel
{
    /// <summary>A computed column: a header and a function producing the cell value from a row.</summary>
    /// <example><code>new ExcelColumn&lt;Shipment&gt;("Receiver city", s =&gt; s.Receiver.City)</code></example>
    public sealed class ExcelColumn<T>
    {
        public ExcelColumn(string header, Func<T, object> value)
        {
            Header = header ?? throw new ArgumentNullException(nameof(header));
            Value = value ?? throw new ArgumentNullException(nameof(value));
        }

        public string Header { get; }

        /// <summary>The cell value. Numbers, booleans and dates keep their Excel type; anything else is written as text.</summary>
        public Func<T, object> Value { get; }

        /// <summary>Fixed width in characters. Overrides auto-fit for this column.</summary>
        public double? Width { get; set; }
    }
}
