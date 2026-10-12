using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

namespace SW.ExportToExcel.Internal
{
    // How the 2.x/8.x API typed a cell. Fixed per column from the property's declared type.
    internal enum LegacyKind
    {
        Text,
        Number,
        DateTime
    }

    internal sealed class Column
    {
        public Column(string header, Func<object, object> getter, LegacyKind legacyKind = LegacyKind.Text,
            double? width = null)
        {
            Header = header;
            Getter = getter;
            LegacyKind = legacyKind;
            Width = width;
        }

        public string Header { get; }
        public Func<object, object> Getter { get; }
        public LegacyKind LegacyKind { get; }
        public double? Width { get; }
    }

    internal sealed class SheetLayout
    {
        public string SheetName { get; set; } = "Exported";

        // Legacy: cells typed by LegacyKind, booleans as text, every cell written (no references),
        // the original two-entry stylesheet. Otherwise cells are typed by their runtime value,
        // empty cells are omitted and every cell carries its reference.
        public bool Legacy { get; set; }
        public bool FreezeHeaderRow { get; set; }
        public bool BoldHeader { get; set; }
        public bool AutoFilter { get; set; }
        public string DateTimeFormat { get; set; }
        public string DateFormat { get; set; }
    }

    // Streams one worksheet straight into the package: rows are written as they are enumerated,
    // so memory stays flat however many rows there are.
    internal sealed class XlsxWriter : IDisposable
    {
        private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string SheetRelId = "rId1";

        private const int DateTimeStyle = 1;
        private const int DateStyle = 2;
        private const int HeaderStyle = 3;

        private static readonly XmlWriterSettings XmlSettings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            NewLineHandling = NewLineHandling.None,
            CloseOutput = true
        };

        private readonly SheetLayout layout;
        private readonly IReadOnlyList<Column> columns;
        private readonly string[] columnNames;
        private readonly ZipArchive zip;

        private XmlWriter xml;
        private int rowNumber;
        private int nextColumn;
        private bool completed;

        public XlsxWriter(Stream stream, SheetLayout layout, IReadOnlyList<Column> columns)
        {
            this.layout = layout;
            this.columns = columns;
            columnNames = new string[columns.Count];
            for (var i = 0; i < columns.Count; i++)
                columnNames[i] = ColumnName(i);

            // The parts are written straight into a create-mode zip, which compresses as it goes and keeps
            // nothing in memory. (The OpenXml SDK only writes through a read-write package, which buffers
            // every part uncompressed until it is disposed.)
            zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        }

        public void Begin(IReadOnlyList<double?> widths)
        {
            xml = XmlWriter.Create(zip.CreateEntry("xl/worksheets/sheet1.xml").Open(), XmlSettings);
            xml.WriteStartDocument(true);
            xml.WriteStartElement("worksheet", MainNs);

            if (layout.FreezeHeaderRow)
            {
                xml.WriteStartElement("sheetViews");
                xml.WriteStartElement("sheetView");
                xml.WriteAttributeString("tabSelected", "1");
                xml.WriteAttributeString("workbookViewId", "0");
                xml.WriteStartElement("pane");
                xml.WriteAttributeString("ySplit", "1");
                xml.WriteAttributeString("topLeftCell", "A2");
                xml.WriteAttributeString("activePane", "bottomLeft");
                xml.WriteAttributeString("state", "frozen");
                xml.WriteEndElement();
                xml.WriteStartElement("selection");
                xml.WriteAttributeString("pane", "bottomLeft");
                xml.WriteAttributeString("activeCell", "A2");
                xml.WriteAttributeString("sqref", "A2");
                xml.WriteEndElement();
                xml.WriteEndElement();
                xml.WriteEndElement();
            }

            xml.WriteStartElement("sheetFormatPr");
            xml.WriteAttributeString("defaultRowHeight", "15");
            xml.WriteEndElement();

            if (widths != null && widths.Any(w => w != null))
            {
                xml.WriteStartElement("cols");
                for (var i = 0; i < widths.Count; i++)
                {
                    if (widths[i] == null)
                        continue;
                    var index = (i + 1).ToString(CultureInfo.InvariantCulture);
                    xml.WriteStartElement("col");
                    xml.WriteAttributeString("min", index);
                    xml.WriteAttributeString("max", index);
                    xml.WriteAttributeString("width", XmlConvert.ToString(Math.Round(widths[i].Value, 2)));
                    xml.WriteAttributeString("customWidth", "1");
                    xml.WriteEndElement();
                }

                xml.WriteEndElement();
            }

            xml.WriteStartElement("sheetData");

            StartRow();
            for (var i = 0; i < columns.Count; i++)
            {
                if (columns[i].Header == null)
                    continue;
                WriteText(i, columns[i].Header, layout.BoldHeader ? HeaderStyle : 0);
            }

            xml.WriteEndElement();
        }

        public void WriteRow(object item)
        {
            StartRow();
            for (var i = 0; i < columns.Count; i++)
            {
                var value = columns[i].Getter(item);
                if (layout.Legacy)
                    WriteLegacyCell(i, columns[i].LegacyKind, value);
                else
                    WriteCell(i, value);
            }

            xml.WriteEndElement();
        }

        public void Complete()
        {
            xml.WriteEndElement(); // sheetData

            var autoFilter = layout.AutoFilter && columns.Count > 0;
            var filterRange = autoFilter ? $"$A$1:${columnNames[columns.Count - 1]}${rowNumber}" : null;
            if (autoFilter)
            {
                xml.WriteStartElement("autoFilter");
                xml.WriteAttributeString("ref", filterRange.Replace("$", ""));
                xml.WriteEndElement();
            }

            xml.WriteEndElement(); // worksheet
            xml.WriteEndDocument();
            xml.Dispose();
            xml = null;

            WritePart("xl/styles.xml", WriteStyles);
            WritePart("xl/workbook.xml", w => WriteWorkbook(w, filterRange));
            WritePart("xl/_rels/workbook.xml.rels", w => WriteRelationships(w,
                (SheetRelId, RelNs + "/worksheet", "worksheets/sheet1.xml"),
                ("rId2", RelNs + "/styles", "styles.xml")));
            WritePart("_rels/.rels", w => WriteRelationships(w,
                ("rId1", RelNs + "/officeDocument", "xl/workbook.xml")));
            WritePart("[Content_Types].xml", WriteContentTypes);

            completed = true;
            zip.Dispose();
        }

        public void Dispose()
        {
            if (completed)
                return;
            xml?.Dispose();
            zip.Dispose();
        }

        private void StartRow()
        {
            rowNumber++;
            nextColumn = 0;
            xml.WriteStartElement("row");
            if (!layout.Legacy)
                xml.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
        }

        private void WriteLegacyCell(int column, LegacyKind kind, object value)
        {
            switch (kind)
            {
                case LegacyKind.Text:
                    WriteText(column, value is bool b ? XmlConvert.ToString(b) : (string)value ?? string.Empty, 0);
                    return;

                case LegacyKind.DateTime:
                    if (value is DateTime dateTime && dateTime.Year >= 100)
                        WriteNumber(column, XmlConvert.ToString(dateTime.ToOADate()), DateTimeStyle);
                    else
                        WriteNumber(column, string.Empty, 0);
                    return;

                default:
                    if (value == null)
                        WriteNumber(column, string.Empty, 0);
                    else if (TryFormatNumber(value, out var number))
                        WriteNumber(column, number, 0);
                    else
                        WriteText(column, number, 0); // NaN / infinity: Excel cannot open a non-finite number
                    return;
            }
        }

        private void WriteCell(int column, object value)
        {
            switch (value)
            {
                case null:
                    return;
                case string text:
                    if (text.Length > 0)
                        WriteText(column, text, 0);
                    return;
                case bool boolean:
                    WriteBoolean(column, boolean);
                    return;
                case DateTime dateTime:
                    WriteDate(column, dateTime, DateTimeStyle);
                    return;
                case DateTimeOffset dateTimeOffset:
                    WriteDate(column, dateTimeOffset.DateTime, DateTimeStyle);
                    return;
#if NET6_0_OR_GREATER
                case DateOnly date:
                    WriteDate(column, date.ToDateTime(TimeOnly.MinValue), DateStyle);
                    return;
#endif
                case Enum enumValue:
                    WriteText(column, enumValue.ToString(), 0);
                    return;
            }

            if (TryFormatNumber(value, out var number))
                WriteNumber(column, number, 0);
            else if (number != null)
                WriteText(column, number, 0);
            else if (value is IFormattable formattable)
                WriteText(column, formattable.ToString(null, CultureInfo.InvariantCulture), 0);
            else
                WriteText(column, value.ToString(), 0);
        }

        private void WriteDate(int column, DateTime value, int style)
        {
            // DateTime.ToOADate cannot represent years below 100; such values (DateTime.MinValue in
            // practice) mean "not set".
            if (value.Year >= 100)
                WriteNumber(column, XmlConvert.ToString(value.ToOADate()), style);
        }

        private void WriteText(int column, string value, int style)
        {
            value = XmlText.Clean(value);
            xml.WriteStartElement("c");
            WriteReference(column);
            if (style != 0)
                xml.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
            xml.WriteAttributeString("t", "inlineStr");
            xml.WriteStartElement("is");
            xml.WriteStartElement("t");
            if (XmlText.NeedsPreserve(value))
                xml.WriteAttributeString("xml", "space", null, "preserve");
            xml.WriteString(value);
            xml.WriteEndElement();
            xml.WriteEndElement();
            xml.WriteEndElement();
        }

        private void WriteNumber(int column, string value, int style)
        {
            xml.WriteStartElement("c");
            WriteReference(column);
            if (style != 0)
                xml.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
            xml.WriteElementString("v", value);
            xml.WriteEndElement();
        }

        private void WriteBoolean(int column, bool value)
        {
            xml.WriteStartElement("c");
            WriteReference(column);
            xml.WriteAttributeString("t", "b");
            xml.WriteElementString("v", value ? "1" : "0");
            xml.WriteEndElement();
        }

        // Cells are positional, so a reference is only needed after an empty cell was skipped.
        // Writing one on every cell made files about four times larger.
        private void WriteReference(int column)
        {
            var positional = column == nextColumn;
            nextColumn = column + 1;
            if (layout.Legacy || positional)
                return;
            xml.WriteStartAttribute("r");
            xml.WriteString(columnNames[column]);
            xml.WriteString(rowNumber.ToString(CultureInfo.InvariantCulture));
            xml.WriteEndAttribute();
        }

        // Formats a numeric value the way XElement did (XmlConvert, culture-invariant). A non-finite
        // float/double returns false with its text in `text`; a non-numeric value returns false and null.
        internal static bool TryFormatNumber(object value, out string text)
        {
            switch (value)
            {
                case int v: text = XmlConvert.ToString(v); return true;
                case long v: text = XmlConvert.ToString(v); return true;
                case decimal v: text = XmlConvert.ToString(v); return true;
                case short v: text = XmlConvert.ToString(v); return true;
                case byte v: text = XmlConvert.ToString(v); return true;
                case sbyte v: text = XmlConvert.ToString(v); return true;
                case ushort v: text = XmlConvert.ToString(v); return true;
                case uint v: text = XmlConvert.ToString(v); return true;
                case ulong v: text = XmlConvert.ToString(v); return true;
                case double v:
                    if (double.IsNaN(v) || double.IsInfinity(v))
                    {
                        text = NonFinite(v);
                        return false;
                    }

                    text = XmlConvert.ToString(v);
                    return true;
                case float v:
                    if (float.IsNaN(v) || float.IsInfinity(v))
                    {
                        text = NonFinite(v);
                        return false;
                    }

                    text = XmlConvert.ToString(v);
                    return true;
                default:
                    text = null;
                    return false;
            }
        }

        private static string NonFinite(double value) =>
            double.IsNaN(value) ? "NaN" : value > 0 ? "Infinity" : "-Infinity";

        private void WritePart(string path, Action<XmlWriter> write)
        {
            using var writer = XmlWriter.Create(zip.CreateEntry(path).Open(), XmlSettings);
            writer.WriteStartDocument(true);
            write(writer);
            writer.WriteEndDocument();
        }

        private void WriteWorkbook(XmlWriter w, string filterRange)
        {
            w.WriteStartElement("workbook", MainNs);
            w.WriteAttributeString("xmlns", "r", null, RelNs);
            w.WriteStartElement("bookViews");
            w.WriteElementString("workbookView", null);
            w.WriteEndElement();
            w.WriteStartElement("sheets");
            w.WriteStartElement("sheet");
            w.WriteAttributeString("name", layout.SheetName);
            w.WriteAttributeString("sheetId", "1");
            w.WriteAttributeString("id", RelNs, SheetRelId);
            w.WriteEndElement();
            w.WriteEndElement();
            if (filterRange != null)
            {
                w.WriteStartElement("definedNames");
                w.WriteStartElement("definedName");
                w.WriteAttributeString("name", "_xlnm._FilterDatabase");
                w.WriteAttributeString("localSheetId", "0");
                w.WriteAttributeString("hidden", "1");
                w.WriteString("'" + layout.SheetName.Replace("'", "''") + "'!" + filterRange);
                w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        private static void WriteRelationships(XmlWriter w, params (string id, string type, string target)[] relationships)
        {
            w.WriteStartElement("Relationships", PackageRelNs);
            foreach (var (id, type, target) in relationships)
            {
                w.WriteStartElement("Relationship");
                w.WriteAttributeString("Id", id);
                w.WriteAttributeString("Type", type);
                w.WriteAttributeString("Target", target);
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        private static void WriteContentTypes(XmlWriter w)
        {
            const string ooxml = "application/vnd.openxmlformats-officedocument.spreadsheetml.";
            w.WriteStartElement("Types", ContentTypesNs);
            foreach (var (extension, type) in new[]
                     {
                         ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
                         ("xml", "application/xml")
                     })
            {
                w.WriteStartElement("Default");
                w.WriteAttributeString("Extension", extension);
                w.WriteAttributeString("ContentType", type);
                w.WriteEndElement();
            }

            foreach (var (part, type) in new[]
                     {
                         ("/xl/workbook.xml", ooxml + "sheet.main+xml"),
                         ("/xl/worksheets/sheet1.xml", ooxml + "worksheet+xml"),
                         ("/xl/styles.xml", ooxml + "styles+xml")
                     })
            {
                w.WriteStartElement("Override");
                w.WriteAttributeString("PartName", part);
                w.WriteAttributeString("ContentType", type);
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        private void WriteStyles(XmlWriter w)
        {
            var legacy = layout.Legacy;
            var customFormats = new List<(int id, string code)>();
            var dateTimeFormatId = legacy ? 22 : FormatId(layout.DateTimeFormat, 22, customFormats);
            var dateFormatId = FormatId(layout.DateFormat, 14, customFormats);

            w.WriteStartElement("styleSheet", MainNs);

            if (!legacy && customFormats.Count > 0)
            {
                w.WriteStartElement("numFmts");
                w.WriteAttributeString("count", customFormats.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var (id, code) in customFormats)
                {
                    w.WriteStartElement("numFmt");
                    w.WriteAttributeString("numFmtId", id.ToString(CultureInfo.InvariantCulture));
                    w.WriteAttributeString("formatCode", code);
                    w.WriteEndElement();
                }

                w.WriteEndElement();
            }

            w.WriteStartElement("fonts");
            w.WriteAttributeString("count", legacy ? "1" : "2");
            WriteFont(w, false);
            if (!legacy)
                WriteFont(w, true);
            w.WriteEndElement();

            w.WriteStartElement("fills");
            w.WriteAttributeString("count", "2");
            foreach (var pattern in new[] { "none", "gray125" })
            {
                w.WriteStartElement("fill");
                w.WriteStartElement("patternFill");
                w.WriteAttributeString("patternType", pattern);
                w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();

            w.WriteStartElement("borders");
            w.WriteAttributeString("count", "1");
            w.WriteStartElement("border");
            foreach (var side in new[] { "left", "right", "top", "bottom", "diagonal" })
                w.WriteElementString(side, null);
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("cellStyleXfs");
            w.WriteAttributeString("count", "1");
            WriteXf(w, 0, 0, false);
            w.WriteEndElement();

            w.WriteStartElement("cellXfs");
            w.WriteAttributeString("count", legacy ? "2" : "4");
            WriteXf(w, 0, 0, true);
            WriteXf(w, dateTimeFormatId, 0, true);
            if (!legacy)
            {
                WriteXf(w, dateFormatId, 0, true);
                WriteXf(w, 0, 1, true);
            }

            w.WriteEndElement();

            w.WriteStartElement("cellStyles");
            w.WriteAttributeString("count", "1");
            w.WriteStartElement("cellStyle");
            w.WriteAttributeString("name", "Normal");
            w.WriteAttributeString("xfId", "0");
            w.WriteAttributeString("builtinId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("dxfs");
            w.WriteAttributeString("count", "0");
            w.WriteEndElement();
            w.WriteStartElement("tableStyles");
            w.WriteAttributeString("count", "0");
            w.WriteAttributeString("defaultTableStyle", "TableStyleMedium9");
            w.WriteAttributeString("defaultPivotStyle", "PivotStyleLight16");
            w.WriteEndElement();

            w.WriteEndElement();
        }

        private static int FormatId(string code, int builtIn, List<(int id, string code)> customFormats)
        {
            if (string.IsNullOrEmpty(code))
                return builtIn;
            var id = 164 + customFormats.Count;
            customFormats.Add((id, code));
            return id;
        }

        private static void WriteFont(XmlWriter w, bool bold)
        {
            w.WriteStartElement("font");
            if (bold)
                w.WriteElementString("b", null);
            WriteVal(w, "sz", "11");
            w.WriteStartElement("color");
            w.WriteAttributeString("theme", "1");
            w.WriteEndElement();
            WriteVal(w, "name", "Calibri");
            WriteVal(w, "family", "2");
            WriteVal(w, "scheme", "minor");
            w.WriteEndElement();
        }

        private static void WriteVal(XmlWriter w, string element, string value)
        {
            w.WriteStartElement(element);
            w.WriteAttributeString("val", value);
            w.WriteEndElement();
        }

        private static void WriteXf(XmlWriter w, int numFmtId, int fontId, bool inCellXfs)
        {
            w.WriteStartElement("xf");
            w.WriteAttributeString("numFmtId", numFmtId.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("fontId", fontId.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("fillId", "0");
            w.WriteAttributeString("borderId", "0");
            if (inCellXfs)
            {
                w.WriteAttributeString("xfId", "0");
                if (numFmtId != 0)
                    w.WriteAttributeString("applyNumberFormat", "1");
                if (fontId != 0)
                    w.WriteAttributeString("applyFont", "1");
            }

            w.WriteEndElement();
        }

        // 0 → A, 25 → Z, 26 → AA, 701 → ZZ, 702 → AAA.
        internal static string ColumnName(int index)
        {
            var name = string.Empty;
            for (var n = index + 1; n > 0; n = (n - 1) / 26)
                name = (char)('A' + (n - 1) % 26) + name;
            return name;
        }

        // Approximate display width, in characters, used to size columns from a sample of rows.
        internal static int DisplayLength(object value)
        {
            switch (value)
            {
                case null: return 0;
                case string text: return LongestLine(text);
                case bool _: return 5;
                case DateTime _:
                case DateTimeOffset _: return 16;
#if NET6_0_OR_GREATER
                case DateOnly _: return 10;
#endif
                case Enum enumValue: return enumValue.ToString().Length;
            }

            if (TryFormatNumber(value, out var number) || number != null)
                return number.Length;
            return value is IFormattable formattable
                ? formattable.ToString(null, CultureInfo.InvariantCulture).Length
                : value.ToString()?.Length ?? 0;
        }

        private static int LongestLine(string text)
        {
            int longest = 0, current = 0;
            foreach (var c in text)
            {
                if (c == '\n')
                {
                    current = 0;
                    continue;
                }

                if (++current > longest)
                    longest = current;
            }

            return longest;
        }
    }
}
