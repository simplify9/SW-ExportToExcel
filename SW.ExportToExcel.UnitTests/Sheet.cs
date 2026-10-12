using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SW.ExportToExcel.UnitTests
{
    // A written workbook read back: validated against the OOXML schema, cells addressed like "B3".
    internal sealed class Sheet
    {
        private readonly Dictionary<string, Cell> cells = new Dictionary<string, Cell>();

        public Sheet(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var document = SpreadsheetDocument.Open(stream, false);

            var errors = new OpenXmlValidator().Validate(document)
                .Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();
            Assert.IsEmpty(errors, string.Join(Environment.NewLine, errors));

            var workbookPart = document.WorkbookPart;
            var sheet = workbookPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Single();
            Name = sheet.Name;
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            Worksheet = worksheetPart.Worksheet;
            Workbook = workbookPart.Workbook;
            NumberFormats = workbookPart.WorkbookStylesPart?.Stylesheet?.CellFormats?.Elements<CellFormat>()
                .Select(f => f.NumberFormatId?.Value ?? 0).ToList() ?? new List<uint>();

            var rowNumber = 0;
            foreach (var row in Worksheet.Descendants<Row>())
            {
                rowNumber = (int?)row.RowIndex?.Value ?? rowNumber + 1;
                var column = 0;
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value is { } reference)
                        column = ColumnIndex(reference);
                    cells[ColumnName(column) + rowNumber] = cell;
                    column++;
                }

                RowCount = rowNumber;
            }
        }

        public string Name { get; }
        public Worksheet Worksheet { get; }
        public Workbook Workbook { get; }
        public List<uint> NumberFormats { get; }
        public int RowCount { get; }

        public Cell this[string reference] => cells.TryGetValue(reference, out var cell) ? cell : null;

        public bool Has(string reference) => cells.ContainsKey(reference);

        public string Text(string reference)
        {
            var cell = this[reference];
            return cell?.InlineString != null
                ? string.Concat(cell.InlineString.Elements<Text>().Select(t => t.Text))
                : cell?.CellValue?.Text;
        }

        public string Type(string reference) => this[reference]?.DataType?.InnerText;

        public uint Style(string reference) => this[reference]?.StyleIndex?.Value ?? 0;

        public uint NumberFormat(string reference) => NumberFormats[(int)Style(reference)];

        public List<string> Row(int row, int width) =>
            Enumerable.Range(0, width).Select(c => Text(ColumnName(c) + row)).ToList();

        private static int ColumnIndex(string reference)
        {
            var index = 0;
            foreach (var c in reference.TakeWhile(char.IsLetter))
                index = index * 26 + (c - 'A' + 1);
            return index - 1;
        }

        private static string ColumnName(int index)
        {
            var name = string.Empty;
            for (var n = index + 1; n > 0; n = (n - 1) / 26)
                name = (char)('A' + (n - 1) % 26) + name;
            return name;
        }
    }
}
