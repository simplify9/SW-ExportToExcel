using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxWriter = SW.ExportToExcel.Internal.XlsxWriter;

namespace SW.ExportToExcel.UnitTests
{
    [TestClass]
    public class ExcelExportTests
    {
        private static readonly DateTime When = new DateTime(2026, 10, 12, 9, 30, 0);

        private static async Task<Sheet> Write(Func<Stream, Task> write)
        {
            using var stream = new MemoryStream();
            await write(stream);
            return new Sheet(stream.ToArray());
        }

        [TestMethod]
        public async Task Columns_come_out_in_the_order_given()
        {
            var sheet = await Write(s => new[] { new Order { Id = 1, Customer = "c", Total = 9.5m } }.WriteExcelAsync(s,
                new Dictionary<string, string> { ["Total"] = "Total €", ["Customer"] = "Customer", ["Nope"] = "x" }));

            CollectionAssert.AreEqual(new[] { "Total €", "Customer" }, sheet.Row(1, 2));
            CollectionAssert.AreEqual(new[] { "9.5", "c" }, sheet.Row(2, 2));
            Assert.IsFalse(sheet.Has("C1"));
        }

        [TestMethod]
        public async Task All_supported_properties_by_default_with_typed_cells()
        {
            var sheet = await Write(s => new[]
            {
                new Order { Id = 1, Customer = "c", Paid = true, Placed = When, Due = new DateOnly(2026, 11, 1), Status = DayOfWeek.Friday, Ref = Guid.Empty },
            }.WriteExcelAsync(s));

            CollectionAssert.AreEqual(new[] { "Id", "Customer", "Total", "Paid", "Placed", "Due", "Status", "Ref", "Note" }, sheet.Row(1, 9));
            Assert.AreEqual("b", sheet.Type("D2"));
            Assert.AreEqual("1", sheet.Text("D2"));
            Assert.AreEqual(22u, sheet.NumberFormat("E2"));
            Assert.AreEqual(14u, sheet.NumberFormat("F2"));
            Assert.AreEqual(new DateTime(2026, 11, 1).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture), sheet.Text("F2"));
            Assert.AreEqual("Friday", sheet.Text("G2"));
            Assert.AreEqual(Guid.Empty.ToString(), sheet.Text("H2"));
            Assert.IsFalse(sheet.Has("J1"), "the List<string> property is not exportable");
        }

        [TestMethod]
        public async Task Empty_values_leave_the_cell_out_and_later_cells_keep_their_column()
        {
            var sheet = await Write(s => new[] { new Order { Id = 3, Customer = null, Note = "n" } }.WriteExcelAsync(s));

            Assert.IsFalse(sheet.Has("B2"));
            Assert.IsFalse(sheet.Has("E2"));
            Assert.AreEqual("n", sheet.Text("I2"));
            Assert.AreEqual("3", sheet.Text("A2"));
        }

        [TestMethod]
        public async Task Computed_columns()
        {
            var sheet = await Write(s => new[] { new Order { Id = 4, Customer = "acme", Total = 10m } }.WriteExcelAsync(s, new[]
            {
                new ExcelColumn<Order>("Upper", o => o.Customer.ToUpperInvariant()),
                new ExcelColumn<Order>("With tax", o => o.Total * 1.2m) { Width = 30 }
            }));

            CollectionAssert.AreEqual(new[] { "ACME", "12.0" }, sheet.Row(2, 2));
            var widths = sheet.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Column>().ToList();
            Assert.AreEqual(30d, widths.Single(c => c.Min == 2).Width.Value);
        }

        [TestMethod]
        public async Task Dictionary_and_expando_rows()
        {
            var rows = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object> { ["a"] = 1, ["b"] = "x" },
                new Dictionary<string, object> { ["b"] = "y", ["c"] = true }
            };
            var sheet = await Write(s => rows.WriteExcelAsync(s));
            CollectionAssert.AreEqual(new[] { "a", "b", "c" }, sheet.Row(1, 3));
            CollectionAssert.AreEqual(new[] { null, "y", "1" }, sheet.Row(3, 3));

            dynamic expando = new ExpandoObject();
            expando.Name = "dyn";
            expando.Count = 2;
            var dynamicSheet = await Write(s => new List<ExpandoObject> { expando }.WriteExcelAsync(s,
                new Dictionary<string, string> { ["Count"] = "How many", ["Name"] = "Name" }));
            CollectionAssert.AreEqual(new[] { "How many", "Name" }, dynamicSheet.Row(1, 2));
            CollectionAssert.AreEqual(new[] { "2", "dyn" }, dynamicSheet.Row(2, 2));
        }

        [TestMethod]
        public async Task Async_enumerable_source()
        {
            var sheet = await Write(s => Orders(5).WriteExcelAsync(s));

            Assert.AreEqual(6, sheet.RowCount);
            Assert.AreEqual("4", sheet.Text("A6"));
        }

        [TestMethod]
        public async Task Options_shape_the_sheet()
        {
            var sheet = await Write(s => new[] { new Order { Placed = When } }.WriteExcelAsync(s, new ExcelExportOptions
            {
                SheetName = "Orders: 2026/10 [final] and a very long tail",
                AutoFilter = true,
                FreezeHeaderRow = false,
                AutoFitColumns = false,
                DateTimeFormat = "yyyy-mm-dd hh:mm"
            }));

            Assert.AreEqual("Orders_ 2026_10 _final_ and a v", sheet.Name);
            Assert.AreEqual("A1:I2", sheet.Worksheet.Descendants<AutoFilter>().Single().Reference.Value);
            Assert.AreEqual("_xlnm._FilterDatabase", sheet.Workbook.DefinedNames.Elements<DefinedName>().Single().Name.Value);
            Assert.IsFalse(sheet.Worksheet.Descendants<Pane>().Any());
            Assert.IsFalse(sheet.Worksheet.Descendants<Columns>().Any());
            Assert.AreEqual(164u, sheet.NumberFormat("E2"));
        }

        [TestMethod]
        public async Task Auto_fit_sizes_from_header_and_sample_and_respects_the_cap()
        {
            var sheet = await Write(s => new[] { new Order { Customer = new string('w', 500) } }.WriteExcelAsync(s));

            var widths = sheet.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Column>().ToDictionary(c => c.Min.Value, c => c.Width.Value);
            Assert.AreEqual(60d, widths[2]);
            Assert.IsTrue(widths[1] >= 6 && widths[1] < 10);
            Assert.AreNotEqual(0u, sheet.Style("A1"), "header carries the bold style");
        }

        [TestMethod]
        public async Task Cancellation_is_honoured()
        {
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();

            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                Enumerable.Range(0, 10).Select(i => new Order { Id = i }).WriteExcelAsync(new MemoryStream(), cancellationToken: cancelled.Token));
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                Orders(10).WriteExcelAsync(new MemoryStream(), cancellationToken: cancelled.Token));
        }

        [TestMethod]
        public async Task Large_exports_stream()
        {
            const int count = 100_000;
            var sheet = await Write(s => Enumerable.Range(0, count)
                .Select(i => new Order { Id = i, Customer = "c" + i, Total = i, Placed = When }).WriteExcelAsync(s));

            Assert.AreEqual(count + 1, sheet.RowCount);
            Assert.AreEqual("c99999", sheet.Text("B100001"));
        }

        [TestMethod]
        [DataRow(0, "A")]
        [DataRow(25, "Z")]
        [DataRow(26, "AA")]
        [DataRow(676, "ZA")]
        [DataRow(701, "ZZ")]
        [DataRow(702, "AAA")]
        [DataRow(16383, "XFD")]
        public void Column_names(int index, string expected) => Assert.AreEqual(expected, XlsxWriter.ColumnName(index));

        private static async IAsyncEnumerable<Order> Orders(int count, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            for (var i = 0; i < count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Yield();
                yield return new Order { Id = i };
            }
        }

        private class Order
        {
            public int Id { get; set; }
            public string Customer { get; set; }
            public decimal Total { get; set; }
            public bool Paid { get; set; }
            public DateTime? Placed { get; set; }
            public DateOnly? Due { get; set; }
            public DayOfWeek? Status { get; set; }
            public Guid? Ref { get; set; }
            public string Note { get; set; }
            public List<string> Tags { get; set; } = new List<string>();
        }
    }
}
