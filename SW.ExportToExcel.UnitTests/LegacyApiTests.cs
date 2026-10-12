using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SW.ExportToExcel.UnitTests
{
    // The original API. These pin the output other applications depend on, plus the cases
    // that used to throw or produce a file Excel refuses to open.
    [TestClass]
    public class LegacyApiTests
    {
        private static async Task<Sheet> Write<T>(IEnumerable<T> data, IDictionary<string, string> columns = null)
        {
            using var stream = new MemoryStream();
            if (columns == null)
                await data.WriteExcel(stream);
            else
                await data.WriteExcel(stream, columns);
            return new Sheet(stream.ToArray());
        }

        [TestMethod]
        public async Task Columns_follow_declaration_order_and_the_map_only_filters_and_renames()
        {
            var sheet = await Write(new[] { new Person { Id = 1, Name = "a" } },
                new Dictionary<string, string> { ["Name"] = "Full name", ["Id"] = "Number", ["Missing"] = "x" });

            CollectionAssert.AreEqual(new[] { "Number", "Full name" }, sheet.Row(1, 2));
            CollectionAssert.AreEqual(new[] { "1", "a" }, sheet.Row(2, 2));
            Assert.IsFalse(sheet.Has("C1"));
        }

        [TestMethod]
        public async Task Cell_types_match_the_original_output()
        {
            var when = new DateTime(2026, 10, 12, 9, 30, 0);
            var sheet = await Write(new[]
            {
                new Mixed { Text = "t", Flag = true, Number = 5, Amount = 1.25m, Ratio = 0.5, When = when },
                new Mixed()
            });

            Assert.AreEqual("inlineStr", sheet.Type("B2"));
            Assert.AreEqual("true", sheet.Text("B2"));
            Assert.AreEqual("5", sheet.Text("D2"));
            Assert.AreEqual("1.25", sheet.Text("E2"));
            Assert.AreEqual(when.ToOADate().ToString("R", System.Globalization.CultureInfo.InvariantCulture), sheet.Text("G2"));
            Assert.AreEqual(22u, sheet.NumberFormat("G2"));

            // Nulls keep a placeholder cell, so nothing shifts left.
            Assert.AreEqual(string.Empty, sheet.Text("A3"));
            Assert.AreEqual("false", sheet.Text("B3"));
            Assert.AreEqual(string.Empty, sheet.Text("C3"));
            Assert.AreEqual(string.Empty, sheet.Text("D3"));
            Assert.AreEqual(string.Empty, sheet.Text("G3"));
            Assert.AreEqual(string.Empty, sheet.Text("H3"));
        }

        [TestMethod]
        public async Task Dates_before_year_100_are_written_empty()
        {
            var sheet = await Write(new[] { new Mixed { When = new DateTime(50, 1, 1) } });

            Assert.AreEqual(string.Empty, sheet.Text("G2"));
            Assert.AreEqual(0u, sheet.Style("G2"));
        }

        [TestMethod]
        public async Task Only_the_original_types_are_exported()
        {
            var sheet = await Write(new[] { new Unsupported() });

            CollectionAssert.AreEqual(new[] { "S" }, sheet.Row(1, 1));
            Assert.IsFalse(sheet.Has("B1"));
        }

        [TestMethod]
        public async Task Header_row_is_frozen_and_no_column_widths_are_forced()
        {
            var sheet = await Write(new[] { new Person() });

            var pane = sheet.Worksheet.Descendants<Pane>().Single();
            Assert.AreEqual(PaneStateValues.Frozen, pane.State.Value);
            Assert.IsInstanceOfType<SheetView>(pane.Parent);
            Assert.IsFalse(sheet.Worksheet.Descendants<Columns>().Any());
            Assert.AreEqual("Exported", sheet.Name);
        }

        [TestMethod]
        public async Task Empty_string_is_an_empty_cell_not_whitespace()
        {
            var sheet = await Write(new[] { new Person { Name = "" } });

            Assert.AreEqual(string.Empty, sheet.Text("B2"));
        }

        [TestMethod]
        public async Task Characters_xml_cannot_carry_are_removed()
        {
            var sheet = await Write(new[] { new Person { Name = "a\u0001b\u000Bc￾d\uD800e😀" } });

            Assert.AreEqual("abcde😀", sheet.Text("B2"));
        }

        [TestMethod]
        public async Task Strings_longer_than_excel_allows_are_truncated()
        {
            var sheet = await Write(new[] { new Person { Name = new string('x', 40000) } });

            Assert.AreEqual(32767, sheet.Text("B2").Length);
        }

        [TestMethod]
        public async Task Leading_and_trailing_spaces_are_preserved()
        {
            var sheet = await Write(new[] { new Person { Name = "  padded  " } });

            Assert.AreEqual("  padded  ", sheet.Text("B2"));
        }

        [TestMethod]
        public async Task Non_finite_numbers_are_written_as_text()
        {
            var sheet = await Write(new[]
            {
                new Mixed { Ratio = double.NaN }, new Mixed { Ratio = double.PositiveInfinity },
                new Mixed { Ratio = double.NegativeInfinity }
            });

            Assert.AreEqual("NaN", sheet.Text("F2"));
            Assert.AreEqual("Infinity", sheet.Text("F3"));
            Assert.AreEqual("-Infinity", sheet.Text("F4"));
            Assert.AreEqual("inlineStr", sheet.Type("F2"));
        }

        [TestMethod]
        public async Task Indexers_and_write_only_properties_are_skipped()
        {
            var sheet = await Write(new[] { new Awkward() });

            CollectionAssert.AreEqual(new[] { "A" }, sheet.Row(1, 1));
            Assert.IsFalse(sheet.Has("B1"));
        }

        [TestMethod]
        public async Task A_type_with_nothing_to_export_still_gives_a_valid_file()
        {
            var sheet = await Write(new object[] { new Person() });

            Assert.AreEqual(2, sheet.RowCount);
            Assert.IsFalse(sheet.Has("A1"));
        }

        [TestMethod]
        public async Task Dictionary_rows_are_matched_by_key_not_position()
        {
            var rows = new List<Dictionary<string, string>>
            {
                new Dictionary<string, string> { ["name"] = "Joe", ["age"] = "42" },
                new Dictionary<string, string> { ["age"] = "20", ["name"] = "Ann" },
                new Dictionary<string, string> { ["name"] = "Bob" }
            };
            using var stream = new MemoryStream();
            await rows.WriteExcel(stream);
            var sheet = new Sheet(stream.ToArray());

            Assert.AreEqual("Sheet 1", sheet.Name);
            CollectionAssert.AreEqual(new[] { "name", "age" }, sheet.Row(1, 2));
            CollectionAssert.AreEqual(new[] { "Ann", "20" }, sheet.Row(3, 2));
            CollectionAssert.AreEqual(new[] { "Bob", "" }, sheet.Row(4, 2));
        }

        [TestMethod]
        public async Task No_dictionary_rows_gives_an_empty_workbook()
        {
            using var stream = new MemoryStream();
            await new List<Dictionary<string, string>>().WriteExcel(stream);

            Assert.AreEqual(1, new Sheet(stream.ToArray()).RowCount);
        }

        [TestMethod]
        public void GenerateExcel_writes_text_by_position()
        {
            var bytes = IEnumerableExtensions.GenerateExcel(new[] { "h1", "h2" },
                new[] { new List<string> { "a", null }, new List<string> { "b", "c", "extra" } });
            var sheet = new Sheet(bytes);

            CollectionAssert.AreEqual(new[] { "h1", "h2", null }, sheet.Row(1, 3));
            CollectionAssert.AreEqual(new[] { "a", "", "" }, sheet.Row(2, 3));
            CollectionAssert.AreEqual(new[] { "b", "c", "extra" }, sheet.Row(3, 3));
        }

        [TestMethod]
        public async Task ExportToExcel_and_file_overloads_write_the_same_sheet()
        {
            var data = new[] { new Person { Id = 7, Name = "x" } };
            var bytes = await data.ExportToExcel(new[] { "Name" });
            var path = Path.GetTempFileName();
            try
            {
                await data.WriteExcel(path, new[] { "Name" });
                CollectionAssert.AreEqual(new Sheet(bytes).Row(2, 1), new Sheet(File.ReadAllBytes(path)).Row(2, 1));
                CollectionAssert.AreEqual(new[] { "x" }, new Sheet(bytes).Row(2, 1));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public async Task Errors_surface_through_the_task()
        {
            var task = new[] { new Person() }.WriteExcel(new MemoryStream(), (IDictionary<string, string>)null);

            await Assert.ThrowsExactlyAsync<ArgumentNullException>(() => task);
        }

        private class Person
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        private class Mixed
        {
            public string Text { get; set; }
            public bool Flag { get; set; }
            public bool? MaybeFlag { get; set; }
            public int? Number { get; set; }
            public decimal Amount { get; set; }
            public double Ratio { get; set; }
            public DateTime? When { get; set; }
            public long? Big { get; set; }
        }

        private class Unsupported
        {
            public string S { get; set; } = "s";
            public Guid G { get; set; } = Guid.NewGuid();
            public DayOfWeek E { get; set; }
            public List<int> L { get; set; } = new List<int>();
        }

        private class Awkward
        {
            private string hidden;
            public string A { get; set; } = "a";
            public string this[int i] => "indexer";
            public string WriteOnly { set => hidden = value; }
        }
    }
}
