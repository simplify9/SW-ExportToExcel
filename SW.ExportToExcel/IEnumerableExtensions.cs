using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using SW.ExportToExcel.Internal;

namespace SW.ExportToExcel
{
    /// <summary>
    /// The original API, kept output-compatible: columns follow the type's property declaration order
    /// (a column map filters and renames them but does not reorder), booleans are written as the text
    /// "true"/"false", and only string, bool, DateTime and numeric properties are exported.
    /// For ordered columns, typed cells and more, use <see cref="ExcelExportExtensions"/>.
    /// </summary>
    public static class IEnumerableExtensions
    {
        public static async Task<byte[]> ExportToExcel<TEntity>(this IEnumerable<TEntity> data)
        {
            var dictionary = typeof(TEntity).GetProperties().ToDictionary(k => k.Name, v => v.Name);
            return await ExportToExcel(data, dictionary);
        }

        public static async Task<byte[]> ExportToExcel<TEntity>(this IEnumerable<TEntity> data,
            IEnumerable<string> columns)
        {
            var dictionary = columns.ToDictionary(k => k, v => v);
            return await ExportToExcel(data, dictionary);
        }

        public static async Task<byte[]> ExportToExcel<TEntity>(this IEnumerable<TEntity> data,
            IDictionary<string, string> columns)
        {
            using var stream = new MemoryStream();
            await WriteExcel(data, stream, columns);
            return stream.ToArray();
        }

        /// <summary>
        /// Rows of string values keyed by header. Headers are the first row's keys, in that row's order;
        /// keys missing from a later row leave its cell empty.
        /// </summary>
        public static Task WriteExcel(this IEnumerable<Dictionary<string, string>> data, Stream stream)
        {
            try
            {
                ExportRunner.Run(data, stream, new SheetLayout { Legacy = true, SheetName = "Sheet 1" }, 1,
                    sample => sample.Count == 0
                        ? Array.Empty<Column>()
                        : sample[0].Keys.Select(key => new Column(key,
                            row => ((Dictionary<string, string>)row).TryGetValue(key, out var value)
                                ? value ?? string.Empty
                                : string.Empty)).ToList(),
                    null);
                return Task.CompletedTask;
            }
            catch (Exception e)
            {
                return Task.FromException(e);
            }
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, string filePath)
        {
            var dictionary = typeof(TEntity).GetProperties().ToDictionary(k => k.Name, v => v.Name);
            await WriteExcel(data, filePath, dictionary);
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, string filePath,
            IEnumerable<string> columns)
        {
            var dictionary = columns.ToDictionary(k => k, v => v);
            await WriteExcel(data, filePath, dictionary);
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, string filePath,
            IDictionary<string, string> columns)
        {
            using var fileStream = File.Open(filePath, FileMode.Create, FileAccess.ReadWrite);
            await WriteExcel(data, fileStream, columns);
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, Stream stream)
        {
            var dictionary = typeof(TEntity).GetProperties().ToDictionary(k => k.Name, v => v.Name);
            await WriteExcel(data, stream, dictionary);
        }

        public static async Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, Stream stream,
            IEnumerable<string> columns)
        {
            var dictionary = columns.ToDictionary(k => k, v => v);
            await WriteExcel(data, stream, dictionary);
        }

        public static Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, Stream stream,
            IDictionary<string, string> columns)
        {
            try
            {
                var planned = ColumnPlanner.Legacy(typeof(TEntity), columns);
                ExportRunner.Run(data, stream, new SheetLayout { Legacy = true, FreezeHeaderRow = true }, 0,
                    _ => planned, null);
                return Task.CompletedTask;
            }
            catch (Exception e)
            {
                return Task.FromException(e);
            }
        }

        /// <summary>A sheet of text cells: one header row, then one row per list, values by position.</summary>
        public static byte[] GenerateExcel(IEnumerable<string> headers,
            IEnumerable<List<string>> dataRows)
        {
            var headerList = headers.ToList();
            var rows = dataRows.ToList();
            var width = rows.Select(r => r?.Count ?? 0).DefaultIfEmpty(0).Max();
            var columns = Enumerable.Range(0, Math.Max(headerList.Count, width))
                .Select(i => new Column(i < headerList.Count ? headerList[i] ?? string.Empty : null,
                    row => row is List<string> values && i < values.Count ? values[i] ?? string.Empty : string.Empty))
                .ToList();

            using var stream = new MemoryStream();
            ExportRunner.Run(rows, stream, new SheetLayout { Legacy = true, SheetName = "Sheet 1" }, 0,
                _ => columns, null);
            return stream.ToArray();
        }
    }
}
