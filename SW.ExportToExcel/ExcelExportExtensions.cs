using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using SW.ExportToExcel.Internal;

namespace SW.ExportToExcel
{
    /// <summary>
    /// Streaming export. Unlike the original <c>WriteExcel</c> methods, columns come out in the order given,
    /// booleans are real Excel booleans, empty values leave the cell empty, and enums, <c>Guid</c>,
    /// <c>DateTimeOffset</c> and <c>DateOnly</c> are exported. Rows can be <c>IDictionary&lt;string, object&gt;</c>
    /// (including <c>ExpandoObject</c>) or <c>IDictionary&lt;string, string&gt;</c>, keyed by column.
    /// </summary>
    /// <remarks>
    /// The package is written synchronously, so the stream must allow synchronous writes:
    /// write to a <see cref="MemoryStream"/> or a <see cref="FileStream"/>, not straight to an HTTP response body.
    /// </remarks>
    public static class ExcelExportExtensions
    {
        /// <summary>Every readable property of a supported type, in declaration order, headed by its name.</summary>
        public static Task WriteExcelAsync<T>(this IEnumerable<T> data, Stream stream,
            ExcelExportOptions options = null, CancellationToken cancellationToken = default)
        {
            options ??= ExcelExportOptions.Default;
            return Run(data, stream, options, ColumnPlanner.FromProperties, cancellationToken);
        }

        /// <summary>
        /// The given columns, in the given order: key is the property (or dictionary key), value the header.
        /// Properties that don't exist or have an unsupported type are skipped.
        /// </summary>
        public static Task WriteExcelAsync<T>(this IEnumerable<T> data, Stream stream,
            IEnumerable<KeyValuePair<string, string>> columns, ExcelExportOptions options = null,
            CancellationToken cancellationToken = default)
        {
            var planned = ColumnPlanner.FromNames<T>(columns);
            return Run(data, stream, options ?? ExcelExportOptions.Default, _ => planned, cancellationToken);
        }

        /// <summary>Computed columns, in the given order.</summary>
        public static Task WriteExcelAsync<T>(this IEnumerable<T> data, Stream stream,
            IEnumerable<ExcelColumn<T>> columns, ExcelExportOptions options = null,
            CancellationToken cancellationToken = default)
        {
            var planned = ColumnPlanner.FromSelectors(columns);
            return Run(data, stream, options ?? ExcelExportOptions.Default, _ => planned, cancellationToken);
        }

        /// <inheritdoc cref="WriteExcelAsync{T}(IEnumerable{T}, Stream, ExcelExportOptions, CancellationToken)"/>
        public static Task WriteExcelAsync<T>(this IAsyncEnumerable<T> data, Stream stream,
            ExcelExportOptions options = null, CancellationToken cancellationToken = default)
        {
            options ??= ExcelExportOptions.Default;
            return RunAsync(data, stream, options, ColumnPlanner.FromProperties, cancellationToken);
        }

        /// <inheritdoc cref="WriteExcelAsync{T}(IEnumerable{T}, Stream, IEnumerable{KeyValuePair{string, string}}, ExcelExportOptions, CancellationToken)"/>
        public static Task WriteExcelAsync<T>(this IAsyncEnumerable<T> data, Stream stream,
            IEnumerable<KeyValuePair<string, string>> columns, ExcelExportOptions options = null,
            CancellationToken cancellationToken = default)
        {
            var planned = ColumnPlanner.FromNames<T>(columns);
            return RunAsync(data, stream, options ?? ExcelExportOptions.Default, _ => planned, cancellationToken);
        }

        /// <inheritdoc cref="WriteExcelAsync{T}(IEnumerable{T}, Stream, IEnumerable{ExcelColumn{T}}, ExcelExportOptions, CancellationToken)"/>
        public static Task WriteExcelAsync<T>(this IAsyncEnumerable<T> data, Stream stream,
            IEnumerable<ExcelColumn<T>> columns, ExcelExportOptions options = null,
            CancellationToken cancellationToken = default)
        {
            var planned = ColumnPlanner.FromSelectors(columns);
            return RunAsync(data, stream, options ?? ExcelExportOptions.Default, _ => planned, cancellationToken);
        }

        private static Task Run<T>(IEnumerable<T> data, Stream stream, ExcelExportOptions options,
            Func<IReadOnlyList<T>, IReadOnlyList<Column>> plan, CancellationToken cancellationToken)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                ExportRunner.Run(data, stream, Layout(options), SampleSize(options), plan, Widths<T>(options),
                    cancellationToken);
                return Task.CompletedTask;
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                return Task.FromCanceled(cancellationToken);
            }
            catch (Exception e)
            {
                return Task.FromException(e);
            }
        }

        private static Task RunAsync<T>(IAsyncEnumerable<T> data, Stream stream, ExcelExportOptions options,
            Func<IReadOnlyList<T>, IReadOnlyList<Column>> plan, CancellationToken cancellationToken) =>
            ExportRunner.RunAsync(data, stream, Layout(options), SampleSize(options), plan, Widths<T>(options),
                cancellationToken);

        private static SheetLayout Layout(ExcelExportOptions options) => new SheetLayout
        {
            SheetName = options.SafeSheetName(),
            FreezeHeaderRow = options.FreezeHeaderRow,
            BoldHeader = options.BoldHeader,
            AutoFilter = options.AutoFilter,
            DateTimeFormat = options.DateTimeFormat,
            DateFormat = options.DateFormat
        };

        // Dictionary rows need at least one row read ahead to know their keys.
        private static int SampleSize(ExcelExportOptions options) =>
            Math.Max(1, options.AutoFitColumns ? options.AutoFitSampleRows : 1);

        private static Func<IReadOnlyList<Column>, IReadOnlyList<T>, double?[]> Widths<T>(ExcelExportOptions options) =>
            (columns, sample) => ColumnPlanner.Widths(columns, sample, options.AutoFitColumns, options.MaxColumnWidth);
    }
}
