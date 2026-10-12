using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace SW.ExportToExcel.Internal
{
    internal static class ExportRunner
    {
        private const int CancellationCheckInterval = 1024;

        // Pulls up to `sampleSize` rows first (to discover dictionary keys and size columns), then streams the rest.
        internal static void Run<T>(IEnumerable<T> data, Stream stream, SheetLayout layout, int sampleSize,
            Func<IReadOnlyList<T>, IReadOnlyList<Column>> plan, Func<IReadOnlyList<Column>, IReadOnlyList<T>, double?[]> widths,
            CancellationToken cancellationToken = default)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            using var enumerator = data.GetEnumerator();
            var sample = new List<T>();
            while (sample.Count < sampleSize && enumerator.MoveNext())
                sample.Add(enumerator.Current);

            var columns = plan(sample);
            using var writer = new XlsxWriter(stream, layout, columns);
            writer.Begin(widths?.Invoke(columns, sample));

            foreach (var item in sample)
                writer.WriteRow(item);

            var count = 0;
            while (enumerator.MoveNext())
            {
                if (++count % CancellationCheckInterval == 0)
                    cancellationToken.ThrowIfCancellationRequested();
                writer.WriteRow(enumerator.Current);
            }

            writer.Complete();
        }

        internal static async Task RunAsync<T>(IAsyncEnumerable<T> data, Stream stream, SheetLayout layout, int sampleSize,
            Func<IReadOnlyList<T>, IReadOnlyList<Column>> plan, Func<IReadOnlyList<Column>, IReadOnlyList<T>, double?[]> widths,
            CancellationToken cancellationToken)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            await using var enumerator = data.GetAsyncEnumerator(cancellationToken);
            var sample = new List<T>();
            while (sample.Count < sampleSize && await enumerator.MoveNextAsync().ConfigureAwait(false))
                sample.Add(enumerator.Current);

            var columns = plan(sample);
            using var writer = new XlsxWriter(stream, layout, columns);
            writer.Begin(widths?.Invoke(columns, sample));

            foreach (var item in sample)
                writer.WriteRow(item);

            var count = 0;
            while (await enumerator.MoveNextAsync().ConfigureAwait(false))
            {
                if (++count % CancellationCheckInterval == 0)
                    cancellationToken.ThrowIfCancellationRequested();
                writer.WriteRow(enumerator.Current);
            }

            writer.Complete();
        }
    }

    internal static class ColumnPlanner
    {
        // The property types the 2.x/8.x API exported; anything else was silently left out.
        private static readonly Dictionary<Type, LegacyKind> LegacyTypes = new Dictionary<Type, LegacyKind>
        {
            [typeof(string)] = LegacyKind.Text,
            [typeof(bool)] = LegacyKind.Text,
            [typeof(bool?)] = LegacyKind.Text,
            [typeof(DateTime)] = LegacyKind.DateTime,
            [typeof(DateTime?)] = LegacyKind.DateTime,
            [typeof(int)] = LegacyKind.Number,
            [typeof(int?)] = LegacyKind.Number,
            [typeof(byte)] = LegacyKind.Number,
            [typeof(byte?)] = LegacyKind.Number,
            [typeof(short)] = LegacyKind.Number,
            [typeof(short?)] = LegacyKind.Number,
            [typeof(long)] = LegacyKind.Number,
            [typeof(long?)] = LegacyKind.Number,
            [typeof(float)] = LegacyKind.Number,
            [typeof(float?)] = LegacyKind.Number,
            [typeof(double)] = LegacyKind.Number,
            [typeof(double?)] = LegacyKind.Number,
            [typeof(decimal)] = LegacyKind.Number,
            [typeof(decimal?)] = LegacyKind.Number,
        };

        // Legacy: declared property order, filtered by `columns`' keys, headers from its values.
        internal static IReadOnlyList<Column> Legacy(Type type, IDictionary<string, string> columns)
        {
            if (columns == null) throw new ArgumentNullException(nameof(columns));
            var result = new List<Column>();
            foreach (var property in type.GetProperties())
            {
                if (!LegacyTypes.TryGetValue(property.PropertyType, out var kind) || !IsReadable(property) ||
                    !columns.ContainsKey(property.Name))
                    continue;
                result.Add(new Column(columns[property.Name], Getter(property), kind));
            }

            return result;
        }

        internal static IReadOnlyList<Column> FromProperties<T>(IReadOnlyList<T> sample)
        {
            var dictionaryGetter = DictionaryGetter(typeof(T));
            if (dictionaryGetter != null)
                return DictionaryKeys(sample).Select(key => new Column(key, dictionaryGetter(key))).ToList();

            return typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => IsReadable(p) && IsSupported(p.PropertyType))
                .Select(p => new Column(p.Name, Getter(p)))
                .ToList();
        }

        internal static IReadOnlyList<Column> FromNames<T>(IEnumerable<KeyValuePair<string, string>> columns)
        {
            if (columns == null) throw new ArgumentNullException(nameof(columns));
            var dictionaryGetter = DictionaryGetter(typeof(T));
            var result = new List<Column>();
            foreach (var column in columns)
            {
                if (dictionaryGetter != null)
                {
                    result.Add(new Column(column.Value ?? column.Key, dictionaryGetter(column.Key)));
                    continue;
                }

                var property = typeof(T).GetProperty(column.Key, BindingFlags.Public | BindingFlags.Instance);
                if (property != null && IsReadable(property) && IsSupported(property.PropertyType))
                    result.Add(new Column(column.Value ?? column.Key, Getter(property)));
            }

            return result;
        }

        internal static IReadOnlyList<Column> FromSelectors<T>(IEnumerable<ExcelColumn<T>> columns)
        {
            if (columns == null) throw new ArgumentNullException(nameof(columns));
            return columns.Select(c => new Column(c.Header, item => c.Value((T)item), width: c.Width)).ToList();
        }

        internal static double?[] Widths<T>(IReadOnlyList<Column> columns, IReadOnlyList<T> sample,
            bool autoFit, double maxWidth)
        {
            if (!autoFit && columns.All(c => c.Width == null))
                return null;

            var widths = new double?[columns.Count];
            for (var i = 0; i < columns.Count; i++)
            {
                if (columns[i].Width != null || !autoFit)
                {
                    widths[i] = columns[i].Width;
                    continue;
                }

                // Header is bold, which renders roughly 10% wider.
                var longest = (columns[i].Header?.Length ?? 0) * 1.1;
                foreach (var item in sample)
                    longest = Math.Max(longest, XlsxWriter.DisplayLength(columns[i].Getter(item)));
                widths[i] = Math.Min(maxWidth, Math.Max(6, longest + 2));
            }

            return widths;
        }

        private static IEnumerable<string> DictionaryKeys<T>(IReadOnlyList<T> sample)
        {
            var seen = new HashSet<string>();
            foreach (var row in sample)
            foreach (var key in Keys(row))
                if (key != null && seen.Add(key))
                    yield return key;
        }

        private static IEnumerable<string> Keys(object row) => row switch
        {
            IDictionary<string, object> d => d.Keys,
            IReadOnlyDictionary<string, object> d => d.Keys,
            IDictionary<string, string> d => d.Keys,
            IReadOnlyDictionary<string, string> d => d.Keys,
            _ => Enumerable.Empty<string>()
        };

        private static Func<string, Func<object, object>> DictionaryGetter(Type type)
        {
            if (typeof(IDictionary<string, object>).IsAssignableFrom(type))
                return key => row => row is IDictionary<string, object> d && d.TryGetValue(key, out var v) ? v : null;
            if (typeof(IReadOnlyDictionary<string, object>).IsAssignableFrom(type))
                return key => row => row is IReadOnlyDictionary<string, object> d && d.TryGetValue(key, out var v) ? v : null;
            if (typeof(IDictionary<string, string>).IsAssignableFrom(type))
                return key => row => row is IDictionary<string, string> d && d.TryGetValue(key, out var v) ? v : null;
            if (typeof(IReadOnlyDictionary<string, string>).IsAssignableFrom(type))
                return key => row => row is IReadOnlyDictionary<string, string> d && d.TryGetValue(key, out var v) ? v : null;
            return null;
        }

        private static bool IsReadable(PropertyInfo property) =>
            property.GetGetMethod() != null && property.GetIndexParameters().Length == 0;

        internal static bool IsSupported(Type type)
        {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type.IsPrimitive && type != typeof(IntPtr) && type != typeof(UIntPtr)
                   || type.IsEnum
                   || type == typeof(string)
                   || type == typeof(decimal)
                   || type == typeof(DateTime)
                   || type == typeof(DateTimeOffset)
                   || type == typeof(Guid)
                   || type == typeof(TimeSpan)
#if NET6_0_OR_GREATER
                   || type == typeof(DateOnly)
                   || type == typeof(TimeOnly)
#endif
                ;
        }

        private static Func<object, object> Getter(PropertyInfo property) => property.GetValue;
    }
}
