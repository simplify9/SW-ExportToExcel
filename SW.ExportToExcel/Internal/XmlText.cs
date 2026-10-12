using System.Text;

namespace SW.ExportToExcel.Internal
{
    internal static class XmlText
    {
        // Excel refuses to open a file holding a longer cell.
        internal const int MaxCellLength = 32767;

        // Removes characters XML 1.0 cannot carry (control characters, lone surrogates, U+FFFE/U+FFFF)
        // and truncates to what Excel accepts. Either would otherwise abort the export or corrupt the file.
        internal static string Clean(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            var firstInvalid = FirstInvalid(value);
            if (firstInvalid >= 0)
            {
                var builder = new StringBuilder(value.Length);
                builder.Append(value, 0, firstInvalid);
                for (var i = firstInvalid; i < value.Length; i++)
                {
                    var c = value[i];
                    if (char.IsHighSurrogate(c))
                    {
                        if (i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                        {
                            builder.Append(c).Append(value[i + 1]);
                            i++;
                        }
                    }
                    else if (IsValid(c))
                        builder.Append(c);
                }

                value = builder.ToString();
            }

            if (value.Length > MaxCellLength)
            {
                var length = char.IsHighSurrogate(value[MaxCellLength - 1]) ? MaxCellLength - 1 : MaxCellLength;
                value = value.Substring(0, length);
            }

            return value;
        }

        internal static bool NeedsPreserve(string value) =>
            value.Length > 0 && (IsXmlWhitespace(value[0]) || IsXmlWhitespace(value[value.Length - 1]));

        private static int FirstInvalid(string value)
        {
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (char.IsHighSurrogate(c) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                {
                    i++;
                    continue;
                }

                if (!IsValid(c))
                    return i;
            }

            return -1;
        }

        private static bool IsValid(char c) =>
            (c >= 0x20 && c <= 0xD7FF) || c == '\t' || c == '\n' || c == '\r' || (c >= 0xE000 && c <= 0xFFFD);

        private static bool IsXmlWhitespace(char c) => c == ' ' || c == '\t' || c == '\n' || c == '\r';
    }
}
