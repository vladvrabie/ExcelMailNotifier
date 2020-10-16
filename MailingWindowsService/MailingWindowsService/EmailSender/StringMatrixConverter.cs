using System.Linq;
using System.Text;
using StringMatrix = System.Collections.Generic.List<System.Collections.Generic.List<string>>;


namespace MailingWindowsService.EmailSender
{
    class StringMatrixConverter
    {
        public static string ToPlainTextTable(StringMatrix data)
        {
            if (data == null)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            int maxLength = data.Max((row) => row?.Max((cell) => cell?.Length ?? 0) ?? 0);

            {
                var headerDataRow = data[0];
                sb.Append("  ");
                headerDataRow?.ForEach((cell) => sb.Append($"| {CenteredString(cell ?? string.Empty, maxLength)} "));
                sb.Append("|\n");
                sb.Append("  ");
                headerDataRow?.ForEach((cell) => sb.Append($"| {new string('-', maxLength)} "));
                sb.Append("|\n");
            }

            foreach (var dataRow in data.Skip(1))
            {
                sb.Append("  ");
                dataRow?.ForEach((cell) =>
                    {
                        sb.Append(string.Format("| {0,-" + maxLength.ToString() + "} ", cell ?? string.Empty));
                    });
                sb.Append("|\n");
            }

            return sb.ToString();
        }

        public static string ToHtmlTable(StringMatrix data)
        {
            if (data == null)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            sb.Append("<table style=\"width:80%\" border=\"1\" cellpadding=\"3\" >");

            {
                var headerDataRow = data[0];
                sb.Append("<tr>");
                headerDataRow?.ForEach((cell) => sb.Append($"<th>{cell ?? string.Empty}</th>"));
                sb.Append("</tr>");
            }

            foreach (var dataRow in data.Skip(1))
            {
                sb.Append("<tr>");
                dataRow?.ForEach((cell) => sb.Append($"<td>{cell ?? string.Empty}</td>"));
                sb.Append("</tr>");
            }

            sb.Append("</table>");
            return sb.ToString();
        }

        private static string CenteredString(string s, int width)
        {
            // Credits https://stackoverflow.com/a/18573196

            if (s == null || s.Length >= width)
            {
                return s;
            }

            int leftPadding = (width - s.Length) / 2;
            int rightPadding = width - s.Length - leftPadding;

            return new string(' ', leftPadding) + s + new string(' ', rightPadding);
        }
    }
}
