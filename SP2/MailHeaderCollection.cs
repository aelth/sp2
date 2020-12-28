using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace SP2
{
    // almost verbatim from https://bitbucket.org/otac0n/mailutilities/src/master/MailUtilities/MailHeaderCollection.cs
    class MailHeaderCollection : List<MailHeader>
    {
        // https://bitbucket.org/otac0n/mailutilities/src/master/MailUtilities/MimeMessage.cs
        private const string m_headersMatch = @"^(?<header_key>[-A-Za-z0-9]+)(?<separator>:[ \t]*)(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";

        public MailHeaderCollection(String rawHeaders)
        {
            MatchCollection headerMatches = Regex.Matches(rawHeaders, m_headersMatch, RegexOptions.Multiline);
            foreach (Match headerMatch in headerMatches)
            {
                string headerKey = headerMatch.Groups["header_key"].Value;
                string headerValue = headerMatch.Groups["header_value"].Value;

                this.Add(new MailHeader(headerKey, headerValue));
            }
        }

        public List<String> GetHeader(String headerName)
        {
            List<String> headerList = new List<String>();
            foreach (MailHeader header in this)
            {
                if (header.Key.ToUpperInvariant() == headerName.ToUpperInvariant())
                {
                    headerList.Add(header.Value);
                }
            }

            return headerList;
        }

        public string GetHeaderString(string headerName)
        {
            StringBuilder headerString = new StringBuilder(100);
            List<string> header = GetHeader(headerName);
            foreach (var headerPart in header)
            {
                headerString.AppendFormat("{0},", headerPart);
            }

            return headerString.ToString().TrimEnd(new char[] { ',' });
        }

        public override String ToString()
        {
            String rawHeasders = "";
            foreach (MailHeader header in this)
            {
                rawHeasders += header.ToString() + "\r\n";
            }

            return rawHeasders;
        }
    }
}
