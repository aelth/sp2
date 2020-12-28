using System;

namespace SP2
{
    class MailHeader
    {
        public MailHeader(string key, string value)
        {
            Key = key;
            Value = value;
        }

        public string Key
        {
            get;
            private set;
        }

        public string Value
        {
            get;
            set;
        }

        public override string ToString()
        {
            return String.Format("{0}: {1}", Key, Value);
        }
    }
}
