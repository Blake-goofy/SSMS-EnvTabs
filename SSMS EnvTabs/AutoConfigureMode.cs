using System;

namespace SSMS_EnvTabs
{
    internal static class AutoConfigureMode
    {
        internal const string Server = "server";
        internal const string ServerDatabase = "server db";
        internal const string Off = "off";

        internal static string Normalize(string value)
        {
            value = value?.Trim();

            if (string.Equals(value, Server, StringComparison.OrdinalIgnoreCase))
            {
                return Server;
            }

            if (string.Equals(value, ServerDatabase, StringComparison.OrdinalIgnoreCase))
            {
                return ServerDatabase;
            }

            if (string.Equals(value, Off, StringComparison.OrdinalIgnoreCase))
            {
                return Off;
            }

            return ServerDatabase;
        }

        internal static bool IsEnabled(string value)
        {
            return !string.Equals(Normalize(value), Off, StringComparison.Ordinal);
        }

        internal static bool UsesDatabase(string value)
        {
            return string.Equals(Normalize(value), ServerDatabase, StringComparison.Ordinal);
        }
    }
}
