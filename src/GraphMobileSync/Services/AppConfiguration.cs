using GraphMobileSync.Configuration;

namespace GraphMobileSync.Services
{
    internal static class AppConfiguration
    {
        public static string ClientId
        {
            get
            {
                return LiveSyncConfiguration.MicrosoftEntraClientId;
            }
        }

        public static string Tenant
        {
            get
            {
                return LiveSyncConfiguration.MicrosoftEntraTenant;
            }
        }
    }
}
