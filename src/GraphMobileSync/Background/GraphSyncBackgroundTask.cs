using System;
using GraphMobileSync.Services;
using System.Threading.Tasks;

namespace GraphMobileSync.Background
{
    internal static class GraphSyncBackgroundTask
    {
        public static async Task RunAsync()
        {
            try
            {
                BackgroundSyncDiagnostics.Record("Checking Microsoft Graph for changes...");
                string accessToken = await OAuthService.GetBackgroundAccessTokenAsync();
                DeltaSyncResult result;

                using (GraphClient graphClient = new GraphClient(accessToken))
                {
                    result = await new DeltaSyncService(graphClient).PollAndApplyAsync();
                }

                BackgroundSyncDiagnostics.Record(result.Message);
            }
            catch (Exception exception)
            {
                string message = "Background sync failed: " + exception.Message;
                BackgroundSyncDiagnostics.Record(message);
                BackgroundFailureNotifier.Notify(message);
            }
        }
    }
}
