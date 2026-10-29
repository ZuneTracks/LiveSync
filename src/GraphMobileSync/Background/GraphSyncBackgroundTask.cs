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
                SyncSettings settings = SyncSettings.Load();
                bool outboundWritesAvailable = false;
                string accessToken;

                if (settings.TwoWaySyncEnabled)
                {
                    try
                    {
                        LiveSyncProEntitlement entitlement =
                            await LiveSyncProService.GetEntitlementAsync();

                        if (entitlement.IsActive)
                        {
                            accessToken = await OAuthService.GetBackgroundAccessTokenAsync(
                                requestWriteAccess: true);
                            outboundWritesAvailable = true;
                        }
                        else
                        {
                            BackgroundSyncDiagnostics.Record(
                                "LiveSync Pro is not active; background sync will not upload local edits.");
                            accessToken = await OAuthService.GetBackgroundAccessTokenAsync();
                        }
                    }
                    catch (Exception exception)
                    {
                        BackgroundSyncDiagnostics.Record(
                            "Background two-way sync is unavailable; continuing with read-only sync: "
                            + exception.Message);
                        accessToken = await OAuthService.GetBackgroundAccessTokenAsync();
                    }
                }
                else
                {
                    accessToken = await OAuthService.GetBackgroundAccessTokenAsync();
                }

                DeltaSyncResult result;

                using (GraphClient graphClient = new GraphClient(accessToken))
                {
                    result = await new DeltaSyncService(graphClient).PollAndApplyAsync(
                        outboundWritesAvailable: outboundWritesAvailable);
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
