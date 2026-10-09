using System;
using System.Globalization;
using Windows.Storage;

namespace GraphMobileSync.Services
{
    internal sealed class GraphThrottledException : Exception
    {
        public GraphThrottledException(TimeSpan retryAfter)
            : base(string.Format(
                "Microsoft Graph is throttling requests. The next check will wait until {0}.",
                DateTimeOffset.UtcNow.Add(retryAfter).ToLocalTime().ToString("g")))
        {
            RetryAfter = retryAfter;
        }

        public TimeSpan RetryAfter { get; private set; }
    }

    internal static class GraphThrottleState
    {
        private const string DeferredUntilKey = "MicrosoftGraphDeferredUntilUtc";

        public static bool TryGetRemainingDelay(out TimeSpan remainingDelay)
        {
            remainingDelay = TimeSpan.Zero;
            ApplicationDataContainer settings = ApplicationData.Current.LocalSettings;
            string serializedTime = settings.Values[DeferredUntilKey] as string;

            if (string.IsNullOrEmpty(serializedTime))
            {
                return false;
            }

            DateTimeOffset deferredUntil;

            if (!DateTimeOffset.TryParse(
                    serializedTime,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out deferredUntil))
            {
                settings.Values.Remove(DeferredUntilKey);
                return false;
            }

            remainingDelay = deferredUntil - DateTimeOffset.UtcNow;

            if (remainingDelay <= TimeSpan.Zero)
            {
                settings.Values.Remove(DeferredUntilKey);
                remainingDelay = TimeSpan.Zero;
                return false;
            }

            return true;
        }

        public static void Defer(TimeSpan retryAfter)
        {
            DateTimeOffset requestedUntil = DateTimeOffset.UtcNow.Add(
                retryAfter > TimeSpan.Zero ? retryAfter : TimeSpan.FromSeconds(60));
            DateTimeOffset existingUntil;
            TimeSpan ignoredDelay;

            if (TryGetRemainingDelay(out ignoredDelay)
                && TryGetDeferredUntil(out existingUntil)
                && existingUntil > requestedUntil)
            {
                requestedUntil = existingUntil;
            }

            ApplicationData.Current.LocalSettings.Values[DeferredUntilKey] =
                requestedUntil.UtcDateTime.ToString("o", CultureInfo.InvariantCulture);
        }

        private static bool TryGetDeferredUntil(out DateTimeOffset deferredUntil)
        {
            string serializedTime =
                ApplicationData.Current.LocalSettings.Values[DeferredUntilKey] as string;

            return DateTimeOffset.TryParse(
                serializedTime,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out deferredUntil);
        }
    }
}
