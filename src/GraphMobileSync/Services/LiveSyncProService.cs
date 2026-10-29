using System;
using System.Linq;
using System.Threading.Tasks;
using Windows.Services.Store;

namespace GraphMobileSync.Services
{
    internal sealed class LiveSyncProEntitlement
    {
        public bool IsActive { get; set; }

        public DateTimeOffset ExpirationDate { get; set; }

        public string StatusMessage { get; set; }
    }

    internal sealed class LiveSyncProPurchaseResult
    {
        public bool IsSuccessful { get; set; }

        public string Message { get; set; }
    }

    internal static class LiveSyncProService
    {
        internal const string ProductStoreId = "9PCFJKMQN57S";

        public static async Task<LiveSyncProEntitlement> GetEntitlementAsync()
        {
            StoreContext context = StoreContext.GetDefault();
            StoreAppLicense appLicense = await context.GetAppLicenseAsync();

            if (appLicense == null)
            {
                return new LiveSyncProEntitlement
                {
                    StatusMessage =
                        "Microsoft Store did not return license information. Check the Store connection and try again."
                };
            }

            StoreLicense license = appLicense.AddOnLicenses
                .Select(entry => entry.Value)
                .FirstOrDefault(candidate =>
                    !string.IsNullOrEmpty(candidate.SkuStoreId)
                    && candidate.SkuStoreId.StartsWith(
                        ProductStoreId,
                        StringComparison.OrdinalIgnoreCase));

            if (license == null)
            {
                return new LiveSyncProEntitlement
                {
                    StatusMessage =
                        "No LiveSync Pro entitlement was found for the Microsoft Store account signed in on this device."
                };
            }

            if (!license.IsActive)
            {
                return new LiveSyncProEntitlement
                {
                    StatusMessage = "The LiveSync Pro subscription is not active."
                };
            }

            return new LiveSyncProEntitlement
            {
                IsActive = true,
                ExpirationDate = license.ExpirationDate,
                StatusMessage = "LiveSync Pro is active."
            };
        }

        public static async Task<LiveSyncProPurchaseResult> PurchaseAsync()
        {
            StoreContext context = StoreContext.GetDefault();
            StorePurchaseResult result = await context.RequestPurchaseAsync(ProductStoreId);

            switch (result.Status)
            {
                case StorePurchaseStatus.Succeeded:
                    return new LiveSyncProPurchaseResult
                    {
                        IsSuccessful = true,
                        Message = "LiveSync Pro is active."
                    };

                case StorePurchaseStatus.AlreadyPurchased:
                    return new LiveSyncProPurchaseResult
                    {
                        IsSuccessful = true,
                        Message = "LiveSync Pro is already active."
                    };

                case StorePurchaseStatus.NotPurchased:
                    return new LiveSyncProPurchaseResult
                    {
                        Message = "The LiveSync Pro purchase was cancelled."
                    };

                case StorePurchaseStatus.NetworkError:
                    return new LiveSyncProPurchaseResult
                    {
                        Message = "The Store could not complete the purchase because of a network error."
                    };

                default:
                    int errorCode = result.ExtendedError == null
                        ? 0
                        : result.ExtendedError.HResult;
                    return new LiveSyncProPurchaseResult
                    {
                        Message = string.Format(
                            "The Store could not complete the LiveSync Pro purchase ({0}, 0x{1:X8}).",
                            result.Status,
                            errorCode)
                    };
            }
        }
    }
}
