using System;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.Background;
using Windows.Storage;

namespace GraphMobileSync.Services
{
    internal static class BackgroundSyncRegistration
    {
        private const string TaskName = "LiveSync Graph delta sync";
        private const string RegistrationVersionKey = "BackgroundSyncRegistrationVersion";
        private const int CurrentRegistrationVersion = 1;

        public static bool IsRegistered
        {
            get
            {
                return BackgroundTaskRegistration.AllTasks.Values.Any(registration => registration.Name == TaskName);
            }
        }

        public static async Task RegisterAsync()
        {
            if (IsCurrentRegistration())
            {
                return;
            }

            BackgroundAccessStatus accessStatus = await BackgroundExecutionManager.RequestAccessAsync();

            if (!accessStatus.ToString().StartsWith("Allowed", StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Windows did not allow background execution (" + accessStatus
                    + "). Enable background activity for LiveSync in Settings.");
            }

            if (IsCurrentRegistration())
            {
                return;
            }

            UnregisterExistingTasks();

            BackgroundTaskBuilder builder = new BackgroundTaskBuilder
            {
                Name = TaskName
            };
            builder.SetTrigger(new TimeTrigger(15, false));
            builder.Register();
            ApplicationData.Current.LocalSettings.Values[RegistrationVersionKey] =
                CurrentRegistrationVersion;
        }

        public static void Unregister()
        {
            UnregisterExistingTasks();
            ApplicationData.Current.LocalSettings.Values.Remove(RegistrationVersionKey);
        }

        private static bool IsCurrentRegistration()
        {
            object version;
            return IsRegistered
                && ApplicationData.Current.LocalSettings.Values.TryGetValue(
                    RegistrationVersionKey,
                    out version)
                && version is int
                && (int)version == CurrentRegistrationVersion;
        }

        private static void UnregisterExistingTasks()
        {
            foreach (IBackgroundTaskRegistration registration in BackgroundTaskRegistration.AllTasks.Values)
            {
                if (registration.Name == TaskName)
                {
                    registration.Unregister(false);
                }
            }
        }
    }
}
