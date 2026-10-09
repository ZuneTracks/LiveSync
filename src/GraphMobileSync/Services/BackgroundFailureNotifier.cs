using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;

namespace GraphMobileSync.Services
{
    internal static class BackgroundFailureNotifier
    {
        public static void Notify(string message)
        {
            if (!SyncSettings.Load().FailureNotificationsEnabled)
            {
                return;
            }

            XmlDocument toastXml = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastText02);
            XmlNodeList textNodes = toastXml.GetElementsByTagName("text");
            textNodes.Item(0).InnerText = "LiveSync needs attention";
            textNodes.Item(1).InnerText = message;
            ToastNotificationManager.CreateToastNotifier().Show(new ToastNotification(toastXml));
        }
    }
}
