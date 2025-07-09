using System.Web.Http;

namespace AddOnGs.Helpers
{
    public static class AddonMiddlewareRegistration
    {
        public static void RegisterMessageHandlers(HttpConfiguration config)
        {
            // ลงทะเบียน Message Handler ของ Addon
            config.MessageHandlers.Add(new AddOnGs.Handlers.AddonApiVersionRedirectHandler());
        }
    }
}
