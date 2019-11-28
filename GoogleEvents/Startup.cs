using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(GoogleEvents.Startup))]
namespace GoogleEvents
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
