using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExcelFileRead.Startup))]
namespace ExcelFileRead
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
