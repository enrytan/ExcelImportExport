using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExcelImportExport.Startup))]
namespace ExcelImportExport
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
