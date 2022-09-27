using Project.Generate.Svc.Interfaces;
using Project.Generate.Svc.Services;

namespace Project.Generate.Svc.Configurations
{
    public static class DependencyInjections
    {
        public static void DiSettings(this IServiceCollection services)
        {
            services.AddScoped<IGenerateFilesService, GenerateFilesService>();
        }
    }
}
