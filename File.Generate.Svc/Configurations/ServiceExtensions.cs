using System.Text.Json.Serialization;

namespace Project.Generate.Svc.Configurations
{
    public static class ServiceExtensions
    {
        public static void ServiceExtensionSettings(this IServiceCollection services)
        {
            services.AddControllers()
                    .AddJsonOptions(options =>
                    {
                        options.JsonSerializerOptions.PropertyNameCaseInsensitive = true;
                        options.JsonSerializerOptions.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull;
                    });
        }
    }
}
