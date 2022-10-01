using Project.Generate.Svc.Models;

namespace Project.Generate.Svc.Converter
{
    public static class GenerateFilesConverter
    {
        public static FileStreamResult Success(this Stream fileStream, string name)
            => new()
            {
                Success = true,
                File = fileStream,
                Name = name
            };

        public static FileStreamResult Error(this string message)
            => new()
            {
                Success = false,                
                Message = message
            };
    }
}
