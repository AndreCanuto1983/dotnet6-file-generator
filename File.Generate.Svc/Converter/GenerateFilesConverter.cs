using Project.Generate.Svc.Models;

namespace Project.Generate.Svc.Converter
{
    public static class GenerateFilesConverter
    {
        public static FileResult Success(this Stream fileStream, string name)
            => new()
            {
                Success = true,
                File = fileStream,
                Name = name
            };

        public static FileResult Error(this string message)
            => new()
            {
                Success = false,                
                Message = message
            };
    }
}
