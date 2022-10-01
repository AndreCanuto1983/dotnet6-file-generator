namespace Project.Generate.Svc.Models
{
    public class FileStreamResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public Stream File { get; set; }
        public string Name { get; set; }
    }
}
