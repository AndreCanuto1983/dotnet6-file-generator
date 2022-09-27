using Project.Generate.Svc.Models;

namespace Project.Generate.Svc.Interfaces
{
    public interface IGenerateFilesService
    {
        string GenerateExcelFile(IEnumerable<Client> client, string path);
        string GenerateExcelFileClosedXml(IEnumerable<Client> client, string path);
        string GenerateCsvFile(IEnumerable<Client> client, string path);
    }
}
