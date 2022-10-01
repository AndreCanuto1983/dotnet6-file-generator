using Project.Generate.Svc.Models;

namespace Project.Generate.Svc.Interfaces
{
    public interface IGenerateFilesService
    {
        string GenerateExcelByInterop(IEnumerable<Client> client, string path);
        string GenerateExcelByClosedXml(IEnumerable<Client> client, string path);
        FileStreamResult GenerateExcelStreamByClosedXml();
        string GenerateCsvFile(IEnumerable<Client> client, string path);
        FileStreamResult GenerateCsvFileStream();
    }
}
