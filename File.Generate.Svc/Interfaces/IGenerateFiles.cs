using Project.Generate.Svc.Models;

namespace Project.Generate.Svc.Interfaces
{
    public interface IGenerateFilesService
    {
        void GenerateExcelByInterop(IEnumerable<Client> client);
        void GenerateExcelByClosedXml(IEnumerable<Client> client);
        FileResult GenerateExcelStreamByClosedXml();
        void GenerateCsvFile(IEnumerable<Client> client);
        FileResult GenerateCsvFileStream();
    }
}
