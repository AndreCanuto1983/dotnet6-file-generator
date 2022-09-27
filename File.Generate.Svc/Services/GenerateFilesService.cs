using Project.Generate.Svc.Interfaces;
using Project.Generate.Svc.Models;
using Project.Generate.Svc.Util;

namespace Project.Generate.Svc.Services
{
    public class GenerateFilesService : IGenerateFilesService
    {
        private readonly ILogger<GenerateFilesService> _logger;

        public GenerateFilesService(ILogger<GenerateFilesService> logger)
        {
            _logger = logger;
        }

        public string GenerateExcelFile(IEnumerable<Client> client, string path)
        {
            try
            {
                var dataTable = GenerateFiles.GenerateDataTable(client);
                dataTable.SaveExcelFile(new FileInfo(@$"{path}\Client.xlsx"));                
                return @$"{path}\Client.xlsx";
            }
            catch (Exception ex)
            {
                _logger.LogError("[GenerateFilesService][GenerateExcelFile] => EXCEPTION: {ex}", ex.Message);
                return ex.Message;
            }

        }

        public string GenerateCsvFile(IEnumerable<Client> client, string path)
        {
            try
            {
                var dataTable = GenerateFiles.GenerateDataTable(client);
                dataTable.SaveCsvFile(@$"{path}\Client.csv");
                return @$"{path}\Client.csv";
            }
            catch (Exception ex)
            {
                _logger.LogError("[GenerateFilesService][GenerateCsvFile] => EXCEPTION: {ex}", ex.Message);
                return ex.Message;
            }            
        }
    }
}
