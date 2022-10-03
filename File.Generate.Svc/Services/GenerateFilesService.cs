using Project.Generate.Svc.Converter;
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

        public void GenerateExcelByInterop(IEnumerable<Client> client)
        {
            try
            {
                var dataTable = GenerateFiles.GenerateDataTable(client);

                dataTable.SaveExcelByInterop();
            }
            catch (Exception ex)
            {
                _logger.LogError("[GenerateFilesService][GenerateExcelByInterop] => EXCEPTION: {ex}", ex.Message);
                throw;
            }
        }

        public void GenerateExcelByClosedXml(IEnumerable<Client> client)
        {
            try
            {
                var dataTable = GenerateFiles.GenerateDataTable(client);

                dataTable.SaveExcelByClosedXml();
            }
            catch (Exception ex)
            {
                _logger.LogError("[GenerateFilesService][GenerateExcelByClosedXml] => EXCEPTION: {ex}", ex.Message);
                throw;
            }
        }

        public FileResult GenerateExcelStreamByClosedXml()
        {
            try
            {
                var dataTable = GenerateFiles.GenerateDataTable(GenerateClientList());

                var result = dataTable.SaveExcelStream();
                
                return result.Success("Client.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError("[GenerateFilesService][GenerateExcelStreamByClosedXml] => EXCEPTION: {ex}", ex.Message);

                return ex.Message.Error();
            }
        }

        public void GenerateCsvFile(IEnumerable<Client> client)
        {
            try
            {
                var dataTable = GenerateFiles.GenerateDataTable(client);

                dataTable.SaveCsvFile();
            }
            catch (Exception ex)
            {
                _logger.LogError("[GenerateFilesService][GenerateCsvFile] => EXCEPTION: {ex}", ex.Message);
                throw;
            }
        }

        public FileResult GenerateCsvFileStream()
        {
            try
            {
                var dataTable = GenerateFiles.GenerateDataTable(GenerateClientList());

                var result = dataTable.SaveCsvStream();

                return result.Success("Client.csv");
            }
            catch (Exception ex)
            {
                _logger.LogError("[GenerateFilesService][GenerateCsvFileStream] => EXCEPTION: {ex}", ex.Message);

                return ex.Message.Error();
            }
        }

        private static List<Client> GenerateClientList()
        => new()
            {
                new Client()
                {
                    ClientId = 1,
                    Cpf = "333.222.111-99",
                    Name = "André",
                    Phone = "(11)99999-8855",
                    Email = "andrecanuto@test.com"
                },
                new Client()
                {
                    ClientId = 2,
                    Cpf = "113.222.111-49",
                    Name = "Miguel",
                    Phone = "(11)99999-9900",
                    Email = "miguel@test.com"
                }
            };
    }
}
