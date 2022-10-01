using Microsoft.AspNetCore.Mvc;
using Project.Generate.Svc.Interfaces;
using Project.Generate.Svc.Models;

namespace Project.Generate.Svc.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class GenerateFilesController : ControllerBase
    {
        private readonly IGenerateFilesService _generateFilesService;

        public GenerateFilesController(IGenerateFilesService generateFilesService)
        {
            _generateFilesService = generateFilesService;
        }

        /// <summary>
        /// Lib Microsoft, you need excel installed on the machine.
        /// </summary>
        /// <param name="client"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        [HttpPost("ExcelByInterop")]
        [ProducesResponseType(typeof(string), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status500InternalServerError)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status502BadGateway)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status503ServiceUnavailable)]
        public IActionResult GenerateExcelByInterop(List<Client> client, string path)
        {
            return Ok(_generateFilesService.GenerateExcelByInterop(client, path));
        }

        [HttpPost("ExcelByClosedXml")]
        [ProducesResponseType(typeof(string), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status500InternalServerError)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status502BadGateway)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status503ServiceUnavailable)]
        public IActionResult GenerateExcelByClosedXml(List<Client> client, string path)
        {
            return Ok(_generateFilesService.GenerateExcelByClosedXml(client, path));
        }

        /// <summary>
        /// Example of how to download the file via javascript
        /// window.open('http://localhost:5000/api/GenerateFiles/ExcelStreamFileByClosedXml', "_blank");
        /// </summary>
        /// <param name="client"></param>
        /// <returns></returns>
        [HttpGet("ExcelStreamFileByClosedXml")]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status500InternalServerError)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status502BadGateway)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status503ServiceUnavailable)]
        public IActionResult GenerateExcelStreamFileByClosedXml()
        {
            var response = _generateFilesService.GenerateExcelStreamByClosedXml();

            return File(response.File, "application/octet-stream", response.Name);
        }

        [HttpPost("Csv")]
        [ProducesResponseType(typeof(string), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status500InternalServerError)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status502BadGateway)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status503ServiceUnavailable)]
        public IActionResult GenerateCsv(List<Client> client, string path)
        {
            return Ok(_generateFilesService.GenerateCsvFile(client, path));
        }

        /// <summary>
        /// Example of how to download the file via javascript
        /// window.open('http://localhost:5000/api/GenerateFiles/CsvStream', "_blank");
        /// </summary>
        /// <returns></returns>
        [HttpGet("CsvStream")]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status500InternalServerError)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status502BadGateway)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status503ServiceUnavailable)]
        public IActionResult GenerateCsvFileStream()
        {
            var response = _generateFilesService.GenerateCsvFileStream();

            return File(response.File, "application/octet-stream", response.Name);
        }
    }
}