using Microsoft.AspNetCore.Mvc;
using Project.Generate.Svc.Interfaces;
using Project.Generate.Svc.Models;
using System.Numerics;

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

        [HttpPost("Excel")]
        [ProducesResponseType(typeof(string), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status500InternalServerError)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status502BadGateway)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status503ServiceUnavailable)]
        public IActionResult GenerateExcel(List<Client> client, string path)
        {
            return Ok(_generateFilesService.GenerateExcelFile(client, path));
        }

        [HttpPost("v2/Excel")]
        [ProducesResponseType(typeof(string), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status401Unauthorized)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status500InternalServerError)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status502BadGateway)]
        [ProducesResponseType(typeof(IActionResult), StatusCodes.Status503ServiceUnavailable)]
        public IActionResult GenerateExcelV2(List<Client> client, string path)
        {
            return Ok(_generateFilesService.GenerateExcelFileClosedXml(client, path));
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
    }
}