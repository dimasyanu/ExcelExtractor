using ExcelExtractor.Repositories;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ViewEngines;
using Newtonsoft.Json;

namespace ExcelExtractor.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class HomeController : Controller
    {

        private readonly Serilog.ILogger _logger;
        private readonly ImportRepository _repo;

        public HomeController(ImportRepository repo)
        {
            _logger = Serilog.Log.Logger;
            _repo = repo;
        }

        public ViewResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            var lowerFileName = Path.GetExtension(file.FileName).ToLowerInvariant();
            var extension = lowerFileName.Split('.').Last();
            if (!new[] { "xls", "xlsx" }.Contains(extension)) {
                return Problem("File not supported");
            }

            try {
                var detail = JsonConvert.DeserializeObject<ImportParam>(data.Detail);
                _repo.ImportExcelFile(data.File, new ImportDetail { Data = null, Detail = detail });

                return Ok("Processing");
            }
            catch (Exception ex) {
                _logger.Error(ex.Message);
                _logger.Error(ex.StackTrace);
                return Problem("Import Failed!");
            }
        }
    }
}