using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using ComptaFlow.Models;

namespace ComptaFlow.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class HomeController : ControllerBase
    {
        private readonly ILogger<HomeController> _logger;
        
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        // GET api/home
        [HttpGet]
        public IActionResult GetStatus()
        {
            return Ok(new { message = "API ComptaFlow opérationnelle" });
        }

        // GET api/home/privacy
        [HttpGet("privacy")]
        public IActionResult GetPrivacyInfo()
        {
            // Tu peux retourner des données JSON ici
            return Ok(new { policy = "Politique de confidentialité à définir" });
        }

        // GET api/home/error
        [HttpGet("error")]
        public IActionResult GetError()
        {
            var requestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier;
            var errorModel = new ErrorViewModel { RequestId = requestId };
            return Ok(errorModel);
        }
    }
}
