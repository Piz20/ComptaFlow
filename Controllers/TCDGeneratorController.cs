using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using ComptaFlow.Services;
using ComptaFlow.Models;

namespace ComptaFlow.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SageTCDController : ControllerBase
    {
        private readonly TCDGeneratorSage _tcdService;

        // Injection possible, sinon création manuelle
        public SageTCDController()
        {
            _tcdService = new TCDGeneratorSage();
        }

        [HttpPost("generer-tcd")]
        public IActionResult GenererTCDDepuisFichier([FromBody] TCDRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.FilePath) || !System.IO.File.Exists(request.FilePath))
                return BadRequest("❌ Le fichier source est introuvable.");

            if (string.IsNullOrWhiteSpace(request.OutputDirectory) || !Directory.Exists(request.OutputDirectory))
                return BadRequest("❌ Le répertoire de sortie est invalide.");

            try
            {
                var outputPath = Path.Combine(request.OutputDirectory, "FICHIER SAGE AVEC TCD.xlsx");

                _tcdService.GenererTCDPourChaqueFeuille(request.FilePath, outputPath);

                return Ok(new
                {
                    message = "✅ Fichier généré avec succès.",
                    cheminFichierGenere = outputPath
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"❌ Erreur : {ex.Message}");
            }
        }
    }
}
