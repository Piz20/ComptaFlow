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

        // Injection possible via DI, sinon création manuelle par défaut
        public SageTCDController(TCDGeneratorSage? tcdService = null)
        {
            _tcdService = tcdService ?? new TCDGeneratorSage();
        }

        [HttpPost("generer-tcd")]
        public IActionResult GenererTCDDepuisFichier([FromBody] TCDRequest request)
        {
            // Validation des paramètres
            if (request == null)
                return BadRequest("❌ Le corps de la requête est vide.");

            if (string.IsNullOrWhiteSpace(request.FilePath) || !System.IO.File.Exists(request.FilePath))
                return BadRequest("❌ Le fichier source est introuvable.");

            if (string.IsNullOrWhiteSpace(request.OutputDirectory) || !Directory.Exists(request.OutputDirectory))
                return BadRequest("❌ Le répertoire de sortie est invalide.");

            try
            {
                var outputPath = Path.Combine(request.OutputDirectory, "FICHIER SAGE AVEC TCD.xlsx");

                _tcdService.GenererTCDAvecFeuilPrecedente(request.FilePath, outputPath);

                return Ok(new
                {
                    message = "✅ Fichier généré avec succès.",
                    cheminFichierGenere = outputPath
                });
            }
            catch (Exception ex)
            {
                // Log l’erreur ici si possible
                return StatusCode(500, new
                {
                    message = "❌ Une erreur est survenue lors de la génération du fichier.",
                    details = ex.Message
                });
            }
        }
    }
}
