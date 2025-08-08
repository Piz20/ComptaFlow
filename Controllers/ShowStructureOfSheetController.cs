using Microsoft.AspNetCore.Mvc;
using Aspose.Cells;
using System.ComponentModel.DataAnnotations;

namespace ComptaFlow.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ShowStructureController : ControllerBase
    {
        [HttpGet("analyser-structure")]
        public IActionResult AnalyserStructureDepuisFichier([FromQuery][Required] string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return BadRequest("❌ Le chemin du fichier est requis.");

            if (!System.IO.File.Exists(filePath))
                return NotFound($"❌ Fichier introuvable : {filePath}");

            try
            {
                var workbook = new Workbook(filePath);

                foreach (var feuille in workbook.Worksheets)
                {
                    ComptaUtils.AfficherStructureGenerale(workbook, feuille.Name);
                }

                return Ok("✅ Structure affichée dans la console avec succès.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"❌ Erreur lors de la lecture du fichier : {ex.Message}");
            }
        }
    }
}
