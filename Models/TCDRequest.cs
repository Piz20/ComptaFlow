
namespace ComptaFlow.Models
{
    /// <summary>
    /// Représente une requête pour générer un TCD à partir d'un fichier Excel.
    /// </summary>
    public class TCDRequest
    {
        public string FilePath { get; set; } = string.Empty;
        public string OutputDirectory { get; set; } = string.Empty;
    }
}