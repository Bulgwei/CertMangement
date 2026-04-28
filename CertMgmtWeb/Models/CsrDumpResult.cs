namespace CertMgmtWeb.Models;

public class CsrDumpResult
{
    public bool SignatureMatchesPublicKey { get; set; }
    public string SubjectCommonName { get; set; } = string.Empty;
    public string SubjectOrganization { get; set; } = string.Empty;
    public string SubjectOrganizationalUnit { get; set; } = string.Empty;
    public string SubjectLocality { get; set; } = string.Empty;
    public string SubjectState { get; set; } = string.Empty;
    public string SubjectStreet { get; set; } = string.Empty;
    public string SubjectCountry { get; set; } = string.Empty;
    public string SubjectEmail { get; set; } = string.Empty;
    public int KeyLength { get; set; }
    public string KeyAlgorithm { get; set; } = string.Empty;
    public string SignatureAlgorithm { get; set; } = string.Empty;
    public List<string> SanEntries { get; set; } = new();
    public List<string> EnhancedKeyUsage { get; set; } = new();
    public List<string> Extensions { get; set; } = new();
    public List<string> TemplateNames { get; set; } = new();
}
