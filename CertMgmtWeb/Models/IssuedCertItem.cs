namespace CertMgmtWeb.Models;

public class IssuedCertItem
{
    public string RequestId { get; set; } = string.Empty;
    public string RequesterName { get; set; } = string.Empty;
    public string CommonName { get; set; } = string.Empty;
    public string San { get; set; } = string.Empty;
}
