namespace CertMgmtWeb.Models;

public class CAInfo
{
    public string CaName { get; set; } = string.Empty;
    public string DnsHostName { get; set; } = string.Empty;
    public string ConfigString => string.IsNullOrWhiteSpace(DnsHostName) ? CaName : $"{DnsHostName}\\{CaName}";
}
