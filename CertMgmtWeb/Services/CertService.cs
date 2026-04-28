using System.Diagnostics;
using System.Text;
using CertMgmtWeb.Models;

namespace CertMgmtWeb.Services;

public class CertService
{
    public async Task<List<CAInfo>> GetEnrollmentCAsAsync()
    {
        var result = await RunProcessAsync("certutil", "-adca");
        var lines = result.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var list = new List<CAInfo>();
        CAInfo? current = null;

        foreach (var line in lines)
        {
            if (line.Contains("cn =", StringComparison.OrdinalIgnoreCase))
            {
                current ??= new CAInfo();
                current.CaName = line.Split('=', 2)[1].Trim();
            }
            else if (line.Contains("dNSHostName =", StringComparison.OrdinalIgnoreCase))
            {
                current ??= new CAInfo();
                current.DnsHostName = line.Split('=', 2)[1].Trim();
            }
            else if (line.Contains("CAIsValid: 1", StringComparison.OrdinalIgnoreCase) || line.Contains("CertUtil: -ADCA command completed successfully", StringComparison.OrdinalIgnoreCase))
            {
                if (current != null && !string.IsNullOrWhiteSpace(current.CaName))
                {
                    list.Add(current);
                    current = null;
                }
            }
        }

        if (current != null && !string.IsNullOrWhiteSpace(current.CaName))
        {
            list.Add(current);
        }

        return list;
    }

    public async Task<List<string>> GetCATemplatesAsync(string targetCaConfig)
    {
        var result = await RunProcessAsync("certutil", $"-config \"{targetCaConfig}\" -catemplates");
        var lines = result.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        return lines
            .Where(line => !line.Contains("CertUtil", StringComparison.OrdinalIgnoreCase))
            .Select(line => line.Split(':')[0].Trim())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .ToList();
    }

    public async Task<object> GetCAStatusAsync(string targetCaConfig)
    {
        var ping = await RunProcessAsync("certutil", $"-config \"{targetCaConfig}\" -ping");
        var pingAdmin = await RunProcessAsync("certutil", $"-config \"{targetCaConfig}\" -pingadmin");
        return new
        {
            IsResponding = !ping.Contains("FAILED", StringComparison.OrdinalIgnoreCase),
            HasAdmin = !pingAdmin.Contains("FAILED", StringComparison.OrdinalIgnoreCase),
            PingOutput = ping,
            PingAdminOutput = pingAdmin,
        };
    }

    public async Task<string> SaveUploadedFileAsync(IFormFile file)
    {
        var tempFolder = Path.Combine(Path.GetTempPath(), "CertMgmtWeb", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempFolder);
        var filePath = Path.Combine(tempFolder, Path.GetFileName(file.FileName));

        await using var fileStream = File.Create(filePath);
        await file.CopyToAsync(fileStream);
        return filePath;
    }

    public async Task<CsrDumpResult?> DumpRequestAsync(string reqFileName)
    {
        var output = await RunProcessAsync("certutil", $"\"{reqFileName}\"");
        if (string.IsNullOrWhiteSpace(output))
        {
            return null;
        }

        var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var result = new CsrDumpResult();
        var subjectSection = false;

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (line.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
            {
                subjectSection = true;
                continue;
            }

            if (subjectSection)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    subjectSection = false;
                }
                else if (line.Contains("= "))
                {
                    var parts = line.Split('=', 2);
                    var key = parts[0].Trim();
                    var value = parts[1].Trim();
                    switch (key)
                    {
                        case "CN": result.SubjectCommonName = value; break;
                        case "O": result.SubjectOrganization = value; break;
                        case "OU": result.SubjectOrganizationalUnit = value; break;
                        case "L": result.SubjectLocality = value; break;
                        case "S": result.SubjectState = value; break;
                        case "ST": result.SubjectState = value; break;
                        case "STREET": result.SubjectStreet = value; break;
                        case "C": result.SubjectCountry = value; break;
                        case "E": result.SubjectEmail = value; break;
                    }
                }
            }

            if (line.StartsWith("Public Key Length:", StringComparison.OrdinalIgnoreCase))
            {
                var parts = line.Split(':', 2);
                if (parts.Length > 1 && int.TryParse(parts[1].Trim().Split(' ')[0], out var keyLen))
                {
                    result.KeyLength = keyLen;
                }
            }
            else if (line.StartsWith("Public Key Algorithm:", StringComparison.OrdinalIgnoreCase))
            {
                result.KeyAlgorithm = line.Split(':', 2)[1].Trim();
            }
            else if (line.StartsWith("Signature Algorithm:", StringComparison.OrdinalIgnoreCase))
            {
                result.SignatureAlgorithm = line.Split(':', 2)[1].Trim();
            }
            else if (line.Contains("Signature matches Public Key", StringComparison.OrdinalIgnoreCase))
            {
                result.SignatureMatchesPublicKey = true;
            }
            else if (line.Contains("DNS Name=", StringComparison.OrdinalIgnoreCase) || line.Contains("IP Address=", StringComparison.OrdinalIgnoreCase) || line.Contains("Other Name=", StringComparison.OrdinalIgnoreCase))
            {
                result.SanEntries.Add(line);
            }
            else if (line.Contains("2.5.29.37", StringComparison.OrdinalIgnoreCase) || line.Contains("Enhanced Key Usage", StringComparison.OrdinalIgnoreCase))
            {
                result.EnhancedKeyUsage.Add(line);
            }
            else if (line.Contains("1.3.6.1.4.1.311.21.7", StringComparison.OrdinalIgnoreCase) || line.Contains("CertificateTemplate", StringComparison.OrdinalIgnoreCase))
            {
                result.TemplateNames.Add(line);
            }
        }

        return result;
    }

    public async Task<string> SubmitRequestAsync(string targetCaConfig, string templateName, string requestFile)
    {
        var requestFilePath = Path.GetFullPath(requestFile);
        var folder = Path.GetDirectoryName(requestFilePath) ?? Path.GetTempPath();
        var baseName = Path.GetFileNameWithoutExtension(requestFilePath);
        var certFile = Path.Combine(folder, baseName + ".cer");
        var p7bFile = Path.Combine(folder, baseName + ".p7b");
        var args = $"-f -config \"{targetCaConfig}\" -submit -attrib \"certificatetemplate:{templateName}\" \"{requestFilePath}\" \"{certFile}\" \"{p7bFile}\"";
        return await RunProcessAsync("certreq", args);
    }

    public async Task<string> RetrieveCertAsync(string targetCaConfig, string requestId, string folderName)
    {
        var folder = Path.GetFullPath(folderName);
        Directory.CreateDirectory(folder);
        var certFile = Path.Combine(folder, $"CertNew-RequestID-{requestId}.cer");
        var p7bFile = Path.Combine(folder, $"CertNew-RequestID-{requestId}.p7b");
        var args = $"-f -retrieve -config \"{targetCaConfig}\" {requestId} \"{certFile}\" \"{p7bFile}\"";
        return await RunProcessAsync("certreq", args);
    }

    public async Task<List<IssuedCertItem>> GetIssuedCertListAsync(string targetCaConfig, string? requesterName, int timeWindow)
    {
        var fromDate = DateTime.Now.AddDays(-timeWindow).ToString("MM/dd/yyyy");
        var filter = requesterName is null ? $"disposition=20,notbefore>={fromDate}" : $"disposition=20,RequesterName={requesterName},notbefore>={fromDate}";
        var args = $"-config \"{targetCaConfig}\" -view -restrict \"{filter}\" -out \"RequestID, Requestername, CommonName\" csv";
        var output = await RunProcessAsync("certutil", args);
        var rows = ParseCsvRows(output);
        var result = new List<IssuedCertItem>();

        foreach (var row in rows.Skip(1))
        {
            if (row.Length < 3) continue;
            result.Add(new IssuedCertItem
            {
                RequestId = row[0],
                RequesterName = row[1],
                CommonName = row[2],
                San = await GetSanForRequestAsync(targetCaConfig, row[0])
            });
        }

        return result;
    }

    public async Task<List<IssuedCertItem>> RetrievePendingCertsAsync(string targetCaConfig)
    {
        var args = $"-config \"{targetCaConfig}\" -view -restrict \"disposition=9\" -out \"RequestID, Requestername, CommonName, RawRequest\" csv";
        var output = await RunProcessAsync("certutil", args);
        var rows = ParseCsvRows(output);
        var result = new List<IssuedCertItem>();

        foreach (var row in rows.Skip(1))
        {
            if (row.Length < 3) continue;
            var requestId = row[0];
            var tempPath = Path.Combine(Path.GetTempPath(), "CertMgmtWeb", "Pending", requestId);
            Directory.CreateDirectory(tempPath);
            var rawFileName = Path.Combine(tempPath, "Request.RawRequest");
            await RunProcessAsync("certutil", $"-config \"{targetCaConfig}\" -view -restrict \"requestid={requestId}\" -out Request.RawRequest", tempPath);

            var sanEntries = File.Exists(rawFileName)
                ? ParseSanEntries(await RunProcessAsync("certutil", $"\"{rawFileName}\""))
                : new List<string>();

            result.Add(new IssuedCertItem
            {
                RequestId = requestId,
                RequesterName = row[1],
                CommonName = row[2],
                San = string.Join("; ", sanEntries)
            });
        }

        return result;
    }

    public async Task<string?> DisplayRawDataAsync(string targetCaConfig, string requestId, string mode)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), "CertMgmtWeb", "Raw", requestId);
        Directory.CreateDirectory(tempPath);
        var rawFileName = Path.Combine(tempPath, mode.Equals("Retrieval", StringComparison.OrdinalIgnoreCase) ? "RawCertificate.cer" : "Request.RawRequest");
        var outType = mode.Equals("Retrieval", StringComparison.OrdinalIgnoreCase) ? "rawcertificate" : "Request.RawRequest";

        await RunProcessAsync("certutil", $"-config \"{targetCaConfig}\" -view -restrict \"requestid={requestId}\" -out {outType}", tempPath);
        if (!File.Exists(rawFileName))
        {
            return null;
        }

        return await RunProcessAsync("certutil", $"\"{rawFileName}\"");
    }

    public async Task<string> IssueCertRequestAsync(string targetCaConfig, string requestId)
    {
        return await RunProcessAsync("certutil", $"-config \"{targetCaConfig}\" -resubmit {requestId}");
    }

    public async Task<string> DenyCertRequestAsync(string targetCaConfig, string requestId)
    {
        return await RunProcessAsync("certutil", $"-config \"{targetCaConfig}\" -deny {requestId}");
    }

    private async Task<string> GetSanForRequestAsync(string targetCaConfig, string requestId)
    {
        var args = $"-config \"{targetCaConfig}\" -view -restrict \"ExtensionRequestId={requestId}, ExtensionName=2.5.29.17\" ext csv";
        var output = await RunProcessAsync("certutil", args);
        var rows = ParseCsvRows(output);
        var sanItems = new List<string>();

        foreach (var row in rows.Skip(1))
        {
            if (row.Length > 0 && !string.IsNullOrWhiteSpace(row[0]))
            {
                sanItems.Add(row[0]);
            }
        }

        return string.Join("|", sanItems);
    }

    private static List<string[]> ParseCsvRows(string csvText)
    {
        var result = new List<string[]>();
        if (string.IsNullOrWhiteSpace(csvText))
        {
            return result;
        }

        var lines = csvText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var line in lines)
        {
            var values = line.Split(',').Select(value => value.Trim(' ', '"')).ToArray();
            result.Add(values);
        }

        return result;
    }

    private static List<string> ParseSanEntries(string dump)
    {
        var results = new List<string>();
        if (string.IsNullOrWhiteSpace(dump))
        {
            return results;
        }

        var lines = dump.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var line in lines)
        {
            if (line.Contains("DNS Name=", StringComparison.OrdinalIgnoreCase) ||
                line.Contains("IP Address=", StringComparison.OrdinalIgnoreCase) ||
                line.Contains("Other Name=", StringComparison.OrdinalIgnoreCase))
            {
                results.Add(line.Trim());
            }
        }

        return results;
    }

    private async Task<string> RunProcessAsync(string fileName, string arguments, string? workingDirectory = null)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = workingDirectory ?? Environment.CurrentDirectory,
        };

        using var process = Process.Start(startInfo);
        if (process == null)
        {
            return string.Empty;
        }

        var output = await process.StandardOutput.ReadToEndAsync();
        var error = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        return string.IsNullOrWhiteSpace(output) ? error : output + Environment.NewLine + error;
    }
}
