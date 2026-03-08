using ExcelDataReader;
using ExcelProcessor.Api.Models;
using System.Text;

namespace ExcelProcessor.Api.Services;

public sealed class ExcelMatchService : IExcelMatchService
{
    private static readonly string[] SourceOutputHeaders =
    {
        "SourceRowNumber",
        "Batch No.:",
        "Description:",
        "Entry No.:",
        "Invoice Description",
        "Vendor:",
        "Document Number:",
        "Document Type:",
        "PO Number:",
        "Document Date:",
        "Posting Date:",
        "Year - Period:",
        "Order Number:",
        "Account Set:",
        "Tax Group:",
        "Exchange Rate:",
        "Terms:",
        "Due Date:",
        "G/L Account",
        "Account Description",
        "Detail Desc/ Tax Auth",
        "Net Dist. Amt.",
        "Dist. Tax.",
        "Inv Total:"
    };

    public MatchResponse Match(IFormFile sourceFile, IFormFile targetFile)
    {
        var result = Compute(sourceFile, targetFile);
        return result.Summary;
    }

    public MatchCsvReport BuildCsvReport(IFormFile sourceFile, IFormFile targetFile)
    {
        var result = Compute(sourceFile, targetFile);
        var csv = BuildCsv(result.CsvRows, result.TargetHeaders);
        var fileName = $"match-report-{DateTime.UtcNow:yyyyMMdd-HHmmss}.csv";
        return new MatchCsvReport(Encoding.UTF8.GetBytes(csv), fileName);
    }

    public IReadOnlyList<string> GetNonEmptyTargetColumns(IFormFile targetFile)
    {
        using var stream = targetFile.OpenReadStream();
        using var reader = ExcelReaderFactory.CreateReader(stream);

        if (!reader.Read())
        {
            throw new InvalidOperationException("Target file is empty or missing a header row.");
        }

        var columns = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < reader.FieldCount; i += 1)
        {
            var header = reader.GetValue(i)?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(header))
            {
                continue;
            }

            if (seen.Add(header))
            {
                columns.Add(header);
            }
        }

        return columns;
    }

    private static MatchComputationResult Compute(IFormFile sourceFile, IFormFile targetFile)
    {
        var startedAt = DateTime.UtcNow;

        var targetData = ReadTargetData(targetFile);
        var sourceData = ReadSourceData(sourceFile);

        var csvRows = new List<CsvRow>(capacity: targetData.Records.Count);
        var vendorMatchedRows = 0;
        var fullyMatchedRows = 0;

        foreach (var targetRecord in targetData.Records)
        {
            sourceData.RecordsByVendor.TryGetValue(targetRecord.Vendor, out var sourceRowsForVendor);
            sourceRowsForVendor ??= new List<SourceRecord>();

            var vendorMatched = !string.IsNullOrEmpty(targetRecord.Vendor) && sourceRowsForVendor.Count > 0;
            if (vendorMatched)
            {
                vendorMatchedRows += 1;
            }

            var matchedSourceRows = new List<SourceRecord>();
            if (vendorMatched && !string.IsNullOrEmpty(targetRecord.Invoice))
            {
                matchedSourceRows = sourceRowsForVendor
                    .Where(r => IsInvoiceMatch(r.DocumentNumber, targetRecord.Invoice))
                    .ToList();
            }

            if (matchedSourceRows.Count > 0)
            {
                fullyMatchedRows += 1;
            }

            csvRows.Add(new CsvRow(targetRecord.RowNumber, targetRecord.TargetValues, null));

            foreach (var sourceRecord in matchedSourceRows)
            {
                csvRows.Add(new CsvRow(targetRecord.RowNumber, targetRecord.TargetValues, sourceRecord));
            }
        }

        var elapsedMs = (int)(DateTime.UtcNow - startedAt).TotalMilliseconds;
        var summary = new MatchResponse(
            sourceData.SourceRows,
            targetData.TargetRows,
            targetData.ValidTargetVendorRows,
            targetData.ValidTargetCompositeRows,
            vendorMatchedRows,
            fullyMatchedRows,
            vendorMatchedRows - fullyMatchedRows,
            elapsedMs
        );

        return new MatchComputationResult(summary, csvRows, targetData.TargetHeaders);
    }

    private static TargetData ReadTargetData(IFormFile targetFile)
    {
        var records = new List<TargetRecord>();
        var targetRows = 0;
        var validVendorRows = 0;
        var validCompositeRows = 0;

        using var stream = targetFile.OpenReadStream();
        using var reader = ExcelReaderFactory.CreateReader(stream);

        if (!reader.Read())
        {
            throw new InvalidOperationException("Target file is empty or missing a header row.");
        }

        var headers = ReadHeaderNames(reader);
        var targetVendorIndex = FindRequiredColumnByHeaders(headers, "Vendor Account");
        var targetInvoiceIndex = FindRequiredColumnByHeaders(headers, "Invoice No.");

        while (reader.Read())
        {
            targetRows += 1;
            var rowNumber = targetRows + 1;
            var values = ReadRowValues(reader, headers.Length);
            var vendorAccount = ExtractVendorAccount(values[targetVendorIndex]);
            var invoice = values[targetInvoiceIndex];

            records.Add(new TargetRecord(rowNumber, vendorAccount, invoice, values));

            if (!string.IsNullOrEmpty(vendorAccount))
            {
                validVendorRows += 1;
            }

            if (!string.IsNullOrEmpty(vendorAccount) && !string.IsNullOrEmpty(invoice))
            {
                validCompositeRows += 1;
            }
        }

        return new TargetData(targetRows, validVendorRows, validCompositeRows, records, headers);
    }

    private static SourceData ReadSourceData(IFormFile sourceFile)
    {
        var recordsByVendor = new Dictionary<string, List<SourceRecord>>(StringComparer.OrdinalIgnoreCase);

        using var stream = sourceFile.OpenReadStream();
        using var reader = ExcelReaderFactory.CreateReader(stream);

        if (!reader.Read())
        {
            throw new InvalidOperationException("Source file is empty or missing a header row.");
        }

        var sourceHeaders = ReadHeaderNames(reader);
        var sourceVendorIndex = FindRequiredColumnByHeaders(sourceHeaders, "Vendor:");
        var sourceDocumentIndex = FindRequiredColumnByHeaders(sourceHeaders, "Document Number:");
        var sourceBatchNoIndex = FindOptionalColumnByHeaders(sourceHeaders, "Batch No.:");
        var sourceDescriptionIndex = FindOptionalColumnByHeaders(sourceHeaders, "Description:");
        var sourceEntryNoIndex = FindOptionalColumnByHeaders(sourceHeaders, "Entry No.:");
        var sourceInvoiceDescriptionIndex = FindOptionalColumnByHeaders(sourceHeaders, "Invoice Description");
        var sourceDocumentTypeIndex = FindOptionalColumnByHeaders(sourceHeaders, "Document Type:");
        var sourcePostingDateIndex = FindOptionalColumnByHeaders(sourceHeaders, "Posting Date:");
        var sourceYearPeriodIndex = FindOptionalColumnByHeaders(sourceHeaders, "Year - Period:");
        var sourceOrderNumberIndex = FindOptionalColumnByHeaders(sourceHeaders, "Order Number:");
        var sourceTermsIndex = FindOptionalColumnByHeaders(sourceHeaders, "Terms:");
        var sourceDueDateIndex = FindOptionalColumnByHeaders(sourceHeaders, "Due Date:");
        var sourceGLAccountIndex = FindOptionalColumnByHeaders(sourceHeaders, "G/L Account");
        var sourceAccountDescriptionIndex = FindOptionalColumnByHeaders(sourceHeaders, "Account Description");
        var sourceDetailDescTaxAuthIndex = FindOptionalColumnByHeaders(sourceHeaders, "Detail Desc/ Tax Auth");
        var sourceAccountSetIndex = FindOptionalColumnByHeaders(sourceHeaders, "Account Set:");
        var sourceTaxGroupIndex = FindOptionalColumnByHeaders(sourceHeaders, "Tax Group:");
        var sourceDocumentDateIndex = FindOptionalColumnByHeaders(sourceHeaders, "Document Date:");
        var sourcePONumberIndex = FindOptionalColumnByHeaders(sourceHeaders, "PO Number:");
        var sourceNetDistAmtIndex = FindOptionalColumnByHeaders(sourceHeaders, "Net Dist. Amt.");
        var sourceDistTaxIndex = FindOptionalColumnByHeaders(sourceHeaders, "Dist. Tax.");
        var sourceInvTotalIndex = FindOptionalColumnByHeaders(sourceHeaders, "Inv Total:");
        var sourceExchangeRateIndex = FindOptionalColumnByHeaders(sourceHeaders, "Exchange Rate:");

        var sourceRows = 0;
        while (reader.Read())
        {
            sourceRows += 1;
            var rowNumber = sourceRows + 1;
            var sourceVendorRaw = Normalize(reader.GetValue(sourceVendorIndex));
            var sourceVendorAccount = ExtractVendorAccount(sourceVendorRaw);
            var sourceDocumentRaw = Normalize(reader.GetValue(sourceDocumentIndex));

            if (string.IsNullOrEmpty(sourceVendorAccount))
            {
                continue;
            }

            var sourceRecord = new SourceRecord(
                rowNumber,
                sourceVendorAccount,
                sourceDocumentRaw,
                ReadOptionalValue(reader, sourceBatchNoIndex),
                ReadOptionalValue(reader, sourceDescriptionIndex),
                ReadOptionalValue(reader, sourceEntryNoIndex),
                ReadOptionalValue(reader, sourceInvoiceDescriptionIndex),
                sourceVendorRaw,
                sourceDocumentRaw,
                ReadOptionalValue(reader, sourceDocumentTypeIndex),
                ReadOptionalValue(reader, sourcePONumberIndex),
                ReadOptionalValue(reader, sourceDocumentDateIndex),
                ReadOptionalValue(reader, sourcePostingDateIndex),
                ReadOptionalValue(reader, sourceYearPeriodIndex),
                ReadOptionalValue(reader, sourceOrderNumberIndex),
                ReadOptionalValue(reader, sourceAccountSetIndex),
                ReadOptionalValue(reader, sourceTaxGroupIndex),
                ReadOptionalValue(reader, sourceExchangeRateIndex),
                ReadOptionalValue(reader, sourceTermsIndex),
                ReadOptionalValue(reader, sourceDueDateIndex),
                ReadOptionalValue(reader, sourceGLAccountIndex),
                ReadOptionalValue(reader, sourceAccountDescriptionIndex),
                ReadOptionalValue(reader, sourceDetailDescTaxAuthIndex),
                ReadOptionalValue(reader, sourceNetDistAmtIndex),
                ReadOptionalValue(reader, sourceDistTaxIndex),
                ReadOptionalValue(reader, sourceInvTotalIndex)
            );

            if (!recordsByVendor.TryGetValue(sourceVendorAccount, out var vendorRows))
            {
                vendorRows = new List<SourceRecord>();
                recordsByVendor[sourceVendorAccount] = vendorRows;
            }
            vendorRows.Add(sourceRecord);
        }

        return new SourceData(sourceRows, recordsByVendor);
    }

    private static string BuildCsv(IEnumerable<CsvRow> rows, IReadOnlyList<string> targetHeaders)
    {
        var sb = new StringBuilder();
        var fullHeaders = targetHeaders.Concat(SourceOutputHeaders);
        sb.AppendLine(string.Join(",", fullHeaders.Select(EscapeCsv)));

        foreach (var row in rows)
        {
            var fields = new List<string>(targetHeaders.Count + SourceOutputHeaders.Length);
            fields.AddRange(row.TargetValues.Select(EscapeCsv));

            if (row.Source is null)
            {
                fields.AddRange(Enumerable.Repeat(string.Empty, SourceOutputHeaders.Length));
            }
            else
            {
                fields.Add(row.Source.RowNumber.ToString());
                fields.Add(EscapeCsv(row.Source.BatchNo));
                fields.Add(EscapeCsv(row.Source.Description));
                fields.Add(EscapeCsv(row.Source.EntryNo));
                fields.Add(EscapeCsv(row.Source.InvoiceDescription));
                fields.Add(EscapeCsv(row.Source.VendorRaw));
                fields.Add(EscapeCsv(row.Source.DocumentNumberRaw));
                fields.Add(EscapeCsv(row.Source.DocumentType));
                fields.Add(EscapeCsv(row.Source.PONumber));
                fields.Add(EscapeCsv(row.Source.DocumentDate));
                fields.Add(EscapeCsv(row.Source.PostingDate));
                fields.Add(EscapeCsv(row.Source.YearPeriod));
                fields.Add(EscapeCsv(row.Source.OrderNumber));
                fields.Add(EscapeCsv(row.Source.AccountSet));
                fields.Add(EscapeCsv(row.Source.TaxGroup));
                fields.Add(EscapeCsv(row.Source.ExchangeRate));
                fields.Add(EscapeCsv(row.Source.Terms));
                fields.Add(EscapeCsv(row.Source.DueDate));
                fields.Add(EscapeCsv(row.Source.GLAccount));
                fields.Add(EscapeCsv(row.Source.AccountDescription));
                fields.Add(EscapeCsv(row.Source.DetailDescTaxAuth));
                fields.Add(EscapeCsv(row.Source.NetDistAmt));
                fields.Add(EscapeCsv(row.Source.DistTax));
                fields.Add(EscapeCsv(row.Source.InvTotal));
            }

            sb.AppendLine(string.Join(",", fields));
        }

        return sb.ToString();
    }

    private static string[] ReadHeaderNames(IExcelDataReader reader)
    {
        var headers = new string[reader.FieldCount];
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < reader.FieldCount; i += 1)
        {
            var raw = reader.GetValue(i)?.ToString()?.Trim();
            var baseHeader = string.IsNullOrEmpty(raw) ? $"Column{i + 1}" : raw;

            if (!seen.TryAdd(baseHeader, 1))
            {
                seen[baseHeader] += 1;
                headers[i] = $"{baseHeader}_{seen[baseHeader]}";
            }
            else
            {
                headers[i] = baseHeader;
            }
        }

        return headers;
    }

    private static string[] ReadRowValues(IExcelDataReader reader, int fieldCount)
    {
        var values = new string[fieldCount];
        for (var i = 0; i < fieldCount; i += 1)
        {
            values[i] = Normalize(reader.GetValue(i));
        }

        return values;
    }

    private static int FindRequiredColumnByHeaders(IReadOnlyList<string> headers, string requiredHeader)
    {
        var index = FindOptionalColumnByHeaders(headers, requiredHeader);
        if (!index.HasValue)
        {
            throw new InvalidOperationException($"Required column '{requiredHeader}' was not found in the header row.");
        }

        return index.Value;
    }

    private static int? FindOptionalColumnByHeaders(IReadOnlyList<string> headers, string header)
    {
        var required = CanonicalizeHeader(header);
        for (var i = 0; i < headers.Count; i += 1)
        {
            if (CanonicalizeHeader(headers[i]) == required)
            {
                return i;
            }
        }

        return null;
    }

    private static string CanonicalizeHeader(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        Span<char> buffer = stackalloc char[value.Length];
        var length = 0;
        foreach (var ch in value)
        {
            if (!char.IsLetterOrDigit(ch))
            {
                continue;
            }

            buffer[length] = char.ToLowerInvariant(ch);
            length += 1;
        }

        return new string(buffer[..length]);
    }

    private static string ExtractVendorAccount(string value)
    {
        var normalized = Normalize(value);
        if (string.IsNullOrEmpty(normalized))
        {
            return string.Empty;
        }

        for (var i = 0; i < normalized.Length; i += 1)
        {
            if (char.IsWhiteSpace(normalized[i]) || IsSeparator(normalized[i]))
            {
                return normalized[..i];
            }
        }

        return normalized;
    }

    private static bool IsSeparator(char c)
    {
        return c is '|' or ',' or ';' or ':' or '(' or ')' or '_' or '/' or '\\';
    }

    private static string Normalize(object? value)
    {
        return value?.ToString()?.Trim() ?? string.Empty;
    }

    private static bool IsInvoiceMatch(string sourceDocument, string targetInvoice)
    {
        if (string.IsNullOrEmpty(sourceDocument) || string.IsNullOrEmpty(targetInvoice))
        {
            return false;
        }

        if (string.Equals(sourceDocument, targetInvoice, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        const string suffix = "*VA";
        if (sourceDocument.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
        {
            var withoutSuffix = sourceDocument[..^suffix.Length].TrimEnd();
            return string.Equals(withoutSuffix, targetInvoice, StringComparison.OrdinalIgnoreCase);
        }

        return false;
    }

    private static string ReadOptionalValue(IExcelDataReader reader, int? index)
    {
        return index.HasValue ? (reader.GetValue(index.Value)?.ToString()?.Trim() ?? string.Empty) : string.Empty;
    }

    private static string EscapeCsv(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        if (!value.Contains(',') && !value.Contains('"') && !value.Contains('\n') && !value.Contains('\r'))
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"")}\"";
    }

    private sealed record TargetData(
        int TargetRows,
        int ValidTargetVendorRows,
        int ValidTargetCompositeRows,
        List<TargetRecord> Records,
        IReadOnlyList<string> TargetHeaders
    );

    private sealed record SourceData(
        int SourceRows,
        Dictionary<string, List<SourceRecord>> RecordsByVendor
    );

    private sealed record SourceRecord(
        int RowNumber,
        string VendorAccount,
        string DocumentNumber,
        string BatchNo,
        string Description,
        string EntryNo,
        string InvoiceDescription,
        string VendorRaw,
        string DocumentNumberRaw,
        string DocumentType,
        string PONumber,
        string DocumentDate,
        string PostingDate,
        string YearPeriod,
        string OrderNumber,
        string AccountSet,
        string TaxGroup,
        string ExchangeRate,
        string Terms,
        string DueDate,
        string GLAccount,
        string AccountDescription,
        string DetailDescTaxAuth,
        string NetDistAmt,
        string DistTax,
        string InvTotal
    );

    private sealed record TargetRecord(
        int RowNumber,
        string Vendor,
        string Invoice,
        string[] TargetValues
    );

    private sealed record CsvRow(
        int TargetRowNumber,
        string[] TargetValues,
        SourceRecord? Source
    );

    private sealed record MatchComputationResult(
        MatchResponse Summary,
        List<CsvRow> CsvRows,
        IReadOnlyList<string> TargetHeaders
    );
}
