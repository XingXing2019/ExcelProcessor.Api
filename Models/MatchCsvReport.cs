namespace ExcelProcessor.Api.Models;

public sealed record MatchCsvReport(
    byte[] Content,
    string FileName
);
