namespace ExcelProcessor.Api.Models;

public sealed record MatchExcelReport(
    byte[] Content,
    string FileName
);
