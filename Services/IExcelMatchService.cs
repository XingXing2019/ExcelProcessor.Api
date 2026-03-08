using ExcelProcessor.Api.Models;

namespace ExcelProcessor.Api.Services;

public interface IExcelMatchService
{
    MatchResponse Match(IFormFile sourceFile, IFormFile targetFile);

    MatchCsvReport BuildCsvReport(IFormFile sourceFile, IFormFile targetFile);
    
    MatchExcelReport BuildExcelReport(IFormFile sourceFile, IFormFile targetFile);

    IReadOnlyList<string> GetNonEmptyTargetColumns(IFormFile targetFile);
}
