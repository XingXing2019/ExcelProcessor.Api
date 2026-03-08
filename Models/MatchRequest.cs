namespace ExcelProcessor.Api.Models;

public sealed class MatchRequest
{
    public IFormFile? SourceFile { get; init; }

    public IFormFile? TargetFile { get; init; }
}
