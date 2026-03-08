using ExcelProcessor.Api.Models;
using ExcelProcessor.Api.Services;
using Microsoft.AspNetCore.Mvc;

namespace ExcelProcessor.Api.Controllers;

[ApiController]
[Route("api/match")]
public sealed class MatchController : ControllerBase
{
    private readonly IExcelMatchService _excelMatchService;

    public MatchController(IExcelMatchService excelMatchService)
    {
        _excelMatchService = excelMatchService;
    }

    [HttpPost]
    [ProducesResponseType(typeof(MatchResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(ErrorResponse), StatusCodes.Status400BadRequest)]
    public IActionResult Match([FromForm] MatchRequest request)
    {
        if (request.SourceFile is null || request.TargetFile is null)
        {
            return BadRequest(new ErrorResponse("Both sourceFile and targetFile are required."));
        }

        try
        {
            var response = _excelMatchService.Match(request.SourceFile, request.TargetFile);
            return Ok(response);
        }
        catch (Exception ex)
        {
            return BadRequest(new ErrorResponse(ex.Message));
        }
    }

    [HttpPost("csv")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(ErrorResponse), StatusCodes.Status400BadRequest)]
    public IActionResult MatchCsv([FromForm] MatchRequest request)
    {
        if (request.SourceFile is null || request.TargetFile is null)
        {
            return BadRequest(new ErrorResponse("Both sourceFile and targetFile are required."));
        }

        try
        {
            var report = _excelMatchService.BuildCsvReport(request.SourceFile, request.TargetFile);
            return File(report.Content, "text/csv", report.FileName);
        }
        catch (Exception ex)
        {
            return BadRequest(new ErrorResponse(ex.Message));
        }
    }

    [HttpPost("target-columns")]
    [ProducesResponseType(typeof(TargetColumnsResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(ErrorResponse), StatusCodes.Status400BadRequest)]
    public IActionResult TargetColumns([FromForm] IFormFile? targetFile)
    {
        if (targetFile is null)
        {
            return BadRequest(new ErrorResponse("targetFile is required."));
        }

        try
        {
            var columns = _excelMatchService.GetNonEmptyTargetColumns(targetFile);
            return Ok(new TargetColumnsResponse(columns));
        }
        catch (Exception ex)
        {
            return BadRequest(new ErrorResponse(ex.Message));
        }
    }
}
