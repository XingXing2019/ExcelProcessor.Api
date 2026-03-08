using Microsoft.AspNetCore.Mvc;

namespace ExcelProcessor.Api.Controllers;

[ApiController]
[Route("")]
public sealed class StatusController : ControllerBase
{
    [HttpGet]
    public IActionResult GetStatus()
    {
        return Ok(new { status = "ok", service = "ExcelProcessor.Api" });
    }
}
