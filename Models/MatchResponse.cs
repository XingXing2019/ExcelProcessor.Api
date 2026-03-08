namespace ExcelProcessor.Api.Models;

public sealed record MatchResponse(
    int SourceRows,
    int TargetRows,
    int TargetRowsWithVendor,
    int TargetRowsWithVendorAndInvoice,
    int TargetRowsVendorMatchedInSource,
    int TargetRowsFullyMatchedInSource,
    int TargetRowsVendorMatchedButInvoiceNotMatchedInSource,
    int ElapsedMs
);
