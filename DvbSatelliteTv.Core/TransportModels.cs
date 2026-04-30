namespace DvbSatelliteTv.Core;

public sealed record TsService(
    int ServiceId,
    string Name,
    string Provider,
    int PmtPid,
    int? PcrPid,
    int? VideoPid,
    IReadOnlyList<int> AudioPids,
    bool IsScrambled);

public sealed record TsParseResult(
    string Source,
    int PacketsRead,
    IReadOnlyList<TsService> Services,
    IReadOnlyList<string> Diagnostics);

public sealed record TsCaptureRequest(
    TuneRequest TuneRequest,
    string OutputPath,
    int DurationSeconds);

public sealed record TsCaptureResult(
    bool Success,
    string OutputPath,
    long BytesWritten,
    IReadOnlyList<string> Diagnostics);

public interface ITransportStreamParser
{
    Task<TsParseResult> ParseFileAsync(string path, int maxPackets = 120000, CancellationToken cancellationToken = default);
}

public interface ITransportStreamRecorder
{
    Task<TsCaptureResult> RecordAsync(TsCaptureRequest request, CancellationToken cancellationToken = default);
}
