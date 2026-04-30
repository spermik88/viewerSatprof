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

public interface ITransportStreamParser
{
    Task<TsParseResult> ParseFileAsync(string path, int maxPackets = 120000, CancellationToken cancellationToken = default);
}
