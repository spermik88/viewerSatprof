namespace DvbSatelliteTv.Core;

public enum Polarization
{
    Horizontal,
    Vertical
}

public enum ScanStatus
{
    NotStarted,
    Running,
    Locked,
    NoSignal,
    Completed,
    Failed
}

public sealed record Satellite(string Name, double OrbitalPositionEast);

public sealed record Transponder(
    int FrequencyMhz,
    int SymbolRateKsps,
    Polarization Polarization,
    string Modulation,
    string Fec,
    string Source);

public sealed record Channel(
    string Name,
    int FrequencyMhz,
    int SymbolRateKsps,
    Polarization Polarization,
    int ServiceId,
    int VideoPid,
    int AudioPid,
    bool IsFreeToAir);

public sealed record DeviceInfo(string Name, string Driver, bool IsPresent, string Notes);

public sealed record BdaFilterInfo(string Category, string FriendlyName, string DevicePath);

public sealed record TuneRequest(
    int FrequencyMhz,
    int SymbolRateKsps,
    Polarization Polarization,
    int LnbLowMhz,
    int LnbHighMhz,
    int SwitchMhz);

public sealed record TuneResult(
    bool CanTune,
    string Stage,
    int IntermediateFrequencyMhz,
    bool Use22KhzTone,
    int LnbVoltage,
    SignalInfo Signal,
    IReadOnlyList<string> Diagnostics);

public sealed record BdaGraphProbeResult(
    bool GraphCreated,
    bool NetworkProviderAdded,
    bool TunerAdded,
    bool TransportAdded,
    bool TunerConnected,
    bool TransportConnected,
    bool TuneRequestSubmitted,
    bool GraphRan,
    bool? SignalLocked,
    int? SignalStrength,
    int? SignalQuality,
    IReadOnlyList<string> Diagnostics);

public sealed record SignalInfo(bool HasLock, int StrengthPercent, int QualityPercent, string Message);

public sealed record ScanProgress(
    Transponder Transponder,
    ScanStatus Status,
    SignalInfo Signal,
    IReadOnlyList<Channel> Channels,
    IReadOnlyList<string>? Diagnostics = null);

public interface IDvbDevice
{
    Task<DeviceInfo> GetDeviceInfoAsync(CancellationToken cancellationToken = default);

    IAsyncEnumerable<ScanProgress> ScanAsync(
        IEnumerable<Transponder> transponders,
        CancellationToken cancellationToken = default);
}

public interface IBdaDeviceDetector
{
    Task<IReadOnlyList<BdaFilterInfo>> DetectAsync(CancellationToken cancellationToken = default);
}

public interface ITuneMonitor
{
    Task<TuneResult> TuneAsync(TuneRequest request, CancellationToken cancellationToken = default);
}

public interface IBdaGraphBuilder
{
    Task<BdaGraphProbeResult> ProbeAsync(TuneRequest request, CancellationToken cancellationToken = default);
}

public static class HotbirdDefaults
{
    public static Satellite Satellite { get; } = new("Hotbird 13E", 13.0);

    public static string TransponderFileName { get; } = "hotbird-13e-transponders.json";

    public static IReadOnlyList<Transponder> FallbackTransponders { get; } =
    [
        new(10719, 27500, Polarization.Vertical, "DVB-S QPSK", "5/6", "Seed list"),
        new(10815, 27500, Polarization.Horizontal, "DVB-S QPSK", "5/6", "Seed list"),
        new(10930, 30000, Polarization.Horizontal, "DVB-S2 8PSK", "2/3", "Seed list"),
        new(11034, 27500, Polarization.Vertical, "DVB-S QPSK", "3/4", "Seed list"),
        new(11179, 27500, Polarization.Horizontal, "DVB-S2 8PSK", "3/4", "Seed list"),
        new(11373, 27500, Polarization.Horizontal, "DVB-S QPSK", "3/4", "Seed list"),
        new(11566, 29900, Polarization.Horizontal, "DVB-S2 8PSK", "3/4", "Seed list"),
        new(11727, 29900, Polarization.Vertical, "DVB-S2 8PSK", "3/4", "Seed list"),
        new(11881, 27500, Polarization.Vertical, "DVB-S QPSK", "3/4", "Seed list"),
        new(12149, 27500, Polarization.Vertical, "DVB-S QPSK", "3/4", "Seed list"),
        new(12476, 29900, Polarization.Horizontal, "DVB-S2 8PSK", "3/4", "Seed list")
    ];
}
