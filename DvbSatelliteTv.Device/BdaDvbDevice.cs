using System.Runtime.CompilerServices;
using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

public sealed class BdaDvbDevice : IDvbDevice
{
    private readonly IBdaDeviceDetector _detector;
    private readonly ITransportStreamRecorder _recorder;
    private readonly ITransportStreamParser _parser;
    private readonly string _captureDirectory;
    private readonly int _captureSeconds;

    public BdaDvbDevice(
        IBdaDeviceDetector detector,
        ITransportStreamRecorder recorder,
        ITransportStreamParser parser,
        string captureDirectory,
        int captureSeconds = 8)
    {
        _detector = detector;
        _recorder = recorder;
        _parser = parser;
        _captureDirectory = captureDirectory;
        _captureSeconds = captureSeconds;
    }

    public async Task<DeviceInfo> GetDeviceInfoAsync(CancellationToken cancellationToken = default)
    {
        var filters = await _detector.DetectAsync(cancellationToken);
        var profFilters = filters
            .Where(x => x.FriendlyName.Contains("Prof", StringComparison.OrdinalIgnoreCase)
                || x.DevicePath.Contains("ven_14f1", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (profFilters.Count == 0)
        {
            return new DeviceInfo(
                "No Prof BDA device detected",
                "BDA/DirectShow",
                IsPresent: false,
                "No Prof BDA tuner/capture filters were found.");
        }

        var tuner = profFilters.FirstOrDefault(x => x.Category.Contains("Tuner", StringComparison.OrdinalIgnoreCase))
            ?? profFilters[0];
        var hasTransport = profFilters.Any(x => x.Category.Contains("Receiver", StringComparison.OrdinalIgnoreCase)
            || x.FriendlyName.Contains("TS", StringComparison.OrdinalIgnoreCase)
            || x.FriendlyName.Contains("Capture", StringComparison.OrdinalIgnoreCase));

        return new DeviceInfo(
            tuner.FriendlyName,
            "Prof BDA/DirectShow",
            IsPresent: hasTransport,
            hasTransport
                ? "Prof BDA tuner and TS capture filters are available for scanning."
                : "Prof tuner exists, but TS capture filter was not found.");
    }

    public async IAsyncEnumerable<ScanProgress> ScanAsync(
        IEnumerable<Transponder> transponders,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(_captureDirectory);

        foreach (var transponder in transponders)
        {
            cancellationToken.ThrowIfCancellationRequested();

            yield return new ScanProgress(
                transponder,
                ScanStatus.Running,
                new SignalInfo(false, 0, 0, "Recording TS sample"),
                []);

            var outputPath = Path.Combine(
                _captureDirectory,
                $"scan-{DateTime.Now:yyyyMMdd-HHmmss}-{transponder.FrequencyMhz}-{transponder.Polarization}.ts");

            var capture = await _recorder.RecordAsync(new TsCaptureRequest(
                new TuneRequest(
                    transponder.FrequencyMhz,
                    transponder.SymbolRateKsps,
                    transponder.Polarization,
                    LnbLowMhz: 9750,
                    LnbHighMhz: 10600,
                    SwitchMhz: 11700),
                outputPath,
                _captureSeconds), cancellationToken);

            if (!capture.Success)
            {
                yield return new ScanProgress(
                    transponder,
                    ScanStatus.NoSignal,
                    new SignalInfo(false, 0, 0, $"No TS data captured. Bytes: {capture.BytesWritten}."),
                    [],
                    capture.Diagnostics);
                continue;
            }

            var parsed = await _parser.ParseFileAsync(capture.OutputPath, cancellationToken: cancellationToken);
            var diagnostics = capture.Diagnostics.Concat(parsed.Diagnostics).ToList();
            var channels = parsed.Services
                .Select(service => new Channel(
                    service.Name,
                    transponder.FrequencyMhz,
                    transponder.SymbolRateKsps,
                    transponder.Polarization,
                    service.ServiceId,
                    service.VideoPid ?? 0,
                    service.AudioPids.FirstOrDefault(),
                    !service.IsScrambled))
                .ToList();

            yield return new ScanProgress(
                transponder,
                channels.Count > 0 ? ScanStatus.Locked : ScanStatus.NoSignal,
                new SignalInfo(
                    channels.Count > 0,
                    0,
                    0,
                    channels.Count > 0
                        ? $"Captured {capture.BytesWritten} bytes; found {channels.Count} service(s)."
                        : $"Captured {capture.BytesWritten} bytes; no services parsed."),
                channels,
                diagnostics);
        }

        yield return new ScanProgress(
            HotbirdDefaults.FallbackTransponders[0],
            ScanStatus.Completed,
            new SignalInfo(false, 0, 0, "Scan completed"),
            []);
    }
}
