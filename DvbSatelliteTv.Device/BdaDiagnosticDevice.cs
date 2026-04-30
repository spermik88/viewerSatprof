using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

public sealed class BdaDiagnosticDevice : IDvbDevice
{
    private readonly IBdaDeviceDetector _detector;
    private readonly FakeDvbDevice _scanSimulator = new();

    public BdaDiagnosticDevice(IBdaDeviceDetector detector)
    {
        _detector = detector;
    }

    public async Task<DeviceInfo> GetDeviceInfoAsync(CancellationToken cancellationToken = default)
    {
        var filters = await _detector.DetectAsync(cancellationToken);
        var profFilters = filters
            .Where(x => x.FriendlyName.Contains("Prof", StringComparison.OrdinalIgnoreCase)
                || x.DevicePath.Contains("ven_14f1", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (profFilters.Count > 0)
        {
            var tuner = profFilters.FirstOrDefault(x => x.Category.Contains("Tuner", StringComparison.OrdinalIgnoreCase))
                ?? profFilters[0];
            var hasTransport = profFilters.Any(x => x.Category.Contains("Receiver", StringComparison.OrdinalIgnoreCase)
                || x.FriendlyName.Contains("TS", StringComparison.OrdinalIgnoreCase)
                || x.FriendlyName.Contains("Capture", StringComparison.OrdinalIgnoreCase));

            return new DeviceInfo(
                tuner.FriendlyName,
                "Prof BDA/DirectShow",
                IsPresent: true,
                hasTransport
                    ? "Prof BDA tuner and transport/capture filter were detected. Tuning graph is not implemented yet."
                    : "Prof BDA tuner was detected, but transport/capture filter was not listed by DirectShow.");
        }

        if (filters.Count > 0)
        {
            return new DeviceInfo(
                filters[0].FriendlyName,
                "BDA/DirectShow",
                IsPresent: true,
                "BDA filters were detected, but no Prof-specific filter was identified.");
        }

        return new DeviceInfo(
            "No BDA DVB device detected",
            "BDA/DirectShow",
            IsPresent: false,
            "No BDA Network Tuner or DVB-like capture filters were found.");
    }

    public IAsyncEnumerable<ScanProgress> ScanAsync(
        IEnumerable<Transponder> transponders,
        CancellationToken cancellationToken = default)
    {
        return _scanSimulator.ScanAsync(transponders, cancellationToken);
    }
}
