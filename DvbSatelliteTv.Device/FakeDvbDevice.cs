using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

public sealed class FakeDvbDevice : IDvbDevice
{
    public Task<DeviceInfo> GetDeviceInfoAsync(CancellationToken cancellationToken = default)
    {
        var info = new DeviceInfo(
            "Simulator: Prof Revolution BDA device",
            "FakeDvbDevice",
            IsPresent: false,
            "Real card is not connected. Simulator is used for UI and scan flow development.");

        return Task.FromResult(info);
    }

    public async IAsyncEnumerable<ScanProgress> ScanAsync(
        IEnumerable<Transponder> transponders,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        foreach (var transponder in transponders)
        {
            cancellationToken.ThrowIfCancellationRequested();

            yield return new ScanProgress(
                transponder,
                ScanStatus.Running,
                new SignalInfo(false, 0, 0, "Tuning transponder"),
                []);

            await Task.Delay(350, cancellationToken);

            var channels = CreateChannels(transponder);
            var hasLock = channels.Count > 0;

            yield return new ScanProgress(
                transponder,
                hasLock ? ScanStatus.Locked : ScanStatus.NoSignal,
                hasLock
                    ? new SignalInfo(true, 72, 64, "Simulated lock")
                    : new SignalInfo(false, 18, 0, "Simulated no signal"),
                channels);
        }

        yield return new ScanProgress(
            HotbirdDefaults.FallbackTransponders[0],
            ScanStatus.Completed,
            new SignalInfo(false, 0, 0, "Scan completed"),
            []);
    }

    private static IReadOnlyList<Channel> CreateChannels(Transponder transponder)
    {
        if (transponder.FrequencyMhz % 3 == 0)
        {
            return [];
        }

        return
        [
            new($"FTA Demo {transponder.FrequencyMhz} News", transponder.FrequencyMhz, transponder.SymbolRateKsps, transponder.Polarization, transponder.FrequencyMhz, 101, 102, true),
            new($"FTA Demo {transponder.FrequencyMhz} Music", transponder.FrequencyMhz, transponder.SymbolRateKsps, transponder.Polarization, transponder.FrequencyMhz + 1, 201, 202, true)
        ];
    }
}
