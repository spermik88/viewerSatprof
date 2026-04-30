using System.Runtime.Versioning;
using DirectShowLib;
using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

public sealed class BdaDeviceDetector : IBdaDeviceDetector
{
    private static readonly Guid BdaNetworkTunerCategory = new("71985F48-1CA1-11d3-9CC8-00C04F7971E0");
    private static readonly Guid BdaReceiverComponentCategory = new("FD0A5AF4-B41D-11d2-9C95-00C04F7971E0");
    private static readonly Guid CaptureCategory = new("65E8773D-8F56-11D0-A3B9-00A0C9223196");

    [SupportedOSPlatform("windows")]
    public Task<IReadOnlyList<BdaFilterInfo>> DetectAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var filters = new List<BdaFilterInfo>();
        filters.AddRange(EnumerateCategory("BDA Network Tuner", BdaNetworkTunerCategory));
        filters.AddRange(EnumerateCategory("BDA Receiver/TS", BdaReceiverComponentCategory));
        filters.AddRange(EnumerateCategory("Capture", CaptureCategory)
            .Where(x => LooksLikeDvbFilter(x.FriendlyName) || LooksLikeDvbFilter(x.DevicePath)));

        return Task.FromResult<IReadOnlyList<BdaFilterInfo>>(
            filters
                .GroupBy(x => $"{x.Category}|{x.FriendlyName}|{x.DevicePath}", StringComparer.OrdinalIgnoreCase)
                .Select(x => x.First())
                .OrderBy(x => x.Category)
                .ThenBy(x => x.FriendlyName)
                .ToList());
    }

    private static IEnumerable<BdaFilterInfo> EnumerateCategory(string categoryName, Guid category)
    {
        DsDevice[] devices;

        try
        {
            devices = DsDevice.GetDevicesOfCat(category);
        }
        catch
        {
            yield break;
        }

        foreach (var device in devices)
        {
            yield return new BdaFilterInfo(
                categoryName,
                string.IsNullOrWhiteSpace(device.Name) ? "(unnamed filter)" : device.Name,
                device.DevicePath ?? string.Empty);
        }
    }

    private static bool LooksLikeDvbFilter(string value)
    {
        return value.Contains("BDA", StringComparison.OrdinalIgnoreCase)
            || value.Contains("DVB", StringComparison.OrdinalIgnoreCase)
            || value.Contains("Prof", StringComparison.OrdinalIgnoreCase)
            || value.Contains("Tuner", StringComparison.OrdinalIgnoreCase)
            || value.Contains("TS", StringComparison.OrdinalIgnoreCase);
    }
}
