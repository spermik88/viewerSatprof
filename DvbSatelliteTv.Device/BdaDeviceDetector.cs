using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

public sealed class BdaDeviceDetector : IBdaDeviceDetector
{
    private static readonly Guid BdaNetworkTunerCategory = new("71985F48-1CA1-11d3-9CC8-00C04F7971E0");
    private static readonly Guid BdaReceiverComponentCategory = new("FD0A5AF4-B41D-11d2-9C95-00C04F7971E0");
    private static readonly Guid CaptureCategory = new("65E8773D-8F56-11D0-A3B9-00A0C9223196");
    private static readonly Guid SystemDeviceEnumClsid = new("62BE5D10-60EB-11D0-BD3B-00A0C911CE86");

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

    [SupportedOSPlatform("windows")]
    private static IEnumerable<BdaFilterInfo> EnumerateCategory(string categoryName, Guid category)
    {
        ICreateDevEnum? deviceEnum = null;
        IEnumMoniker? enumMoniker = null;

        try
        {
            var type = Type.GetTypeFromCLSID(SystemDeviceEnumClsid, throwOnError: true)!;
            deviceEnum = (ICreateDevEnum)Activator.CreateInstance(type)!;
            var hr = deviceEnum.CreateClassEnumerator(ref category, out enumMoniker, 0);
            if (hr != 0 || enumMoniker is null)
            {
                yield break;
            }

            var monikers = new IMoniker[1];
            while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
            {
                var moniker = monikers[0];
                IPropertyBag? propertyBag = null;

                try
                {
                    var propertyBagId = typeof(IPropertyBag).GUID;
                    moniker.BindToStorage(IntPtr.Zero, IntPtr.Zero, ref propertyBagId, out var bagObject);
                    propertyBag = (IPropertyBag)bagObject;

                    var friendlyName = ReadProperty(propertyBag, "FriendlyName");
                    var devicePath = ReadProperty(propertyBag, "DevicePath");
                    if (!string.IsNullOrWhiteSpace(friendlyName) || !string.IsNullOrWhiteSpace(devicePath))
                    {
                        yield return new BdaFilterInfo(
                            categoryName,
                            friendlyName.Length > 0 ? friendlyName : "(unnamed filter)",
                            devicePath);
                    }
                }
                finally
                {
                    if (propertyBag is not null)
                    {
                        Marshal.ReleaseComObject(propertyBag);
                    }

                    Marshal.ReleaseComObject(moniker);
                }
            }
        }
        finally
        {
            if (enumMoniker is not null)
            {
                Marshal.ReleaseComObject(enumMoniker);
            }

            if (deviceEnum is not null)
            {
                Marshal.ReleaseComObject(deviceEnum);
            }
        }
    }

    private static string ReadProperty(IPropertyBag propertyBag, string name)
    {
        object value = string.Empty;
        return propertyBag.Read(name, ref value, IntPtr.Zero) == 0 ? value?.ToString() ?? string.Empty : string.Empty;
    }

    private static bool LooksLikeDvbFilter(string value)
    {
        return value.Contains("BDA", StringComparison.OrdinalIgnoreCase)
            || value.Contains("DVB", StringComparison.OrdinalIgnoreCase)
            || value.Contains("Prof", StringComparison.OrdinalIgnoreCase)
            || value.Contains("Tuner", StringComparison.OrdinalIgnoreCase)
            || value.Contains("TS", StringComparison.OrdinalIgnoreCase);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("29840822-5B84-11D0-BD3B-00A0C911CE86")]
    private interface ICreateDevEnum
    {
        [PreserveSig]
        int CreateClassEnumerator([In] ref Guid classType, out IEnumMoniker? enumMoniker, int flags);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("00000102-0000-0000-C000-000000000046")]
    private interface IEnumMoniker
    {
        [PreserveSig]
        int Next(int celt, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IMoniker[] rgelt, IntPtr pceltFetched);

        [PreserveSig]
        int Skip(int celt);

        [PreserveSig]
        int Reset();

        [PreserveSig]
        int Clone(out IEnumMoniker enumMoniker);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("0000000F-0000-0000-C000-000000000046")]
    private interface IMoniker
    {
        void GetClassID(out Guid pClassID);
        void IsDirty();
        void Load(IntPtr stream);
        void Save(IntPtr stream, bool clearDirty);
        void GetSizeMax(out long size);
        void BindToObject(IntPtr bindContext, IntPtr monikerToLeft, ref Guid riidResult, out object result);
        void BindToStorage(IntPtr bindContext, IntPtr monikerToLeft, ref Guid riid, out object result);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("55272A00-42CB-11CE-8135-00AA004BB851")]
    private interface IPropertyBag
    {
        [PreserveSig]
        int Read([MarshalAs(UnmanagedType.LPWStr)] string propertyName, ref object value, IntPtr errorLog);

        [PreserveSig]
        int Write([MarshalAs(UnmanagedType.LPWStr)] string propertyName, ref object value);
    }
}
