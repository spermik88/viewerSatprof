using System.Runtime.InteropServices;
using System.Reflection;
using System.Runtime.Versioning;
using DirectShowLib;

namespace DvbSatelliteTv.Device;

internal static class DirectShowDiagnostics
{
    private static readonly Lazy<IReadOnlyDictionary<Guid, string>> GuidNames = new(BuildGuidNames);

    [SupportedOSPlatform("windows")]
    public static void DumpPins(IBaseFilter filter, string label, ICollection<string> diagnostics)
    {
        IEnumPins? enumPins = null;

        try
        {
            var hr = filter.EnumPins(out enumPins);
            if (hr != 0 || enumPins is null)
            {
                diagnostics.Add($"{label}: EnumPins returned 0x{hr:X8}.");
                return;
            }

            var pins = new IPin[1];
            var index = 0;
            while (enumPins.Next(1, pins, IntPtr.Zero) == 0)
            {
                var pin = pins[0];
                try
                {
                    pin.QueryPinInfo(out var info);
                    pin.QueryDirection(out var direction);
                    var connected = pin.ConnectedTo(out var connectedPin) == 0;
                    ReleaseCom(connectedPin);
                    diagnostics.Add($"{label} pin[{index}]: {direction} '{info.name}', connected={connected}");
                    DumpPinMediaTypes(pin, $"{label} pin[{index}]", diagnostics);
                    ReleaseCom(info.filter);
                }
                catch (Exception ex)
                {
                    diagnostics.Add($"{label} pin[{index}]: failed to read pin info: {ex.Message}");
                }
                finally
                {
                    ReleaseCom(pin);
                }

                index++;
            }

            if (index == 0)
            {
                diagnostics.Add($"{label}: no pins.");
            }
        }
        catch (Exception ex)
        {
            diagnostics.Add($"{label}: pin dump failed: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            ReleaseCom(enumPins);
        }
    }

    [SupportedOSPlatform("windows")]
    public static void DumpPinMediaTypes(IPin pin, string label, ICollection<string> diagnostics)
    {
        IEnumMediaTypes? enumMediaTypes = null;

        try
        {
            var hr = pin.EnumMediaTypes(out enumMediaTypes);
            if (hr != 0 || enumMediaTypes is null)
            {
                diagnostics.Add($"{label}: EnumMediaTypes returned 0x{hr:X8}.");
                return;
            }

            var mediaTypes = new AMMediaType[1];
            var index = 0;
            while (enumMediaTypes.Next(1, mediaTypes, IntPtr.Zero) == 0)
            {
                var mediaType = mediaTypes[0];
                try
                {
                    diagnostics.Add(
                        $"{label} media[{index}]: major={NameGuid(mediaType.majorType)}, subtype={NameGuid(mediaType.subType)}, format={NameGuid(mediaType.formatType)}, sampleSize={mediaType.sampleSize}, fixed={mediaType.fixedSizeSamples}");
                }
                finally
                {
                    DsUtils.FreeAMMediaType(mediaType);
                }

                index++;
            }

            if (index == 0)
            {
                diagnostics.Add($"{label}: no enumerated media types.");
            }
        }
        catch (Exception ex)
        {
            diagnostics.Add($"{label}: media type dump failed: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            ReleaseCom(enumMediaTypes);
        }
    }

    [SupportedOSPlatform("windows")]
    public static void ReleaseCom(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }

    private static string NameGuid(Guid guid)
    {
        return GuidNames.Value.TryGetValue(guid, out var name)
            ? name
            : guid.ToString("D");
    }

    private static IReadOnlyDictionary<Guid, string> BuildGuidNames()
    {
        var result = new Dictionary<Guid, string>();
        AddGuidNames(typeof(MediaType), "MediaType", result);
        AddGuidNames(typeof(MediaSubType), "MediaSubType", result);
        AddGuidNames(typeof(FormatType), "FormatType", result);
        return result;
    }

    private static void AddGuidNames(Type type, string prefix, Dictionary<Guid, string> result)
    {
        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Static))
        {
            if (field.FieldType != typeof(Guid) || field.GetValue(null) is not Guid guid)
            {
                continue;
            }

            result.TryAdd(guid, $"{prefix}.{field.Name}");
        }
    }
}
