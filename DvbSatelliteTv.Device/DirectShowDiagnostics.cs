using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using DirectShowLib;

namespace DvbSatelliteTv.Device;

internal static class DirectShowDiagnostics
{
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
    public static void ReleaseCom(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }
}
