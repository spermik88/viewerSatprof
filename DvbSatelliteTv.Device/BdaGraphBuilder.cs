using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using DirectShowLib;
using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

public sealed class BdaGraphBuilder : IBdaGraphBuilder
{
    private static readonly Guid BdaNetworkTunerCategory = new("71985F48-1CA1-11d3-9CC8-00C04F7971E0");
    private static readonly Guid BdaReceiverComponentCategory = new("FD0A5AF4-B41D-11d2-9C95-00C04F7971E0");

    [SupportedOSPlatform("windows")]
    public Task<BdaGraphProbeResult> ProbeAsync(TuneRequest request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var diagnostics = new List<string>();
        IGraphBuilder? graph = null;
        IBaseFilter? networkProvider = null;
        IBaseFilter? tuner = null;
        IBaseFilter? transport = null;
        var graphCreated = false;
        var networkProviderAdded = false;
        var tunerAdded = false;
        var transportAdded = false;
        var tunerConnected = false;
        var transportConnected = false;

        try
        {
            graph = (IGraphBuilder)new FilterGraph();
            graphCreated = true;
            diagnostics.Add("FilterGraph created.");

            networkProvider = (IBaseFilter)new DVBSNetworkProvider();
            var hr = graph.AddFilter(networkProvider, "Microsoft DVB-S Network Provider");
            DsError.ThrowExceptionForHR(hr);
            networkProviderAdded = true;
            diagnostics.Add("DVB-S Network Provider added.");

            tuner = FindAndCreateFilter(BdaNetworkTunerCategory, "Prof", diagnostics);
            if (tuner is null)
            {
                diagnostics.Add("Prof BDA tuner filter was not found in BDA Network Tuner category.");
                return Task.FromResult(CreateResult());
            }

            hr = graph.AddFilter(tuner, "Prof BDA Tuner/Demod");
            DsError.ThrowExceptionForHR(hr);
            tunerAdded = true;
            diagnostics.Add("Prof BDA tuner filter added.");

            transport = FindAndCreateFilter(BdaReceiverComponentCategory, "Prof", diagnostics);
            if (transport is null)
            {
                diagnostics.Add("Prof TS capture/receiver filter was not found in BDA Receiver/TS category.");
                return Task.FromResult(CreateResult());
            }

            hr = graph.AddFilter(transport, "Prof TS Capture");
            DsError.ThrowExceptionForHR(hr);
            transportAdded = true;
            diagnostics.Add("Prof TS capture/receiver filter added.");

            tunerConnected = TryConnect(graph, networkProvider, tuner, diagnostics, "Network Provider -> Tuner");
            transportConnected = TryConnect(graph, tuner, transport, diagnostics, "Tuner -> TS Capture");

            diagnostics.Add("BDA graph probe completed. Tune request submission is not implemented yet.");
            return Task.FromResult(CreateResult());
        }
        catch (Exception ex)
        {
            diagnostics.Add($"Graph probe failed: {ex.GetType().Name}: {ex.Message}");
            return Task.FromResult(CreateResult());
        }
        finally
        {
            ReleaseCom(transport);
            ReleaseCom(tuner);
            ReleaseCom(networkProvider);
            ReleaseCom(graph);
        }

        BdaGraphProbeResult CreateResult()
        {
            return new BdaGraphProbeResult(
                graphCreated,
                networkProviderAdded,
                tunerAdded,
                transportAdded,
                tunerConnected,
                transportConnected,
                diagnostics);
        }
    }

    private static IBaseFilter? FindAndCreateFilter(Guid category, string namePart, List<string> diagnostics)
    {
        foreach (var device in DsDevice.GetDevicesOfCat(category))
        {
            diagnostics.Add($"Candidate filter: {device.Name}");
            if (!device.Name.Contains(namePart, StringComparison.OrdinalIgnoreCase)
                && !device.DevicePath.Contains("ven_14f1", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var bindCtx = IntPtr.Zero;
            try
            {
                var filterId = typeof(IBaseFilter).GUID;
                device.Mon.BindToObject(null!, null!, ref filterId, out var filterObject);
                diagnostics.Add($"Selected filter: {device.Name}");
                return (IBaseFilter)filterObject;
            }
            catch (Exception ex)
            {
                diagnostics.Add($"Could not bind filter {device.Name}: {ex.Message}");
            }
            finally
            {
                if (bindCtx != IntPtr.Zero)
                {
                    Marshal.Release(bindCtx);
                }
            }
        }

        return null;
    }

    [SupportedOSPlatform("windows")]
    private static bool TryConnect(IGraphBuilder graph, IBaseFilter upstream, IBaseFilter downstream, List<string> diagnostics, string label)
    {
        var outputPin = DsFindPin.ByDirection(upstream, PinDirection.Output, 0);
        var inputPin = DsFindPin.ByDirection(downstream, PinDirection.Input, 0);

        if (outputPin is null || inputPin is null)
        {
            diagnostics.Add($"{label}: pin lookup failed.");
            ReleaseCom(outputPin);
            ReleaseCom(inputPin);
            return false;
        }

        try
        {
            var hr = graph.Connect(outputPin, inputPin);
            if (hr == 0)
            {
                diagnostics.Add($"{label}: connected.");
                return true;
            }

            diagnostics.Add($"{label}: Connect returned 0x{hr:X8}.");
            return false;
        }
        finally
        {
            ReleaseCom(outputPin);
            ReleaseCom(inputPin);
        }
    }

    [SupportedOSPlatform("windows")]
    private static void ReleaseCom(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }
}
