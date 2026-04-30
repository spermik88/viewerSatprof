using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using DirectShowLib;
using DirectShowLib.BDA;
using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

public sealed class BdaGraphBuilder : IBdaGraphBuilder
{
    private static readonly Guid BdaNetworkTunerCategory = new("71985F48-1CA1-11d3-9CC8-00C04F7971E0");
    private static readonly Guid BdaReceiverComponentCategory = new("FD0A5AF4-B41D-11d2-9C95-00C04F7971E0");
    private static readonly Guid DvbsNetworkProvider = new("FA4B375A-45B4-4D45-8440-263957B11623");

    [SupportedOSPlatform("windows")]
    public Task<BdaGraphProbeResult> ProbeAsync(DvbSatelliteTv.Core.TuneRequest request, CancellationToken cancellationToken = default)
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
        var tuneRequestSubmitted = false;
        var graphRan = false;
        bool? signalLocked = null;
        int? signalStrength = null;
        int? signalQuality = null;

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
            DirectShowDiagnostics.DumpPins(networkProvider, "Network Provider", diagnostics);

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
            DirectShowDiagnostics.DumpPins(tuner, "Prof BDA Tuner/Demod", diagnostics);

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
            DirectShowDiagnostics.DumpPins(transport, "Prof TS Capture", diagnostics);

            tunerConnected = TryConnect(graph, networkProvider, tuner, diagnostics, "Network Provider -> Tuner");
            transportConnected = TryConnect(graph, tuner, transport, diagnostics, "Tuner -> TS Capture");

            tuneRequestSubmitted = TrySubmitTuneRequest(networkProvider, request, diagnostics);

            var mediaControl = (IMediaControl)graph;
            hr = mediaControl.Run();
            if (hr == 0)
            {
                graphRan = true;
                diagnostics.Add("FilterGraph running.");
                Thread.Sleep(2500);
                ReadSignalStatistics(tuner, diagnostics, out signalLocked, out signalStrength, out signalQuality);
                mediaControl.Stop();
                diagnostics.Add("FilterGraph stopped.");
            }
            else
            {
                diagnostics.Add($"FilterGraph Run returned 0x{hr:X8}.");
            }

            diagnostics.Add("BDA graph probe completed.");
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
                tuneRequestSubmitted,
                graphRan,
                signalLocked,
                signalStrength,
                signalQuality,
                diagnostics);
        }
    }

    [SupportedOSPlatform("windows")]
    private static bool TrySubmitTuneRequest(IBaseFilter networkProvider, DvbSatelliteTv.Core.TuneRequest request, List<string> diagnostics)
    {
        if (networkProvider is not ITuner tuner)
        {
            diagnostics.Add("Network Provider does not expose ITuner.");
            return false;
        }

        IDVBSTuningSpace? tuningSpace = null;
        IDVBSLocator? locator = null;
        ITuneRequest? tuneRequest = null;

        try
        {
            tuningSpace = (IDVBSTuningSpace)new DVBSTuningSpace();
            DsError.ThrowExceptionForHR(tuningSpace.put_UniqueName("ViewerSatprof Hotbird 13E"));
            DsError.ThrowExceptionForHR(tuningSpace.put_FriendlyName("ViewerSatprof DVB-S/S2"));
            DsError.ThrowExceptionForHR(tuningSpace.put__NetworkType(DvbsNetworkProvider));
            DsError.ThrowExceptionForHR(tuningSpace.put_SystemType(DVBSystemType.Satellite));
            DsError.ThrowExceptionForHR(tuningSpace.put_LowOscillator(request.LnbLowMhz * 1000));
            DsError.ThrowExceptionForHR(tuningSpace.put_HighOscillator(request.LnbHighMhz * 1000));
            DsError.ThrowExceptionForHR(tuningSpace.put_LNBSwitch(request.SwitchMhz * 1000));

            locator = (IDVBSLocator)new DVBSLocator();
            DsError.ThrowExceptionForHR(locator.put_CarrierFrequency(request.FrequencyMhz * 1000));
            DsError.ThrowExceptionForHR(locator.put_SymbolRate(request.SymbolRateKsps));
            DsError.ThrowExceptionForHR(locator.put_SignalPolarisation(ToBdaPolarisation(request.Polarization)));
            DsError.ThrowExceptionForHR(locator.put_Modulation(ModulationType.ModNotSet));
            DsError.ThrowExceptionForHR(locator.put_InnerFEC(FECMethod.MethodNotSet));
            DsError.ThrowExceptionForHR(locator.put_InnerFECRate(BinaryConvolutionCodeRate.RateNotSet));
            DsError.ThrowExceptionForHR(locator.put_OuterFEC(FECMethod.MethodNotSet));
            DsError.ThrowExceptionForHR(locator.put_OuterFECRate(BinaryConvolutionCodeRate.RateNotSet));
            DsError.ThrowExceptionForHR(locator.put_OrbitalPosition(130));
            DsError.ThrowExceptionForHR(locator.put_WestPosition(false));

            DsError.ThrowExceptionForHR(((ITuningSpace)tuningSpace).put_DefaultLocator(locator));
            DsError.ThrowExceptionForHR(tuner.put_TuningSpace((ITuningSpace)tuningSpace));
            DsError.ThrowExceptionForHR(((ITuningSpace)tuningSpace).CreateTuneRequest(out tuneRequest));
            DsError.ThrowExceptionForHR(tuneRequest.put_Locator(locator));
            DsError.ThrowExceptionForHR(tuner.Validate(tuneRequest));
            DsError.ThrowExceptionForHR(tuner.put_TuneRequest(tuneRequest));

            diagnostics.Add($"Tune request submitted: {request.FrequencyMhz} MHz {request.Polarization}, SR {request.SymbolRateKsps} KS/s.");
            return true;
        }
        catch (Exception ex)
        {
            diagnostics.Add($"Tune request failed: {ex.GetType().Name}: {ex.Message}");
            return false;
        }
        finally
        {
            ReleaseCom(tuneRequest);
            ReleaseCom(locator);
            ReleaseCom(tuningSpace);
        }
    }

    private static void ReadSignalStatistics(
        IBaseFilter tuner,
        List<string> diagnostics,
        out bool? signalLocked,
        out int? signalStrength,
        out int? signalQuality)
    {
        signalLocked = null;
        signalStrength = null;
        signalQuality = null;

        if (tuner is not IBDA_SignalStatistics statistics)
        {
            diagnostics.Add("Tuner does not expose IBDA_SignalStatistics.");
            return;
        }

        var locked = false;
        var strength = 0;
        var quality = 0;

        var lockHr = statistics.get_SignalLocked(out locked);
        if (lockHr == 0)
        {
            signalLocked = locked;
            diagnostics.Add($"Signal locked: {locked}.");
        }
        else
        {
            diagnostics.Add($"Signal lock read returned 0x{lockHr:X8}.");
        }

        var strengthHr = statistics.get_SignalStrength(out strength);
        if (strengthHr == 0)
        {
            signalStrength = strength;
            diagnostics.Add($"Signal strength: {strength}.");
        }
        else
        {
            diagnostics.Add($"Signal strength read returned 0x{strengthHr:X8}.");
        }

        var qualityHr = statistics.get_SignalQuality(out quality);
        if (qualityHr == 0)
        {
            signalQuality = quality;
            diagnostics.Add($"Signal quality: {quality}.");
        }
        else
        {
            diagnostics.Add($"Signal quality read returned 0x{qualityHr:X8}.");
        }
    }

    private static DirectShowLib.BDA.Polarisation ToBdaPolarisation(DvbSatelliteTv.Core.Polarization polarization)
    {
        return polarization == DvbSatelliteTv.Core.Polarization.Horizontal
            ? DirectShowLib.BDA.Polarisation.LinearH
            : DirectShowLib.BDA.Polarisation.LinearV;
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
