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
        ICaptureGraphBuilder2? captureGraphBuilder = null;
        IBaseFilter? networkProvider = null;
        IBaseFilter? tuner = null;
        IBaseFilter? transport = null;
        IBaseFilter? nullRenderer = null;
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
            captureGraphBuilder = (ICaptureGraphBuilder2)new CaptureGraphBuilder2();
            var hr = captureGraphBuilder.SetFiltergraph(graph);
            diagnostics.Add($"CaptureGraphBuilder2 SetFiltergraph returned 0x{hr:X8}.");

            networkProvider = (IBaseFilter)new DVBSNetworkProvider();
            hr = graph.AddFilter(networkProvider, "Microsoft DVB-S Network Provider");
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
            BdaTopologyDiagnostics.DumpTunerCapabilities(tuner, diagnostics);

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
            if (!tunerConnected && captureGraphBuilder is not null)
            {
                tunerConnected = TryRenderStream(captureGraphBuilder, networkProvider, tuner, diagnostics, "Network Provider -> Tuner");
            }

            tuneRequestSubmitted = TrySubmitTuneRequest(networkProvider, request, diagnostics);
            if (tuneRequestSubmitted && !tunerConnected)
            {
                tunerConnected = TryConnect(graph, networkProvider, tuner, diagnostics, "Network Provider -> Tuner after tune");
                if (!tunerConnected && captureGraphBuilder is not null)
                {
                    tunerConnected = TryRenderStream(captureGraphBuilder, networkProvider, tuner, diagnostics, "Network Provider -> Tuner after tune");
                }
            }

            transportConnected = TryConnect(graph, tuner, transport, diagnostics, "Tuner -> TS Capture");
            if (!transportConnected && captureGraphBuilder is not null)
            {
                transportConnected = TryRenderStream(captureGraphBuilder, tuner, transport, diagnostics, "Tuner -> TS Capture");
            }

            nullRenderer = (IBaseFilter)new NullRenderer();
            hr = graph.AddFilter(nullRenderer, "TS Null Renderer");
            diagnostics.Add($"TS Null Renderer AddFilter returned 0x{hr:X8}.");
            if (hr == 0)
            {
                DirectShowDiagnostics.DumpPins(nullRenderer, "TS Null Renderer", diagnostics);
                TryConnect(graph, transport, nullRenderer, diagnostics, "TS Capture -> NullRenderer");
            }

            var mediaControl = (IMediaControl)graph;
            hr = mediaControl.Run();
            if (hr >= 0)
            {
                graphRan = true;
                diagnostics.Add($"FilterGraph running, Run returned 0x{hr:X8}.");
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
            ReleaseCom(nullRenderer);
            ReleaseCom(transport);
            ReleaseCom(tuner);
            ReleaseCom(networkProvider);
            ReleaseCom(captureGraphBuilder);
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

            TryRegisterTuningSpace((ITuningSpace)tuningSpace, diagnostics);
            if (!TryApplyTuningSpace(tuner, (ITuningSpace)tuningSpace, locator, diagnostics, "registered/default-locator-first"))
            {
                diagnostics.Add("Tune request: retrying with tuning-space-first order.");
                DsError.ThrowExceptionForHR(tuner.put_TuningSpace((ITuningSpace)tuningSpace));
                diagnostics.Add("Tune request: tuning-space-first put_TuningSpace succeeded.");
                DsError.ThrowExceptionForHR(((ITuningSpace)tuningSpace).put_DefaultLocator(locator));
                diagnostics.Add("Tune request: tuning-space-first default locator assigned.");
            }

            DsError.ThrowExceptionForHR(((ITuningSpace)tuningSpace).CreateTuneRequest(out tuneRequest));
            diagnostics.Add("Tune request: CreateTuneRequest succeeded.");
            DsError.ThrowExceptionForHR(tuneRequest.put_Locator(locator));
            diagnostics.Add("Tune request: tune request locator assigned.");
            DsError.ThrowExceptionForHR(tuner.Validate(tuneRequest));
            diagnostics.Add("Tune request: Validate succeeded.");
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

    [SupportedOSPlatform("windows")]
    private static bool TryApplyTuningSpace(
        ITuner tuner,
        ITuningSpace tuningSpace,
        IDVBSLocator locator,
        List<string> diagnostics,
        string label)
    {
        try
        {
            DsError.ThrowExceptionForHR(tuningSpace.put_DefaultLocator(locator));
            diagnostics.Add($"Tune request {label}: default locator assigned.");
            DsError.ThrowExceptionForHR(tuner.put_TuningSpace(tuningSpace));
            diagnostics.Add($"Tune request {label}: put_TuningSpace succeeded.");
            return true;
        }
        catch (Exception ex)
        {
            diagnostics.Add($"Tune request {label}: failed: {ex.GetType().Name}: {ex.Message}");
            return false;
        }
    }

    [SupportedOSPlatform("windows")]
    private static void TryRegisterTuningSpace(ITuningSpace tuningSpace, List<string> diagnostics)
    {
        ITuningSpaceContainer? container = null;

        try
        {
            container = (ITuningSpaceContainer)new SystemTuningSpaces();
            var hr = container.FindID(tuningSpace, out var existingId);
            diagnostics.Add($"SystemTuningSpaces FindID returned 0x{hr:X8}, id={existingId}.");
            if (hr == 0 && existingId > 0)
            {
                return;
            }

            hr = container.Add(tuningSpace, out var newId);
            diagnostics.Add($"SystemTuningSpaces Add returned 0x{hr:X8}, id={newId}.");
        }
        catch (Exception ex)
        {
            diagnostics.Add($"SystemTuningSpaces registration failed: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            ReleaseCom(container);
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
    private static bool TryRenderStream(
        ICaptureGraphBuilder2 captureGraphBuilder,
        IBaseFilter upstream,
        IBaseFilter downstream,
        List<string> diagnostics,
        string label)
    {
        try
        {
            var hr = captureGraphBuilder.RenderStream(DsGuid.Empty, DsGuid.Empty, upstream, null!, downstream);
            if (hr == 0)
            {
                diagnostics.Add($"{label}: CaptureGraphBuilder2 RenderStream connected.");
                return true;
            }

            diagnostics.Add($"{label}: CaptureGraphBuilder2 RenderStream returned 0x{hr:X8}.");
            return false;
        }
        catch (Exception ex)
        {
            diagnostics.Add($"{label}: CaptureGraphBuilder2 RenderStream failed: {ex.GetType().Name}: {ex.Message}");
            return false;
        }
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
