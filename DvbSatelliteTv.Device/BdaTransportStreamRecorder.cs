using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using DirectShowLib;
using DirectShowLib.BDA;
using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

public sealed class BdaTransportStreamRecorder : ITransportStreamRecorder
{
    private static readonly Guid BdaNetworkTunerCategory = new("71985F48-1CA1-11d3-9CC8-00C04F7971E0");
    private static readonly Guid BdaReceiverComponentCategory = new("FD0A5AF4-B41D-11d2-9C95-00C04F7971E0");
    private static readonly Guid DvbsNetworkProvider = new("FA4B375A-45B4-4D45-8440-263957B11623");

    [SupportedOSPlatform("windows")]
    public async Task<TsCaptureResult> RecordAsync(TsCaptureRequest request, CancellationToken cancellationToken = default)
    {
        var diagnostics = new List<string>();
        IGraphBuilder? graph = null;
        IBaseFilter? networkProvider = null;
        IBaseFilter? tuner = null;
        IBaseFilter? transport = null;
        IBaseFilter? fileWriter = null;
        IMediaControl? mediaControl = null;
        var graphStarted = false;

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(request.OutputPath)!);
            if (File.Exists(request.OutputPath))
            {
                File.Delete(request.OutputPath);
            }

            graph = (IGraphBuilder)new FilterGraph();
            diagnostics.Add("FilterGraph created.");

            networkProvider = (IBaseFilter)new DVBSNetworkProvider();
            DsError.ThrowExceptionForHR(graph.AddFilter(networkProvider, "Microsoft DVB-S Network Provider"));
            diagnostics.Add("DVB-S Network Provider added.");
            DirectShowDiagnostics.DumpPins(networkProvider, "Network Provider", diagnostics);

            tuner = FindAndCreateFilter(BdaNetworkTunerCategory, "Prof", diagnostics)
                ?? throw new InvalidOperationException("Prof BDA tuner filter was not found.");
            DsError.ThrowExceptionForHR(graph.AddFilter(tuner, "Prof BDA Tuner/Demod"));
            diagnostics.Add("Prof BDA tuner filter added.");
            DirectShowDiagnostics.DumpPins(tuner, "Prof BDA Tuner/Demod", diagnostics);

            transport = FindAndCreateFilter(BdaReceiverComponentCategory, "Prof", diagnostics)
                ?? throw new InvalidOperationException("Prof TS capture/receiver filter was not found.");
            DsError.ThrowExceptionForHR(graph.AddFilter(transport, "Prof TS Capture"));
            diagnostics.Add("Prof TS capture/receiver filter added.");
            DirectShowDiagnostics.DumpPins(transport, "Prof TS Capture", diagnostics);

            fileWriter = (IBaseFilter)new FileWriter();
            DsError.ThrowExceptionForHR(((IFileSinkFilter)fileWriter).SetFileName(request.OutputPath, null!));
            DsError.ThrowExceptionForHR(graph.AddFilter(fileWriter, "TS File Writer"));
            diagnostics.Add($"FileWriter added: {request.OutputPath}");
            DirectShowDiagnostics.DumpPins(fileWriter, "TS File Writer", diagnostics);

            var networkProviderConnected = TryConnect(graph, networkProvider, tuner, diagnostics, "Network Provider -> Tuner pre-tune");
            SubmitTuneRequest(networkProvider, request.TuneRequest, diagnostics);
            if (!networkProviderConnected)
            {
                TryConnectOrThrow(graph, networkProvider, tuner, diagnostics, "Network Provider -> Tuner");
            }

            TryConnectOrThrow(graph, tuner, transport, diagnostics, "Tuner -> TS Capture");
            TryConnectOrThrow(graph, transport, fileWriter, diagnostics, "TS Capture -> FileWriter");

            mediaControl = (IMediaControl)graph;
            DsError.ThrowExceptionForHR(mediaControl.Run());
            graphStarted = true;
            diagnostics.Add($"Recording graph running for {request.DurationSeconds} second(s).");
            await Task.Delay(TimeSpan.FromSeconds(Math.Max(1, request.DurationSeconds)), cancellationToken);
            mediaControl.Stop();
            graphStarted = false;
            diagnostics.Add("Recording graph stopped.");

            var bytesWritten = File.Exists(request.OutputPath) ? new FileInfo(request.OutputPath).Length : 0;
            diagnostics.Add($"Bytes written: {bytesWritten}.");
            return new TsCaptureResult(bytesWritten > 0, request.OutputPath, bytesWritten, diagnostics);
        }
        catch (Exception ex)
        {
            diagnostics.Add($"TS capture failed: {ex.GetType().Name}: {ex.Message}");
            var bytesWritten = File.Exists(request.OutputPath) ? new FileInfo(request.OutputPath).Length : 0;
            return new TsCaptureResult(false, request.OutputPath, bytesWritten, diagnostics);
        }
        finally
        {
            if (graphStarted && mediaControl is not null)
            {
                try
                {
                    mediaControl.Stop();
                    diagnostics.Add("Recording graph stopped during cleanup.");
                }
                catch (Exception ex)
                {
                    diagnostics.Add($"Recording graph cleanup stop failed: {ex.GetType().Name}: {ex.Message}");
                }
            }

            ReleaseCom(fileWriter);
            ReleaseCom(transport);
            ReleaseCom(tuner);
            ReleaseCom(networkProvider);
            ReleaseCom(graph);
        }
    }

    [SupportedOSPlatform("windows")]
    private static bool TryConnect(IGraphBuilder graph, IBaseFilter upstream, IBaseFilter downstream, List<string> diagnostics, string label)
    {
        var outputPin = DsFindPin.ByDirection(upstream, PinDirection.Output, 0);
        var inputPin = DsFindPin.ByDirection(downstream, PinDirection.Input, 0);

        try
        {
            if (outputPin is null || inputPin is null)
            {
                diagnostics.Add($"{label}: pin lookup failed.");
                return false;
            }

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
    private static void SubmitTuneRequest(IBaseFilter networkProvider, DvbSatelliteTv.Core.TuneRequest request, List<string> diagnostics)
    {
        if (networkProvider is not ITuner tuner)
        {
            throw new InvalidOperationException("Network Provider does not expose ITuner.");
        }

        IDVBSTuningSpace? tuningSpace = null;
        IDVBSLocator? locator = null;
        ITuneRequest? tuneRequest = null;

        try
        {
            diagnostics.Add("Tune request: creating DVB-S tuning space.");
            tuningSpace = (IDVBSTuningSpace)new DVBSTuningSpace();
            DsError.ThrowExceptionForHR(tuningSpace.put_UniqueName("ViewerSatprof Hotbird 13E Recorder"));
            DsError.ThrowExceptionForHR(tuningSpace.put_FriendlyName("ViewerSatprof DVB-S/S2 Recorder"));
            DsError.ThrowExceptionForHR(tuningSpace.put__NetworkType(DvbsNetworkProvider));
            DsError.ThrowExceptionForHR(tuningSpace.put_SystemType(DVBSystemType.Satellite));
            DsError.ThrowExceptionForHR(tuningSpace.put_LowOscillator(request.LnbLowMhz * 1000));
            DsError.ThrowExceptionForHR(tuningSpace.put_HighOscillator(request.LnbHighMhz * 1000));
            DsError.ThrowExceptionForHR(tuningSpace.put_LNBSwitch(request.SwitchMhz * 1000));

            diagnostics.Add("Tune request: creating DVB-S locator.");
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

            diagnostics.Add("Tune request: assigning default locator.");
            DsError.ThrowExceptionForHR(((ITuningSpace)tuningSpace).put_DefaultLocator(locator));
            diagnostics.Add("Tune request: assigning tuning space to network provider.");
            DsError.ThrowExceptionForHR(tuner.put_TuningSpace((ITuningSpace)tuningSpace));
            diagnostics.Add("Tune request: creating tune request.");
            DsError.ThrowExceptionForHR(((ITuningSpace)tuningSpace).CreateTuneRequest(out tuneRequest));
            diagnostics.Add("Tune request: assigning locator.");
            DsError.ThrowExceptionForHR(tuneRequest.put_Locator(locator));
            diagnostics.Add("Tune request: validating.");
            DsError.ThrowExceptionForHR(tuner.Validate(tuneRequest));
            diagnostics.Add("Tune request: submitting to network provider.");
            DsError.ThrowExceptionForHR(tuner.put_TuneRequest(tuneRequest));
            diagnostics.Add($"Tune request submitted: {request.FrequencyMhz} MHz {request.Polarization}, SR {request.SymbolRateKsps} KS/s.");
        }
        finally
        {
            ReleaseCom(tuneRequest);
            ReleaseCom(locator);
            ReleaseCom(tuningSpace);
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
        }

        return null;
    }

    [SupportedOSPlatform("windows")]
    private static void TryConnectOrThrow(IGraphBuilder graph, IBaseFilter upstream, IBaseFilter downstream, List<string> diagnostics, string label)
    {
        var outputPin = DsFindPin.ByDirection(upstream, PinDirection.Output, 0);
        var inputPin = DsFindPin.ByDirection(downstream, PinDirection.Input, 0);

        try
        {
            if (outputPin is null || inputPin is null)
            {
                throw new InvalidOperationException($"{label}: pin lookup failed.");
            }

            var hr = graph.Connect(outputPin, inputPin);
            if (hr != 0)
            {
                diagnostics.Add($"{label}: connect failed with 0x{hr:X8}.");
                DirectShowDiagnostics.DumpPinMediaTypes(outputPin, $"{label} upstream", diagnostics);
                DirectShowDiagnostics.DumpPinMediaTypes(inputPin, $"{label} downstream", diagnostics);

                hr = TryConnectDirectWithPinTypes(graph, outputPin, inputPin, diagnostics, label);
                if (hr != 0)
                {
                    DsError.ThrowExceptionForHR(hr);
                }
            }

            diagnostics.Add($"{label}: connected.");
        }
        finally
        {
            ReleaseCom(outputPin);
            ReleaseCom(inputPin);
        }
    }

    [SupportedOSPlatform("windows")]
    private static int TryConnectDirectWithPinTypes(
        IGraphBuilder graph,
        IPin outputPin,
        IPin inputPin,
        List<string> diagnostics,
        string label)
    {
        var hr = graph.ConnectDirect(outputPin, inputPin, null!);
        diagnostics.Add($"{label}: ConnectDirect without media type returned 0x{hr:X8}.");
        if (hr == 0)
        {
            return 0;
        }

        hr = TryConnectDirectWithEnumeratedTypes(graph, outputPin, inputPin, inputPin, diagnostics, $"{label} downstream");
        if (hr == 0)
        {
            return 0;
        }

        return TryConnectDirectWithEnumeratedTypes(graph, outputPin, inputPin, outputPin, diagnostics, $"{label} upstream");
    }

    [SupportedOSPlatform("windows")]
    private static int TryConnectDirectWithEnumeratedTypes(
        IGraphBuilder graph,
        IPin outputPin,
        IPin inputPin,
        IPin mediaTypeSourcePin,
        List<string> diagnostics,
        string label)
    {
        IEnumMediaTypes? enumMediaTypes = null;
        var lastHr = -1;

        try
        {
            var hr = mediaTypeSourcePin.EnumMediaTypes(out enumMediaTypes);
            if (hr != 0 || enumMediaTypes is null)
            {
                diagnostics.Add($"{label}: cannot enumerate media types for ConnectDirect, 0x{hr:X8}.");
                return hr;
            }

            var mediaTypes = new AMMediaType[1];
            var index = 0;
            while (enumMediaTypes.Next(1, mediaTypes, IntPtr.Zero) == 0)
            {
                var mediaType = mediaTypes[0];
                try
                {
                    lastHr = graph.ConnectDirect(outputPin, inputPin, mediaType);
                    diagnostics.Add($"{label}: ConnectDirect media[{index}] returned 0x{lastHr:X8}.");
                    if (lastHr == 0)
                    {
                        return 0;
                    }
                }
                finally
                {
                    DsUtils.FreeAMMediaType(mediaType);
                }

                index++;
            }

            return lastHr;
        }
        finally
        {
            ReleaseCom(enumMediaTypes);
        }
    }

    private static DirectShowLib.BDA.Polarisation ToBdaPolarisation(DvbSatelliteTv.Core.Polarization polarization)
    {
        return polarization == DvbSatelliteTv.Core.Polarization.Horizontal
            ? DirectShowLib.BDA.Polarisation.LinearH
            : DirectShowLib.BDA.Polarisation.LinearV;
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
