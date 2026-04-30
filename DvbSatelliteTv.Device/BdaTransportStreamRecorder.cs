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

            TryConnectOrThrow(graph, networkProvider, tuner, diagnostics, "Network Provider -> Tuner");
            TryConnectOrThrow(graph, tuner, transport, diagnostics, "Tuner -> TS Capture");
            TryConnectOrThrow(graph, transport, fileWriter, diagnostics, "TS Capture -> FileWriter");
            SubmitTuneRequest(networkProvider, request.TuneRequest, diagnostics);

            var mediaControl = (IMediaControl)graph;
            DsError.ThrowExceptionForHR(mediaControl.Run());
            diagnostics.Add($"Recording graph running for {request.DurationSeconds} second(s).");
            await Task.Delay(TimeSpan.FromSeconds(Math.Max(1, request.DurationSeconds)), cancellationToken);
            mediaControl.Stop();
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
            ReleaseCom(fileWriter);
            ReleaseCom(transport);
            ReleaseCom(tuner);
            ReleaseCom(networkProvider);
            ReleaseCom(graph);
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
            tuningSpace = (IDVBSTuningSpace)new DVBSTuningSpace();
            DsError.ThrowExceptionForHR(tuningSpace.put_UniqueName("ViewerSatprof Hotbird 13E Recorder"));
            DsError.ThrowExceptionForHR(tuningSpace.put_FriendlyName("ViewerSatprof DVB-S/S2 Recorder"));
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
            DsError.ThrowExceptionForHR(hr);
            diagnostics.Add($"{label}: connected.");
        }
        finally
        {
            ReleaseCom(outputPin);
            ReleaseCom(inputPin);
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
