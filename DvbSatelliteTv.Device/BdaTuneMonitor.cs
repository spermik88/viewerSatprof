using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

public sealed class BdaTuneMonitor : ITuneMonitor
{
    private readonly IBdaDeviceDetector _detector;
    private readonly IBdaGraphBuilder _graphBuilder;

    public BdaTuneMonitor(IBdaDeviceDetector detector, IBdaGraphBuilder? graphBuilder = null)
    {
        _detector = detector;
        _graphBuilder = graphBuilder ?? new BdaGraphBuilder();
    }

    public async Task<TuneResult> TuneAsync(TuneRequest request, CancellationToken cancellationToken = default)
    {
        var diagnostics = new List<string>();
        var filters = await _detector.DetectAsync(cancellationToken);
        var profFilters = filters
            .Where(x => IsProfFilter(x.FriendlyName) || IsProfFilter(x.DevicePath))
            .ToList();

        var tuner = profFilters.FirstOrDefault(x => x.Category.Contains("Tuner", StringComparison.OrdinalIgnoreCase));
        var transport = profFilters.FirstOrDefault(x =>
            x.Category.Contains("Receiver", StringComparison.OrdinalIgnoreCase)
            || x.FriendlyName.Contains("TS", StringComparison.OrdinalIgnoreCase)
            || x.FriendlyName.Contains("Capture", StringComparison.OrdinalIgnoreCase));

        diagnostics.Add($"Requested: {request.FrequencyMhz} MHz {request.Polarization}, SR {request.SymbolRateKsps} KS/s");
        diagnostics.Add($"LNB: low {request.LnbLowMhz} MHz, high {request.LnbHighMhz} MHz, switch {request.SwitchMhz} MHz");

        var useHighBand = request.FrequencyMhz >= request.SwitchMhz;
        var localOscillator = useHighBand ? request.LnbHighMhz : request.LnbLowMhz;
        var intermediateFrequency = Math.Abs(request.FrequencyMhz - localOscillator);
        var voltage = request.Polarization == Polarization.Horizontal ? 18 : 13;

        diagnostics.Add($"Band: {(useHighBand ? "high" : "low")}; IF {intermediateFrequency} MHz; 22 kHz {(useHighBand ? "on" : "off")}; LNB voltage {voltage}V");

        if (tuner is null)
        {
            diagnostics.Add("Prof BDA tuner filter was not found.");
            return CreateResult(false, "BDA filter check", intermediateFrequency, useHighBand, voltage, "Prof BDA tuner filter is missing.", diagnostics);
        }

        diagnostics.Add($"Tuner: {tuner.FriendlyName}");

        if (transport is null)
        {
            diagnostics.Add("Prof TS capture/receiver filter was not found.");
            return CreateResult(false, "BDA filter check", intermediateFrequency, useHighBand, voltage, "Prof TS capture/receiver filter is missing.", diagnostics);
        }

        diagnostics.Add($"Transport: {transport.FriendlyName}");
        diagnostics.Add("Starting BDA graph probe.");
        var graphProbe = await _graphBuilder.ProbeAsync(request, cancellationToken);
        foreach (var graphDiagnostic in graphProbe.Diagnostics)
        {
            diagnostics.Add($"Graph: {graphDiagnostic}");
        }

        var canTune = graphProbe.GraphCreated
            && graphProbe.NetworkProviderAdded
            && graphProbe.TunerAdded
            && graphProbe.TransportAdded
            && graphProbe.TuneRequestSubmitted;
        var message = graphProbe.SignalLocked.HasValue
            ? $"Tune request submitted; lock={graphProbe.SignalLocked.Value}."
            : graphProbe.GraphRan
                ? "Tune request submitted and graph ran; signal lock was not reported."
                : graphProbe.TuneRequestSubmitted
                    ? "Tune request submitted, but graph did not run."
                    : "BDA graph probe completed, but tune request was not submitted.";

        return new TuneResult(
            canTune,
            graphProbe.TransportConnected ? "Graph connected" : "Graph probe",
            intermediateFrequency,
            useHighBand,
            voltage,
            new SignalInfo(
                graphProbe.SignalLocked == true,
                NormalizeSignalValue(graphProbe.SignalStrength),
                NormalizeSignalValue(graphProbe.SignalQuality),
                message),
            diagnostics);
    }

    private static TuneResult CreateResult(
        bool canTune,
        string stage,
        int intermediateFrequencyMhz,
        bool use22KhzTone,
        int lnbVoltage,
        string message,
        IReadOnlyList<string> diagnostics)
    {
        return new TuneResult(
            canTune,
            stage,
            intermediateFrequencyMhz,
            use22KhzTone,
            lnbVoltage,
            new SignalInfo(false, 0, 0, message),
            diagnostics);
    }

    private static int NormalizeSignalValue(int? rawValue)
    {
        if (!rawValue.HasValue)
        {
            return 0;
        }

        return Math.Clamp(rawValue.Value, 0, 100);
    }

    private static bool IsProfFilter(string value)
    {
        return value.Contains("Prof", StringComparison.OrdinalIgnoreCase)
            || value.Contains("ven_14f1", StringComparison.OrdinalIgnoreCase)
            || value.Contains("3034b034", StringComparison.OrdinalIgnoreCase);
    }
}
