using System.Runtime.Versioning;
using DirectShowLib;
using DirectShowLib.BDA;
using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Device;

internal static class BdaTopologyDiagnostics
{
    [SupportedOSPlatform("windows")]
    public static void DumpTunerCapabilities(IBaseFilter tuner, ICollection<string> diagnostics)
    {
        diagnostics.Add("BDA tuner capability probe started.");
        DumpSupportedInterface<IBDA_Topology>(tuner, diagnostics);
        DumpSupportedInterface<IBDA_DeviceControl>(tuner, diagnostics);
        DumpSupportedInterface<IBDA_FrequencyFilter>(tuner, diagnostics);
        DumpSupportedInterface<IBDA_DigitalDemodulator>(tuner, diagnostics);
        DumpSupportedInterface<IBDA_LNBInfo>(tuner, diagnostics);
        DumpSupportedInterface<IBDA_SignalProperties>(tuner, diagnostics);
        DumpSupportedInterface<IBDA_SignalStatistics>(tuner, diagnostics);
        DumpSupportedInterface<IBDA_AutoDemodulate>(tuner, diagnostics);

        if (tuner is IBDA_Topology topology)
        {
            DumpTopology(topology, diagnostics);
        }

        DumpFrequencyState(tuner, diagnostics);
        DumpDemodulatorState(tuner, diagnostics);
        DumpLnbState(tuner, diagnostics);
        DumpSignalState(tuner, diagnostics);
        diagnostics.Add("BDA tuner capability probe completed.");
    }

    [SupportedOSPlatform("windows")]
    public static bool TryApplyDirectNodeTune(IBaseFilter tuner, DvbSatelliteTv.Core.TuneRequest request, ICollection<string> diagnostics)
    {
        diagnostics.Add("BDA direct node tune started.");

        if (tuner is not IBDA_DeviceControl deviceControl || tuner is not IBDA_Topology topology)
        {
            diagnostics.Add("BDA direct node tune skipped: tuner does not expose IBDA_DeviceControl and IBDA_Topology.");
            return false;
        }

        var success = false;
        var hr = deviceControl.StartChanges();
        diagnostics.Add($"BDA direct StartChanges: hr=0x{hr:X8}.");

        try
        {
            foreach (var nodeType in GetNodeTypes(topology, diagnostics))
            {
                success |= TryApplyDirectNodeTune(topology, nodeType, request, diagnostics);
            }

            hr = deviceControl.CheckChanges();
            diagnostics.Add($"BDA direct CheckChanges: hr=0x{hr:X8}.");
            if (hr == 0)
            {
                hr = deviceControl.CommitChanges();
                diagnostics.Add($"BDA direct CommitChanges: hr=0x{hr:X8}.");
                success &= hr == 0;
            }
            else
            {
                success = false;
            }
        }
        catch (Exception ex)
        {
            diagnostics.Add($"BDA direct node tune failed: {ex.GetType().Name}: {ex.Message}");
            success = false;
        }

        diagnostics.Add($"BDA direct node tune completed: success={success}.");
        return success;
    }

    private static void DumpSupportedInterface<T>(object instance, ICollection<string> diagnostics)
    {
        diagnostics.Add($"BDA interface {typeof(T).Name}: {instance is T}");
    }

    [SupportedOSPlatform("windows")]
    private static void DumpTopology(IBDA_Topology topology, ICollection<string> diagnostics)
    {
        var nodeTypes = new int[32];
        var hr = topology.GetNodeTypes(out var nodeCount, nodeTypes.Length, nodeTypes);
        diagnostics.Add($"BDA node types: hr=0x{hr:X8}, count={nodeCount}, values={FormatInts(nodeTypes, nodeCount)}.");

        var pinTypes = new int[32];
        hr = topology.GetPinTypes(out var pinCount, pinTypes.Length, pinTypes);
        diagnostics.Add($"BDA pin types: hr=0x{hr:X8}, count={pinCount}, values={FormatInts(pinTypes, pinCount)}.");

        var nodeDescriptors = new BdaNodeDescriptor[32];
        hr = topology.GetNodeDescriptors(out var nodeDescriptorCount, nodeDescriptors.Length, nodeDescriptors);
        diagnostics.Add($"BDA node descriptors: hr=0x{hr:X8}, count={nodeDescriptorCount}.");
        if (hr == 0)
        {
            for (var i = 0; i < Math.Min(nodeDescriptorCount, nodeDescriptors.Length); i++)
            {
                var descriptor = nodeDescriptors[i];
                diagnostics.Add($"BDA node descriptor[{i}]: type={descriptor.ulBdaNodeType}, function={descriptor.guidFunction}, name={descriptor.guidName}.");
            }
        }

        var connections = new BDATemplateConnection[64];
        hr = topology.GetTemplateConnections(out var connectionCount, connections.Length, connections);
        diagnostics.Add($"BDA template connections: hr=0x{hr:X8}, count={connectionCount}.");
        if (hr == 0)
        {
            for (var i = 0; i < Math.Min(connectionCount, connections.Length); i++)
            {
                var connection = connections[i];
                diagnostics.Add(
                    $"BDA template connection[{i}]: {connection.FromNodeType}:{connection.FromNodePinType} -> {connection.ToNodeType}:{connection.ToNodePinType}.");
            }
        }

        if (hr != 0)
        {
            return;
        }

        for (var i = 0; i < Math.Min(nodeCount, nodeTypes.Length); i++)
        {
            var interfaces = new Guid[32];
            hr = topology.GetNodeInterfaces(nodeTypes[i], out var interfaceCount, interfaces.Length, interfaces);
            diagnostics.Add($"BDA node {nodeTypes[i]} interfaces: hr=0x{hr:X8}, count={interfaceCount}.");
            if (hr != 0)
            {
                continue;
            }

            for (var j = 0; j < Math.Min(interfaceCount, interfaces.Length); j++)
            {
                diagnostics.Add($"BDA node {nodeTypes[i]} interface[{j}]: {interfaces[j]}.");
            }

            DumpControlNode(topology, nodeTypes[i], diagnostics);
        }
    }

    private static IEnumerable<int> GetNodeTypes(IBDA_Topology topology, ICollection<string> diagnostics)
    {
        var nodeTypes = new int[32];
        var hr = topology.GetNodeTypes(out var nodeCount, nodeTypes.Length, nodeTypes);
        if (hr != 0)
        {
            diagnostics.Add($"BDA direct GetNodeTypes failed: hr=0x{hr:X8}.");
            yield break;
        }

        for (var i = 0; i < Math.Min(nodeCount, nodeTypes.Length); i++)
        {
            yield return nodeTypes[i];
        }
    }

    [SupportedOSPlatform("windows")]
    private static bool TryApplyDirectNodeTune(IBDA_Topology topology, int nodeType, DvbSatelliteTv.Core.TuneRequest request, ICollection<string> diagnostics)
    {
        object? controlNode = null;

        try
        {
            var hr = topology.GetControlNode(0, 1, nodeType, out controlNode);
            diagnostics.Add($"BDA direct control node {nodeType}: hr=0x{hr:X8}.");
            if (hr != 0 || controlNode is null)
            {
                return false;
            }

            var applied = false;
            if (controlNode is IBDA_LNBInfo lnbInfo)
            {
                applied |= AddWrite("LNB low oscillator", request.LnbLowMhz * 1000, value => lnbInfo.put_LocalOscilatorFrequencyLowBand(value), diagnostics);
                applied |= AddWrite("LNB high oscillator", request.LnbHighMhz * 1000, value => lnbInfo.put_LocalOscilatorFrequencyHighBand(value), diagnostics);
                applied |= AddWrite("LNB switch", request.SwitchMhz * 1000, value => lnbInfo.put_HighLowSwitchFrequency(value), diagnostics);
            }

            if (controlNode is IBDA_FrequencyFilter frequencyFilter)
            {
                applied |= AddWrite("Frequency", request.FrequencyMhz * 1000, value => frequencyFilter.put_Frequency(value), diagnostics);
                applied |= AddWrite("Polarity", ToBdaPolarisation(request.Polarization), value => frequencyFilter.put_Polarity(value), diagnostics);
            }

            if (controlNode is IBDA_DigitalDemodulator demodulator)
            {
                var symbolRate = request.SymbolRateKsps;
                var modulation = ModulationType.ModNotSet;
                var fecMethod = FECMethod.MethodNotSet;
                var fecRate = BinaryConvolutionCodeRate.RateNotSet;
                var inversion = SpectralInversion.NotSet;
                applied |= AddWrite("SymbolRate", symbolRate, value => demodulator.put_SymbolRate(ref value), diagnostics);
                applied |= AddWrite("ModulationType", modulation, value => demodulator.put_ModulationType(ref value), diagnostics);
                applied |= AddWrite("InnerFECMethod", fecMethod, value => demodulator.put_InnerFECMethod(ref value), diagnostics);
                applied |= AddWrite("InnerFECRate", fecRate, value => demodulator.put_InnerFECRate(ref value), diagnostics);
                applied |= AddWrite("SpectralInversion", inversion, value => demodulator.put_SpectralInversion(ref value), diagnostics);
            }

            return applied;
        }
        finally
        {
            DirectShowDiagnostics.ReleaseCom(controlNode);
        }
    }

    [SupportedOSPlatform("windows")]
    private static void DumpControlNode(IBDA_Topology topology, int nodeType, ICollection<string> diagnostics)
    {
        object? controlNode = null;

        try
        {
            var hr = topology.GetControlNode(0, 1, nodeType, out controlNode);
            diagnostics.Add($"BDA control node {nodeType}: hr=0x{hr:X8}, type={controlNode?.GetType().FullName ?? "(none)"}.");
            if (hr != 0 || controlNode is null)
            {
                return;
            }

            DumpSupportedInterface<IBDA_FrequencyFilter>(controlNode, diagnostics);
            DumpSupportedInterface<IBDA_DigitalDemodulator>(controlNode, diagnostics);
            DumpSupportedInterface<IBDA_LNBInfo>(controlNode, diagnostics);
            DumpSupportedInterface<IBDA_SignalStatistics>(controlNode, diagnostics);
            DumpSupportedInterface<IBDA_AutoDemodulate>(controlNode, diagnostics);
            DumpFrequencyState(controlNode, diagnostics);
            DumpDemodulatorState(controlNode, diagnostics);
            DumpLnbState(controlNode, diagnostics);
            DumpSignalState(controlNode, diagnostics);
        }
        catch (Exception ex)
        {
            diagnostics.Add($"BDA control node {nodeType}: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            DirectShowDiagnostics.ReleaseCom(controlNode);
        }
    }

    private static string FormatInts(int[] values, int count)
    {
        return string.Join(",", values.Take(Math.Min(count, values.Length)));
    }

    private static void DumpFrequencyState(object tuner, ICollection<string> diagnostics)
    {
        if (tuner is not IBDA_FrequencyFilter frequencyFilter)
        {
            return;
        }

        AddRead("Frequency", () => frequencyFilter.get_Frequency(out var value), diagnostics);
        AddRead("Autotune", () => frequencyFilter.get_Autotune(out var value), diagnostics);
        AddRead("Bandwidth", () => frequencyFilter.get_Bandwidth(out var value), diagnostics);
        AddRead("FrequencyMultiplier", () => frequencyFilter.get_FrequencyMultiplier(out var value), diagnostics);
        AddRead("Polarity", () => frequencyFilter.get_Polarity(out var value), diagnostics);
    }

    private static void DumpDemodulatorState(object tuner, ICollection<string> diagnostics)
    {
        if (tuner is not IBDA_DigitalDemodulator demodulator)
        {
            return;
        }

        AddRead("SymbolRate", () => demodulator.get_SymbolRate(out var value), diagnostics);
        AddRead("ModulationType", () => demodulator.get_ModulationType(out var value), diagnostics);
        AddRead("InnerFECMethod", () => demodulator.get_InnerFECMethod(out var value), diagnostics);
        AddRead("InnerFECRate", () => demodulator.get_InnerFECRate(out var value), diagnostics);
        AddRead("SpectralInversion", () => demodulator.get_SpectralInversion(out var value), diagnostics);
    }

    private static void DumpLnbState(object tuner, ICollection<string> diagnostics)
    {
        if (tuner is not IBDA_LNBInfo lnbInfo)
        {
            return;
        }

        AddRead("LNB low oscillator", () => lnbInfo.get_LocalOscilatorFrequencyLowBand(out var value), diagnostics);
        AddRead("LNB high oscillator", () => lnbInfo.get_LocalOscilatorFrequencyHighBand(out var value), diagnostics);
        AddRead("LNB switch", () => lnbInfo.get_HighLowSwitchFrequency(out var value), diagnostics);
    }

    private static void DumpSignalState(object tuner, ICollection<string> diagnostics)
    {
        if (tuner is not IBDA_SignalStatistics signalStatistics)
        {
            return;
        }

        AddRead("SignalPresent", () => signalStatistics.get_SignalPresent(out var value), diagnostics);
        AddRead("SignalLocked", () => signalStatistics.get_SignalLocked(out var value), diagnostics);
        AddRead("SignalStrength", () => signalStatistics.get_SignalStrength(out var value), diagnostics);
        AddRead("SignalQuality", () => signalStatistics.get_SignalQuality(out var value), diagnostics);
        AddRead("SignalSampleTime", () => signalStatistics.get_SampleTime(out var value), diagnostics);
    }

    private static void AddRead(string label, Func<int> read, ICollection<string> diagnostics)
    {
        try
        {
            var hr = read();
            diagnostics.Add($"BDA read {label}: hr=0x{hr:X8}.");
        }
        catch (Exception ex)
        {
            diagnostics.Add($"BDA read {label}: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static bool AddWrite<T>(string label, T value, Func<T, int> write, ICollection<string> diagnostics)
    {
        try
        {
            var hr = write(value);
            diagnostics.Add($"BDA write {label}={value}: hr=0x{hr:X8}.");
            return hr == 0;
        }
        catch (Exception ex)
        {
            diagnostics.Add($"BDA write {label}={value}: {ex.GetType().Name}: {ex.Message}");
            return false;
        }
    }

    private static DirectShowLib.BDA.Polarisation ToBdaPolarisation(Polarization polarization)
    {
        return polarization == Polarization.Horizontal
            ? DirectShowLib.BDA.Polarisation.LinearH
            : DirectShowLib.BDA.Polarisation.LinearV;
    }
}
