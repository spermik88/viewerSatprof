# DVB-S/S2 Satellite TV

Windows 10 WPF prototype for DVB-S/S2 reception through a BDA-compatible Prof Revolution 7300/7301 card, targeting Hotbird 13E.

## Current Status

- WPF application shell is available.
- Prof/BDA device detection is implemented.
- Manual Hotbird 13E transponder list is loaded from JSON.
- Manual transponder entry is available.
- Tune Monitor calculates IF, 22 kHz tone, LNB 13/18V, builds a BDA graph, submits a DVB-S tune request, starts the graph, and reads lock/signal when the driver exposes it.
- Scan flow now uses the BDA recording path: tune transponder, record a short TS sample, parse services, and append channels.
- Found channels and edited transponders are stored under `%LOCALAPPDATA%\DvbSatelliteTv`.
- MPEG-TS file parsing is available from the UI through `Parse TS`.
- BDA TS recording is available from the UI through `Record TS`; it writes a short capture to `%LOCALAPPDATA%\DvbSatelliteTv\captures` and parses it when bytes are produced.
- Tune, record, and scan diagnostics include DirectShow filter pin dumps for the BDA graph.
- TS parser currently reads multi-packet PAT, PMT, SDT sections, service name, provider, video PID, audio PIDs, and basic scrambled flag.
- Built-in TV preview is still a placeholder. libVLC will be connected after a live TS path exists.

## Projects

- `DvbSatelliteTv.App` - WPF UI.
- `DvbSatelliteTv.Core` - domain models and interfaces.
- `DvbSatelliteTv.Device` - BDA/DirectShow device detection, tune monitor, and graph probe.
- `DvbSatelliteTv.Storage` - local JSON storage.
- `DvbSatelliteTv.Transport` - MPEG-TS parser.
- `DvbSatelliteTv.Tests` - unit tests for transport parsing.
- `DvbSatelliteTv.App\Data\hotbird-13e-transponders.json` - bundled offline Hotbird 13E transponder seed list.

## Key Files

- `DvbSatelliteTv.Device\BdaDeviceDetector.cs` - enumerates BDA Network Tuner, BDA Receiver/TS, and DVB-like Capture filters.
- `DvbSatelliteTv.Device\BdaGraphBuilder.cs` - creates FilterGraph, adds DVB-S Network Provider, Prof tuner, TS capture, connects pins, submits tune request, runs graph, and reads `IBDA_SignalStatistics`.
- `DvbSatelliteTv.Device\BdaTransportStreamRecorder.cs` - builds a recording graph and connects Prof TS Capture to DirectShow FileWriter.
- `DvbSatelliteTv.Device\DirectShowDiagnostics.cs` - dumps DirectShow filter pins and releases COM objects safely.
- `DvbSatelliteTv.Device\BdaDvbDevice.cs` - real scan flow over transponders using record TS then parse services.
- `DvbSatelliteTv.Device\BdaTuneMonitor.cs` - wraps manual tune diagnostics for the UI.
- `DvbSatelliteTv.Transport\TransportStreamParser.cs` - parses `.ts` files and extracts services from PAT/PMT/SDT.
- `DvbSatelliteTv.Tests\TransportStreamParserTests.cs` - synthetic TS coverage for PAT/PMT/SDT parsing.

## Next Steps

1. Verify Tune Monitor with a real Hotbird 13E dish and inspect diagnostics.
2. Verify `Tune`, `Record TS`, and `Scan Hotbird` diagnostics once a dish is connected and confirm the Prof driver can connect TS Capture to FileWriter directly.
3. If FileWriter is not accepted by the driver, replace it with a SampleGrabber/custom sink.
4. Verify `Scan Hotbird` with dish signal and tune/record diagnostics.
5. Connect libVLC for FTA channel preview.

## Commands

```powershell
dotnet build .\DvbSatelliteTv.slnx
dotnet test .\DvbSatelliteTv.slnx
dotnet run --project .\DvbSatelliteTv.App\DvbSatelliteTv.App.csproj
```
