using System.Windows;
using System.Diagnostics;
using DvbSatelliteTv.Core;
using DvbSatelliteTv.Device;
using DvbSatelliteTv.Storage;
using DvbSatelliteTv.Transport;
using LibVLCSharp.Shared;
using Microsoft.Win32;

namespace DvbSatelliteTv.App;

public partial class MainWindow : Window
{
    private readonly IBdaDeviceDetector _bdaDeviceDetector = new BdaDeviceDetector();
    private readonly IDvbDevice _device;
    private readonly ITuneMonitor _tuneMonitor;
    private readonly ITransportStreamParser _transportStreamParser = new TransportStreamParser();
    private readonly ITransportStreamRecorder _transportStreamRecorder = new BdaTransportStreamRecorder();
    private readonly ChannelStore _channelStore;
    private readonly TransponderStore _transponderStore;
    private readonly ReceiverSettingsStore _settingsStore;
    private readonly string _diagnosticsDirectory;
    private readonly string _capturesDirectory;
    private readonly List<Transponder> _transponders = [];
    private readonly List<Channel> _channels = [];
    private readonly List<BdaFilterInfo> _bdaFilters = [];
    private readonly HashSet<string> _channelKeys = [];
    private LibVLC? _libVlc;
    private MediaPlayer? _previewPlayer;
    private string? _lastCapturePath;
    private CancellationTokenSource? _scanCancellation;
    private CancellationTokenSource? _recordCancellation;
    private ReceiverSettings _settings = ReceiverSettings.Default;

    public MainWindow()
    {
        InitializeComponent();
        var appDataPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DvbSatelliteTv");
        _device = new BdaDvbDevice(
            _bdaDeviceDetector,
            _transportStreamRecorder,
            _transportStreamParser,
            System.IO.Path.Combine(appDataPath, "scan-captures"),
            () => _settings);
        _tuneMonitor = new BdaTuneMonitor(_bdaDeviceDetector);
        var bundledTranspondersPath = System.IO.Path.Combine(
            AppContext.BaseDirectory,
            "Data",
            HotbirdDefaults.TransponderFileName);
        _channelStore = new ChannelStore(System.IO.Path.Combine(appDataPath, "channels-hotbird-13e.json"));
        _settingsStore = new ReceiverSettingsStore(System.IO.Path.Combine(appDataPath, "receiver-settings.json"));
        _diagnosticsDirectory = System.IO.Path.Combine(appDataPath, "diagnostics");
        _capturesDirectory = System.IO.Path.Combine(appDataPath, "captures");
        _transponderStore = new TransponderStore(
            System.IO.Path.Combine(appDataPath, HotbirdDefaults.TransponderFileName),
            bundledTranspondersPath);

        TranspondersGrid.ItemsSource = _transponders;
        ChannelsGrid.ItemsSource = _channels;
        BdaFiltersGrid.ItemsSource = _bdaFilters;
        Log("Application started in BDA mode.");
        Loaded += MainWindow_Loaded;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        InitializePreview();

        _settings = await _settingsStore.LoadAsync();
        ApplySettingsToUi(_settings);

        var transponders = await _transponderStore.LoadOrSeedAsync();
        _transponders.Clear();
        _transponders.AddRange(transponders);
        TranspondersGrid.Items.Refresh();

        var savedChannels = await _channelStore.LoadAsync();
        _channels.Clear();
        _channels.AddRange(savedChannels);
        RebuildChannelKeys();
        ChannelsGrid.Items.Refresh();

        Log($"Loaded {_transponders.Count} transponders, {_channels.Count} saved channels, and receiver settings.");
    }

    protected override void OnClosed(EventArgs e)
    {
        _previewPlayer?.Stop();
        PreviewVideoView.MediaPlayer = null;
        _previewPlayer?.Dispose();
        _libVlc?.Dispose();
        base.OnClosed(e);
    }

    private async void DetectButton_Click(object sender, RoutedEventArgs e)
    {
        DetectButton.IsEnabled = false;

        try
        {
            var filters = await _bdaDeviceDetector.DetectAsync();
            _bdaFilters.Clear();
            _bdaFilters.AddRange(filters);
            BdaFiltersGrid.Items.Refresh();

            var bdaSummary = filters.Count == 0
                ? "No BDA/DirectShow DVB filters found."
                : $"Found {filters.Count} BDA/DirectShow filter(s).";

            foreach (var filter in filters)
            {
                Log($"BDA filter: {filter.Category} - {filter.FriendlyName}");
            }

            var info = await _device.GetDeviceInfoAsync();
            DeviceStatusText.Text = $"{info.Name}\nDriver: {info.Driver}\nPresent: {info.IsPresent}\n{bdaSummary}\n{info.Notes}";
            Log($"Device check: {info.Name}; present={info.IsPresent}; bdaFilters={filters.Count}.");
        }
        catch (Exception ex)
        {
            DeviceStatusText.Text = $"Device detection failed\n{ex.GetType().Name}: {ex.Message}";
            Log($"Device detection failed: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            DetectButton.IsEnabled = true;
        }
    }

    private async void ScanButton_Click(object sender, RoutedEventArgs e)
    {
        await RunScanAsync(_transponders.ToList(), $"{HotbirdDefaults.Satellite.Name}", clearChannels: true);
    }

    private async void ScanSelectedButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = TranspondersGrid.SelectedItems
            .OfType<Transponder>()
            .ToList();

        if (selected.Count == 0)
        {
            Log("Selected scan was not started: no transponders selected.");
            return;
        }

        await RunScanAsync(selected, $"{selected.Count} selected transponder(s)", clearChannels: false);
    }

    private void AddTransponderButton_Click(object sender, RoutedEventArgs e)
    {
        if (!TryReadManualTransponder(out var transponder))
        {
            Log("Manual transponder was not added: frequency or symbol rate is invalid.");
            return;
        }

        _transponders.Add(transponder);
        TranspondersGrid.Items.Refresh();
        Log($"Manual transponder added: {transponder.FrequencyMhz} {transponder.Polarization} SR {transponder.SymbolRateKsps}.");
    }

    private async void TuneButton_Click(object sender, RoutedEventArgs e)
    {
        if (!TryReadManualTransponder(out var transponder))
        {
            Log("Tune was not started: frequency or symbol rate is invalid.");
            return;
        }

        TuneButton.IsEnabled = false;

        try
        {
            if (!UpdateSettingsFromUi())
            {
                Log("Tune was not started: receiver settings are invalid.");
                return;
            }

            var result = await _tuneMonitor.TuneAsync(new TuneRequest(
                transponder.FrequencyMhz,
                transponder.SymbolRateKsps,
                transponder.Polarization,
                _settings.LnbLowMhz,
                _settings.LnbHighMhz,
                _settings.LnbSwitchMhz));

            TuneIfText.Text = $"{result.IntermediateFrequencyMhz} MHz";
            TuneToneText.Text = result.Use22KhzTone ? "On" : "Off";
            TuneVoltageText.Text = $"{result.LnbVoltage}V";
            TuneStageText.Text = result.Stage;
            ScanStatusText.Text = result.Signal.Message;

            foreach (var diagnostic in result.Diagnostics)
            {
                Log($"Tune: {diagnostic}");
            }

            await SaveDiagnosticsAsync("tune", result.Diagnostics);
            Log($"Tune monitor result: canTune={result.CanTune}; {result.Signal.Message}");
        }
        catch (Exception ex)
        {
            ScanStatusText.Text = "Tune failed";
            Log($"Tune failed: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            TuneButton.IsEnabled = true;
        }
    }

    private async void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        await _channelStore.SaveAsync(_channels);
        await _transponderStore.SaveAsync(_transponders);
        if (TryReadSettings(out var settings))
        {
            _settings = settings;
            await _settingsStore.SaveAsync(_settings);
            Log($"Saved {_channels.Count} channels, {_transponders.Count} transponders, and receiver settings to local app data.");
        }
        else
        {
            Log("Channels and transponders were saved, but receiver settings are invalid and were not saved.");
        }
    }

    private async void ParseTsButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Open MPEG-TS file",
            Filter = "Transport Stream (*.ts;*.mts)|*.ts;*.mts|All files (*.*)|*.*"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        ParseTsButton.IsEnabled = false;

        try
        {
            Log($"TS parse started: {dialog.FileName}");
            var result = await _transportStreamParser.ParseFileAsync(dialog.FileName);

            _channels.Clear();
            _channelKeys.Clear();
            foreach (var channel in result.Services.Select(service => new Channel(
                         service.Name,
                         FrequencyMhz: 0,
                         SymbolRateKsps: 0,
                         Polarization.Vertical,
                         service.ServiceId,
                         service.VideoPid ?? 0,
                         service.AudioPids.FirstOrDefault(),
                         !service.IsScrambled)))
            {
                TryAddChannel(channel);
            }

            ChannelsGrid.Items.Refresh();

            foreach (var diagnostic in result.Diagnostics)
            {
                Log($"TS: {diagnostic}");
            }

            Log($"TS parse completed: {result.Services.Count} service(s).");
        }
        catch (Exception ex)
        {
            Log($"TS parse failed: {ex.Message}");
        }
        finally
        {
            ParseTsButton.IsEnabled = true;
        }
    }

    private async void RecordTsButton_Click(object sender, RoutedEventArgs e)
    {
        if (!TryReadManualTransponder(out var transponder))
        {
            Log("TS recording was not started: frequency or symbol rate is invalid.");
            return;
        }

        await RecordTransponderAsync(transponder, "manual");
    }

    private async void RecordSelectedButton_Click(object sender, RoutedEventArgs e)
    {
        if (TranspondersGrid.SelectedItem is not Transponder transponder)
        {
            Log("Selected TS recording was not started: no transponder selected.");
            return;
        }

        await RecordTransponderAsync(transponder, "selected");
    }

    private async void PreviewSelectedButton_Click(object sender, RoutedEventArgs e)
    {
        if (TranspondersGrid.SelectedItem is not Transponder transponder)
        {
            Log("Preview was not started: no transponder selected.");
            return;
        }

        await RecordTransponderAsync(transponder, "preview", playWhileRecording: true);
    }

    private async Task RecordTransponderAsync(Transponder transponder, string label, bool playWhileRecording = false)
    {
        RecordTsButton.IsEnabled = false;
        RecordSelectedButton.IsEnabled = false;
        PreviewSelectedButton.IsEnabled = false;
        CancelScanButton.IsEnabled = true;
        _recordCancellation = new CancellationTokenSource();

        try
        {
            if (!UpdateSettingsFromUi())
            {
                Log("TS recording was not started: receiver settings are invalid.");
                return;
            }

            var capturePath = System.IO.Path.Combine(
                _capturesDirectory,
                $"hotbird-{DateTime.Now:yyyyMMdd-HHmmss}-{transponder.FrequencyMhz}-{transponder.Polarization}.ts");

            Log($"TS recording started ({label}): {capturePath}");
            var recordTask = _transportStreamRecorder.RecordAsync(new TsCaptureRequest(
                new TuneRequest(
                    transponder.FrequencyMhz,
                    transponder.SymbolRateKsps,
                    transponder.Polarization,
                    _settings.LnbLowMhz,
                    _settings.LnbHighMhz,
                    _settings.LnbSwitchMhz),
                capturePath,
                _settings.CaptureSeconds),
                _recordCancellation.Token);

            if (playWhileRecording)
            {
                await StartTimeshiftPreviewAsync(capturePath, _recordCancellation.Token);
            }

            var result = await recordTask;

            foreach (var diagnostic in result.Diagnostics)
            {
                Log($"Record: {diagnostic}");
            }

            await SaveDiagnosticsAsync("record", result.Diagnostics);
            if (!result.Success)
            {
                Log($"TS recording did not produce data. Bytes written: {result.BytesWritten}.");
                return;
            }

            Log($"TS recording completed: {result.BytesWritten} byte(s). Parsing captured file.");
            _lastCapturePath = result.OutputPath;
            PlayPreview(result.OutputPath);
            var parseResult = await _transportStreamParser.ParseFileAsync(result.OutputPath);
            _channels.Clear();
            _channelKeys.Clear();
            foreach (var channel in parseResult.Services.Select(service => new Channel(
                         service.Name,
                         transponder.FrequencyMhz,
                         transponder.SymbolRateKsps,
                         transponder.Polarization,
                         service.ServiceId,
                         service.VideoPid ?? 0,
                         service.AudioPids.FirstOrDefault(),
                         !service.IsScrambled)))
            {
                TryAddChannel(channel);
            }

            ChannelsGrid.Items.Refresh();

            foreach (var diagnostic in parseResult.Diagnostics)
            {
                Log($"TS: {diagnostic}");
            }

            Log($"Captured TS parse completed: {parseResult.Services.Count} service(s).");
        }
        catch (OperationCanceledException)
        {
            Log("TS recording cancelled.");
        }
        catch (Exception ex)
        {
            Log($"TS recording failed: {ex.Message}");
        }
        finally
        {
            RecordTsButton.IsEnabled = true;
            RecordSelectedButton.IsEnabled = TranspondersGrid.SelectedItems.Count > 0;
            PreviewSelectedButton.IsEnabled = TranspondersGrid.SelectedItems.Count > 0;
            CancelScanButton.IsEnabled = _scanCancellation is not null;
            _recordCancellation?.Dispose();
            _recordCancellation = null;
        }
    }

    private void CancelScanButton_Click(object sender, RoutedEventArgs e)
    {
        _scanCancellation?.Cancel();
        _recordCancellation?.Cancel();
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        _channels.Clear();
        _channelKeys.Clear();
        ChannelsGrid.Items.Refresh();
        ScanProgressBar.Value = 0;
        ScanStatusText.Text = "Idle";
        Log("Channel list cleared.");
    }

    private void DeleteTransponderButton_Click(object sender, RoutedEventArgs e)
    {
        var selected = TranspondersGrid.SelectedItems
            .OfType<Transponder>()
            .ToList();

        if (selected.Count == 0)
        {
            Log("No transponders selected for deletion.");
            return;
        }

        foreach (var transponder in selected)
        {
            _transponders.Remove(transponder);
        }

        TranspondersGrid.Items.Refresh();
        ScanSelectedButton.IsEnabled = false;
        DeleteTransponderButton.IsEnabled = false;
        Log($"Deleted {selected.Count} transponder(s). Use Save Channels to persist the list.");
    }

    private async void ResetTranspondersButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var seed = await _transponderStore.ResetToSeedAsync();
            _transponders.Clear();
            _transponders.AddRange(seed);
            TranspondersGrid.Items.Refresh();
            Log($"Transponder list reset to seed: {_transponders.Count} item(s).");
        }
        catch (Exception ex)
        {
            Log($"Transponder reset failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private void OpenCapturesButton_Click(object sender, RoutedEventArgs e)
    {
        OpenFolder(_capturesDirectory, "captures");
    }

    private void OpenDiagnosticsButton_Click(object sender, RoutedEventArgs e)
    {
        OpenFolder(_diagnosticsDirectory, "diagnostics");
    }

    private void PlayLastCaptureButton_Click(object sender, RoutedEventArgs e)
    {
        var path = _lastCapturePath;
        if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
        {
            path = System.IO.Directory.Exists(_capturesDirectory)
                ? System.IO.Directory.GetFiles(_capturesDirectory, "*.ts")
                    .OrderByDescending(System.IO.File.GetLastWriteTime)
                    .FirstOrDefault()
                : null;
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            Log("Preview was not started: no capture file found.");
            PreviewStatusText.Text = "No capture file found";
            return;
        }

        PlayPreview(path);
    }

    private void StopPreviewButton_Click(object sender, RoutedEventArgs e)
    {
        _previewPlayer?.Stop();
        PreviewStatusText.Text = "Preview stopped";
        PreviewStatusText.Visibility = Visibility.Visible;
        Log("Preview stopped.");
    }

    private async Task RunScanAsync(IReadOnlyList<Transponder> transponders, string label, bool clearChannels)
    {
        if (transponders.Count == 0)
        {
            Log("Scan was not started: transponder list is empty.");
            return;
        }

        ScanButton.IsEnabled = false;
        ScanSelectedButton.IsEnabled = false;
        CancelScanButton.IsEnabled = true;
        ScanProgressBar.Value = 0;
        _scanCancellation = new CancellationTokenSource();

        if (clearChannels)
        {
            _channels.Clear();
            _channelKeys.Clear();
            ChannelsGrid.Items.Refresh();
        }

        try
        {
            if (!UpdateSettingsFromUi())
            {
                Log("Scan was not started: receiver settings are invalid.");
                return;
            }

            Log($"Scan started: {label}, {transponders.Count} transponder(s).");
            var completed = 0;
            var total = Math.Max(transponders.Count, 1);
            var locked = 0;
            var noSignal = 0;
            var failed = 0;
            var addedChannels = 0;

            await foreach (var progress in _device.ScanAsync(transponders, _scanCancellation.Token))
            {
                if (progress.Status == ScanStatus.Completed)
                {
                    ScanProgressBar.Value = 100;
                    ScanStatusText.Text = $"Scan completed: {locked} locked, {noSignal} no signal, {failed} failed, {addedChannels} channel(s)";
                    Log($"Scan completed: {completed}/{transponders.Count} transponder(s), locked={locked}, noSignal={noSignal}, failed={failed}, channels={addedChannels}.");
                    continue;
                }

                var isFinalTransponderStatus = progress.Status is ScanStatus.Locked or ScanStatus.NoSignal or ScanStatus.Failed;
                if (isFinalTransponderStatus)
                {
                    completed++;
                }

                if (progress.Status == ScanStatus.Locked)
                {
                    locked++;
                }
                else if (progress.Status == ScanStatus.NoSignal)
                {
                    noSignal++;
                }
                else if (progress.Status == ScanStatus.Failed)
                {
                    failed++;
                }

                ScanProgressBar.Value = isFinalTransponderStatus
                    ? Math.Min(100, completed * 100.0 / total)
                    : ScanProgressBar.Value;
                ScanStatusText.Text = $"{progress.Transponder.FrequencyMhz} MHz {progress.Transponder.Polarization}: {progress.Status}";
                Log($"{progress.Transponder.FrequencyMhz} {progress.Transponder.Polarization} SR {progress.Transponder.SymbolRateKsps}: {progress.Status}, {progress.Signal.Message}");
                foreach (var diagnostic in progress.Diagnostics ?? [])
                {
                    Log($"Scan: {diagnostic}");
                }

                foreach (var channel in progress.Channels)
                {
                    if (!TryAddChannel(channel))
                    {
                        Log($"Skipped duplicate channel: {channel.Name}");
                        continue;
                    }

                    addedChannels++;
                    Log($"Found FTA channel: {channel.Name}");
                }

                ChannelsGrid.Items.Refresh();
            }
        }
        catch (OperationCanceledException)
        {
            ScanStatusText.Text = "Scan cancelled";
            Log("Scan cancelled.");
        }
        catch (Exception ex)
        {
            ScanStatusText.Text = "Scan failed";
            Log($"Scan failed: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            ScanButton.IsEnabled = true;
            ScanSelectedButton.IsEnabled = TranspondersGrid.SelectedItems.Count > 0;
            CancelScanButton.IsEnabled = false;
            _scanCancellation?.Dispose();
            _scanCancellation = null;
        }
    }

    private void TranspondersGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        ScanSelectedButton.IsEnabled = TranspondersGrid.SelectedItems.Count > 0 && ScanButton.IsEnabled;
        RecordSelectedButton.IsEnabled = TranspondersGrid.SelectedItems.Count > 0 && RecordTsButton.IsEnabled;
        PreviewSelectedButton.IsEnabled = TranspondersGrid.SelectedItems.Count > 0 && RecordTsButton.IsEnabled;
        DeleteTransponderButton.IsEnabled = TranspondersGrid.SelectedItems.Count > 0;

        if (TranspondersGrid.SelectedItem is not Transponder transponder)
        {
            return;
        }

        FrequencyBox.Text = transponder.FrequencyMhz.ToString();
        SymbolRateBox.Text = transponder.SymbolRateKsps.ToString();
        PolarizationBox.SelectedIndex = transponder.Polarization == Polarization.Horizontal ? 0 : 1;
        Log($"Selected transponder: {transponder.FrequencyMhz} {transponder.Polarization} SR {transponder.SymbolRateKsps}.");
    }

    private void OpenFolder(string path, string label)
    {
        try
        {
            System.IO.Directory.CreateDirectory(path);
            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
            Log($"Opened {label} folder: {path}");
        }
        catch (Exception ex)
        {
            Log($"Could not open {label} folder: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private void InitializePreview()
    {
        try
        {
            LibVLCSharp.Shared.Core.Initialize();
            _libVlc = new LibVLC("--no-video-title-show");
            _previewPlayer = new MediaPlayer(_libVlc);
            PreviewVideoView.MediaPlayer = _previewPlayer;
            PreviewStatusText.Text = "No preview loaded";
            Log("libVLC preview initialized.");
        }
        catch (Exception ex)
        {
            PreviewStatusText.Text = "Preview unavailable";
            Log($"libVLC preview initialization failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private void PlayPreview(string path)
    {
        if (_libVlc is null || _previewPlayer is null)
        {
            Log("Preview was not started: libVLC is not initialized.");
            return;
        }

        if (!System.IO.File.Exists(path))
        {
            Log($"Preview was not started: file not found: {path}");
            return;
        }

        try
        {
            using var media = new Media(_libVlc, new Uri(path));
            _previewPlayer.Play(media);
            _lastCapturePath = path;
            PreviewStatusText.Text = System.IO.Path.GetFileName(path);
            PreviewStatusText.Visibility = Visibility.Collapsed;
            Log($"Preview started: {path}");
        }
        catch (Exception ex)
        {
            PreviewStatusText.Text = "Preview failed";
            PreviewStatusText.Visibility = Visibility.Visible;
            Log($"Preview failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private async Task StartTimeshiftPreviewAsync(string path, CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(TimeSpan.FromSeconds(2), cancellationToken);
            if (!System.IO.File.Exists(path))
            {
                Log("Timeshift preview is waiting: capture file was not created yet.");
                return;
            }

            PlayPreview(path);
            Log("Timeshift preview started while recording.");
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            Log($"Timeshift preview failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private bool TryReadManualTransponder(out Transponder transponder)
    {
        transponder = new Transponder(0, 0, Polarization.Vertical, string.Empty, string.Empty, string.Empty);

        if (!int.TryParse(FrequencyBox.Text, out var frequency) || !int.TryParse(SymbolRateBox.Text, out var symbolRate))
        {
            return false;
        }

        var polarization = PolarizationBox.SelectedIndex == 0 ? Polarization.Horizontal : Polarization.Vertical;
        transponder = new Transponder(frequency, symbolRate, polarization, "DVB-S/S2", "Auto", "Manual");
        return true;
    }

    private bool TryAddChannel(Channel channel)
    {
        var key = $"{channel.FrequencyMhz}:{channel.Polarization}:{channel.ServiceId}";
        if (!_channelKeys.Add(key))
        {
            return false;
        }

        _channels.Add(channel);
        return true;
    }

    private void RebuildChannelKeys()
    {
        _channelKeys.Clear();
        foreach (var channel in _channels)
        {
            _channelKeys.Add($"{channel.FrequencyMhz}:{channel.Polarization}:{channel.ServiceId}");
        }
    }

    private bool TryReadSettings(out ReceiverSettings settings)
    {
        settings = _settings;

        if (!int.TryParse(LnbLowBox.Text, out var lnbLow)
            || !int.TryParse(LnbHighBox.Text, out var lnbHigh)
            || !int.TryParse(LnbSwitchBox.Text, out var lnbSwitch)
            || !int.TryParse(CaptureSecondsBox.Text, out var captureSeconds))
        {
            return false;
        }

        if (lnbLow <= 0 || lnbHigh <= 0 || lnbSwitch <= 0 || captureSeconds <= 0 || captureSeconds > 120)
        {
            return false;
        }

        settings = new ReceiverSettings(lnbLow, lnbHigh, lnbSwitch, captureSeconds);
        return true;
    }

    private void ApplySettingsToUi(ReceiverSettings settings)
    {
        LnbLowBox.Text = settings.LnbLowMhz.ToString();
        LnbHighBox.Text = settings.LnbHighMhz.ToString();
        LnbSwitchBox.Text = settings.LnbSwitchMhz.ToString();
        CaptureSecondsBox.Text = settings.CaptureSeconds.ToString();
    }

    private bool UpdateSettingsFromUi()
    {
        if (!TryReadSettings(out var settings))
        {
            return false;
        }

        _settings = settings;
        return true;
    }

    private void Log(string message)
    {
        LogList.Items.Insert(0, $"{DateTime.Now:HH:mm:ss}  {message}");
    }

    private async Task SaveDiagnosticsAsync(string prefix, IReadOnlyList<string> diagnostics)
    {
        System.IO.Directory.CreateDirectory(_diagnosticsDirectory);
        var path = System.IO.Path.Combine(_diagnosticsDirectory, $"{prefix}-{DateTime.Now:yyyyMMdd-HHmmss}.log");
        await System.IO.File.WriteAllLinesAsync(path, diagnostics);
        Log($"Diagnostics saved: {path}");
    }
}
