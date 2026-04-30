using System.Windows;
using DvbSatelliteTv.Core;
using DvbSatelliteTv.Device;
using DvbSatelliteTv.Storage;
using DvbSatelliteTv.Transport;
using Microsoft.Win32;

namespace DvbSatelliteTv.App;

public partial class MainWindow : Window
{
    private readonly IBdaDeviceDetector _bdaDeviceDetector = new BdaDeviceDetector();
    private readonly IDvbDevice _device;
    private readonly ITuneMonitor _tuneMonitor;
    private readonly ITransportStreamParser _transportStreamParser = new TransportStreamParser();
    private readonly ChannelStore _channelStore;
    private readonly TransponderStore _transponderStore;
    private readonly List<Transponder> _transponders = [];
    private readonly List<Channel> _channels = [];
    private readonly List<BdaFilterInfo> _bdaFilters = [];
    private CancellationTokenSource? _scanCancellation;

    public MainWindow()
    {
        InitializeComponent();
        _device = new BdaDiagnosticDevice(_bdaDeviceDetector);
        _tuneMonitor = new BdaTuneMonitor(_bdaDeviceDetector);
        var appDataPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DvbSatelliteTv");
        var bundledTranspondersPath = System.IO.Path.Combine(
            AppContext.BaseDirectory,
            "Data",
            HotbirdDefaults.TransponderFileName);
        _channelStore = new ChannelStore(System.IO.Path.Combine(appDataPath, "channels-hotbird-13e.json"));
        _transponderStore = new TransponderStore(
            System.IO.Path.Combine(appDataPath, HotbirdDefaults.TransponderFileName),
            bundledTranspondersPath);

        TranspondersGrid.ItemsSource = _transponders;
        ChannelsGrid.ItemsSource = _channels;
        BdaFiltersGrid.ItemsSource = _bdaFilters;
        Log("Application started in simulator mode.");
        Loaded += MainWindow_Loaded;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        var transponders = await _transponderStore.LoadOrSeedAsync();
        _transponders.Clear();
        _transponders.AddRange(transponders);
        TranspondersGrid.Items.Refresh();

        var savedChannels = await _channelStore.LoadAsync();
        _channels.Clear();
        _channels.AddRange(savedChannels);
        ChannelsGrid.Items.Refresh();

        Log($"Loaded {_transponders.Count} transponders and {_channels.Count} saved channels.");
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
        finally
        {
            DetectButton.IsEnabled = true;
        }
    }

    private async void ScanButton_Click(object sender, RoutedEventArgs e)
    {
        ScanButton.IsEnabled = false;
        CancelScanButton.IsEnabled = true;
        ScanProgressBar.Value = 0;
        _scanCancellation = new CancellationTokenSource();
        _channels.Clear();
        ChannelsGrid.Items.Refresh();

        try
        {
            Log($"Scan started: {HotbirdDefaults.Satellite.Name}, {_transponders.Count} transponders.");
            var completed = 0;
            var total = Math.Max(_transponders.Count, 1);

            await foreach (var progress in _device.ScanAsync(_transponders, _scanCancellation.Token))
            {
                if (progress.Status == ScanStatus.Completed)
                {
                    ScanProgressBar.Value = 100;
                    ScanStatusText.Text = "Scan completed";
                    Log("Scan completed.");
                    continue;
                }

                completed++;
                ScanProgressBar.Value = Math.Min(100, completed * 100.0 / total);
                ScanStatusText.Text = $"{progress.Transponder.FrequencyMhz} MHz {progress.Transponder.Polarization}: {progress.Status}";
                Log($"{progress.Transponder.FrequencyMhz} {progress.Transponder.Polarization} SR {progress.Transponder.SymbolRateKsps}: {progress.Status}, {progress.Signal.Message}");

                foreach (var channel in progress.Channels)
                {
                    _channels.Add(channel);
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
        finally
        {
            ScanButton.IsEnabled = true;
            CancelScanButton.IsEnabled = false;
        }
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
            var result = await _tuneMonitor.TuneAsync(new TuneRequest(
                transponder.FrequencyMhz,
                transponder.SymbolRateKsps,
                transponder.Polarization,
                LnbLowMhz: 9750,
                LnbHighMhz: 10600,
                SwitchMhz: 11700));

            TuneIfText.Text = $"{result.IntermediateFrequencyMhz} MHz";
            TuneToneText.Text = result.Use22KhzTone ? "On" : "Off";
            TuneVoltageText.Text = $"{result.LnbVoltage}V";
            TuneStageText.Text = result.Stage;
            ScanStatusText.Text = result.Signal.Message;

            foreach (var diagnostic in result.Diagnostics)
            {
                Log($"Tune: {diagnostic}");
            }

            Log($"Tune monitor result: canTune={result.CanTune}; {result.Signal.Message}");
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
        Log($"Saved {_channels.Count} channels and {_transponders.Count} transponders to local app data.");
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
            _channels.AddRange(result.Services.Select(service => new Channel(
                service.Name,
                FrequencyMhz: 0,
                SymbolRateKsps: 0,
                Polarization.Vertical,
                service.ServiceId,
                service.VideoPid ?? 0,
                service.AudioPids.FirstOrDefault(),
                !service.IsScrambled)));
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

    private void CancelScanButton_Click(object sender, RoutedEventArgs e)
    {
        _scanCancellation?.Cancel();
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        _channels.Clear();
        ChannelsGrid.Items.Refresh();
        ScanProgressBar.Value = 0;
        ScanStatusText.Text = "Idle";
        Log("Channel list cleared.");
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

    private void Log(string message)
    {
        LogList.Items.Insert(0, $"{DateTime.Now:HH:mm:ss}  {message}");
    }
}
