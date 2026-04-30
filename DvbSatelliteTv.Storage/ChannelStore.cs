using System.Text.Json;
using System.Text.Json.Serialization;
using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Storage;

public sealed class ChannelStore
{
    private static readonly JsonSerializerOptions Options = JsonOptionsFactory.CreateJsonOptions();
    private readonly string _path;

    public ChannelStore(string path)
    {
        _path = path;
    }

    public async Task SaveAsync(IEnumerable<Channel> channels, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        await using var stream = File.Create(_path);
        await JsonSerializer.SerializeAsync(stream, channels.OrderBy(x => x.FrequencyMhz).ThenBy(x => x.Name), Options, cancellationToken);
    }

    public async Task<IReadOnlyList<Channel>> LoadAsync(CancellationToken cancellationToken = default)
    {
        if (!File.Exists(_path))
        {
            return [];
        }

        await using var stream = File.OpenRead(_path);
        return await JsonSerializer.DeserializeAsync<List<Channel>>(stream, Options, cancellationToken) ?? [];
    }
}

public sealed class TransponderStore
{
    private static readonly JsonSerializerOptions Options = JsonOptionsFactory.CreateJsonOptions();
    private readonly string _path;
    private readonly string? _seedPath;

    public TransponderStore(string path, string? seedPath = null)
    {
        _path = path;
        _seedPath = seedPath;
    }

    public async Task<IReadOnlyList<Transponder>> LoadOrSeedAsync(CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);

        if (!File.Exists(_path))
        {
            var seed = await LoadSeedAsync(cancellationToken);
            await SaveAsync(seed, cancellationToken);
            return seed;
        }

        await using var stream = File.OpenRead(_path);
        var transponders = await JsonSerializer.DeserializeAsync<List<Transponder>>(stream, Options, cancellationToken);
        return transponders is { Count: > 0 } ? transponders : HotbirdDefaults.FallbackTransponders;
    }

    private async Task<IReadOnlyList<Transponder>> LoadSeedAsync(CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(_seedPath) && File.Exists(_seedPath))
        {
            await using var stream = File.OpenRead(_seedPath);
            var transponders = await JsonSerializer.DeserializeAsync<List<Transponder>>(stream, Options, cancellationToken);
            if (transponders is { Count: > 0 })
            {
                return transponders;
            }
        }

        return HotbirdDefaults.FallbackTransponders;
    }

    public async Task SaveAsync(IEnumerable<Transponder> transponders, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        await using var stream = File.Create(_path);
        await JsonSerializer.SerializeAsync(
            stream,
            transponders.OrderBy(x => x.FrequencyMhz).ThenBy(x => x.Polarization),
            Options,
            cancellationToken);
    }
}

public sealed class ReceiverSettingsStore
{
    private static readonly JsonSerializerOptions Options = JsonOptionsFactory.CreateJsonOptions();
    private readonly string _path;

    public ReceiverSettingsStore(string path)
    {
        _path = path;
    }

    public async Task<ReceiverSettings> LoadAsync(CancellationToken cancellationToken = default)
    {
        if (!File.Exists(_path))
        {
            await SaveAsync(ReceiverSettings.Default, cancellationToken);
            return ReceiverSettings.Default;
        }

        try
        {
            await using var stream = File.OpenRead(_path);
            return await JsonSerializer.DeserializeAsync<ReceiverSettings>(stream, Options, cancellationToken)
                ?? ReceiverSettings.Default;
        }
        catch
        {
            return ReceiverSettings.Default;
        }
    }

    public async Task SaveAsync(ReceiverSettings settings, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        await using var stream = File.Create(_path);
        await JsonSerializer.SerializeAsync(stream, settings, Options, cancellationToken);
    }
}

file static class JsonOptionsFactory
{
    public static JsonSerializerOptions CreateJsonOptions()
    {
        var options = new JsonSerializerOptions { WriteIndented = true };
        options.Converters.Add(new JsonStringEnumConverter());
        return options;
    }
}
