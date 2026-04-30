using System.Text;
using DvbSatelliteTv.Core;

namespace DvbSatelliteTv.Transport;

public sealed class TransportStreamParser : ITransportStreamParser
{
    private const int PacketSize = 188;
    private readonly Dictionary<int, int> _programMapPids = [];
    private readonly Dictionary<int, PmtInfo> _pmts = [];
    private readonly Dictionary<int, SdtInfo> _sdts = [];
    private readonly Dictionary<int, SectionAssembler> _sectionAssemblers = [];
    private readonly HashSet<string> _diagnostics = [];

    public async Task<TsParseResult> ParseFileAsync(string path, int maxPackets = 120000, CancellationToken cancellationToken = default)
    {
        _programMapPids.Clear();
        _pmts.Clear();
        _sdts.Clear();
        _sectionAssemblers.Clear();
        _diagnostics.Clear();

        var packetsRead = 0;
        await using var stream = File.OpenRead(path);
        var packet = new byte[PacketSize];

        while (packetsRead < maxPackets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var read = await stream.ReadExactlyOrLessAsync(packet, cancellationToken);
            if (read < PacketSize)
            {
                break;
            }

            packetsRead++;
            if (packet[0] != 0x47)
            {
                _diagnostics.Add($"Sync byte mismatch at packet {packetsRead}.");
                continue;
            }

            ParsePacket(packet);
        }

        var services = _programMapPids
            .OrderBy(x => x.Key)
            .Select(x =>
            {
                _pmts.TryGetValue(x.Key, out var pmt);
                _sdts.TryGetValue(x.Key, out var sdt);
                return new TsService(
                    x.Key,
                    sdt?.ServiceName ?? $"Service {x.Key}",
                    sdt?.ProviderName ?? string.Empty,
                    x.Value,
                    pmt?.PcrPid,
                    pmt?.VideoPid,
                    pmt?.AudioPids ?? [],
                    pmt?.IsScrambled ?? false);
            })
            .ToList();

        _diagnostics.Add($"Packets read: {packetsRead}.");
        _diagnostics.Add($"Programs found: {_programMapPids.Count}.");
        _diagnostics.Add($"PMTs parsed: {_pmts.Count}.");
        _diagnostics.Add($"SDT names parsed: {_sdts.Count}.");

        return new TsParseResult(path, packetsRead, services, _diagnostics.ToList());
    }

    private void ParsePacket(byte[] packet)
    {
        var payloadUnitStart = (packet[1] & 0x40) != 0;
        var pid = ((packet[1] & 0x1F) << 8) | packet[2];
        var adaptationControl = (packet[3] >> 4) & 0x03;
        if (adaptationControl is 0 or 2)
        {
            return;
        }

        var offset = 4;
        if (adaptationControl == 3)
        {
            offset += 1 + packet[offset];
        }

        if (offset >= PacketSize)
        {
            return;
        }

        if (offset >= PacketSize)
        {
            return;
        }

        foreach (var section in GetAssembler(pid).Push(packet, offset, PacketSize - offset, payloadUnitStart))
        {
            if (pid == 0)
            {
                ParsePat(section);
                continue;
            }

            if (pid == 0x11)
            {
                ParseSdt(section);
                continue;
            }

            var programId = _programMapPids.FirstOrDefault(x => x.Value == pid).Key;
            if (programId != 0)
            {
                ParsePmt(programId, section);
            }
        }
    }

    private SectionAssembler GetAssembler(int pid)
    {
        if (!_sectionAssemblers.TryGetValue(pid, out var assembler))
        {
            assembler = new SectionAssembler();
            _sectionAssemblers[pid] = assembler;
        }

        return assembler;
    }

    private void ParsePat(byte[] data)
    {
        if (data[0] != 0x00)
        {
            return;
        }

        var sectionLength = GetSectionLength(data, 0);
        var end = Math.Min(3 + sectionLength - 4, data.Length);
        var cursor = 8;

        while (cursor + 4 <= end)
        {
            var programNumber = (data[cursor] << 8) | data[cursor + 1];
            var pmtPid = ((data[cursor + 2] & 0x1F) << 8) | data[cursor + 3];
            if (programNumber != 0)
            {
                _programMapPids[programNumber] = pmtPid;
            }

            cursor += 4;
        }
    }

    private void ParsePmt(int programId, byte[] data)
    {
        if (data[0] != 0x02 || data.Length < 12)
        {
            return;
        }

        var sectionLength = GetSectionLength(data, 0);
        var end = Math.Min(3 + sectionLength - 4, data.Length);
        var pcrPid = ((data[8] & 0x1F) << 8) | data[9];
        var programInfoLength = ((data[10] & 0x0F) << 8) | data[11];
        var cursor = 12 + programInfoLength;
        int? videoPid = null;
        var audioPids = new List<int>();
        var scrambled = false;

        while (cursor + 5 <= end)
        {
            var streamType = data[cursor];
            var elementaryPid = ((data[cursor + 1] & 0x1F) << 8) | data[cursor + 2];
            var esInfoLength = ((data[cursor + 3] & 0x0F) << 8) | data[cursor + 4];
            if (IsVideo(streamType) && videoPid is null)
            {
                videoPid = elementaryPid;
            }

            if (IsAudio(streamType))
            {
                audioPids.Add(elementaryPid);
            }

            if (streamType == 0x06 && ContainsCaDescriptor(data, cursor + 5, esInfoLength))
            {
                scrambled = true;
            }

            cursor += 5 + esInfoLength;
        }

        _pmts[programId] = new PmtInfo(pcrPid, videoPid, audioPids, scrambled);
    }

    private void ParseSdt(byte[] data)
    {
        if (data[0] is not (0x42 or 0x46) || data.Length < 11)
        {
            return;
        }

        var sectionLength = GetSectionLength(data, 0);
        var end = Math.Min(3 + sectionLength - 4, data.Length);
        var cursor = 11;

        while (cursor + 5 <= end)
        {
            var serviceId = (data[cursor] << 8) | data[cursor + 1];
            var descriptorsLoopLength = ((data[cursor + 3] & 0x0F) << 8) | data[cursor + 4];
            var provider = string.Empty;
            var name = string.Empty;
            var descriptorsEnd = Math.Min(cursor + 5 + descriptorsLoopLength, end);
            var descriptorCursor = cursor + 5;

            while (descriptorCursor + 2 <= descriptorsEnd)
            {
                var tag = data[descriptorCursor];
                var length = data[descriptorCursor + 1];
                if (descriptorCursor + 2 + length > descriptorsEnd)
                {
                    break;
                }

                if (tag == 0x48 && length >= 3)
                {
                    var providerLength = data[descriptorCursor + 3];
                    var providerOffset = descriptorCursor + 4;
                    var nameLengthOffset = providerOffset + providerLength;
                    if (nameLengthOffset < descriptorCursor + 2 + length)
                    {
                        var nameLength = data[nameLengthOffset];
                        provider = DecodeDvbText(data, providerOffset, providerLength);
                        name = DecodeDvbText(data, nameLengthOffset + 1, nameLength);
                    }
                }

                descriptorCursor += 2 + length;
            }

            if (!string.IsNullOrWhiteSpace(name))
            {
                _sdts[serviceId] = new SdtInfo(provider, name);
            }

            cursor = descriptorsEnd;
        }
    }

    private static int GetSectionLength(byte[] data, int offset)
    {
        return ((data[offset + 1] & 0x0F) << 8) | data[offset + 2];
    }

    private static bool IsVideo(byte streamType)
    {
        return streamType is 0x01 or 0x02 or 0x10 or 0x1B or 0x24;
    }

    private static bool IsAudio(byte streamType)
    {
        return streamType is 0x03 or 0x04 or 0x0F or 0x11 or 0x81 or 0x87;
    }

    private static bool ContainsCaDescriptor(byte[] data, int offset, int length)
    {
        var end = Math.Min(offset + length, data.Length);
        for (var cursor = offset; cursor + 2 <= end;)
        {
            var tag = data[cursor];
            var descriptorLength = data[cursor + 1];
            if (tag == 0x09)
            {
                return true;
            }

            cursor += 2 + descriptorLength;
        }

        return false;
    }

    private static string DecodeDvbText(byte[] data, int offset, int length)
    {
        if (length <= 0 || offset < 0 || offset + length > data.Length)
        {
            return string.Empty;
        }

        return Encoding.GetEncoding("ISO-8859-1").GetString(data, offset, length).Trim('\0', ' ');
    }

    private sealed record PmtInfo(int PcrPid, int? VideoPid, IReadOnlyList<int> AudioPids, bool IsScrambled);

    private sealed record SdtInfo(string ProviderName, string ServiceName);

    private sealed class SectionAssembler
    {
        private readonly List<byte> _buffer = [];
        private int? _expectedLength;

        public IEnumerable<byte[]> Push(byte[] payload, int offset, int length, bool payloadUnitStart)
        {
            if (length <= 0)
            {
                yield break;
            }

            var cursor = offset;
            var end = offset + length;

            if (payloadUnitStart)
            {
                var pointerField = payload[cursor];
                cursor++;
                var sectionStart = cursor + pointerField;
                if (sectionStart > end)
                {
                    Reset();
                    yield break;
                }

                Reset();
                cursor = sectionStart;
            }

            while (cursor < end)
            {
                if (_buffer.Count == 0 && payload[cursor] == 0xFF)
                {
                    yield break;
                }

                var needed = _expectedLength.HasValue
                    ? _expectedLength.Value - _buffer.Count
                    : Math.Min(3 - _buffer.Count, end - cursor);
                var take = Math.Min(needed, end - cursor);
                for (var i = 0; i < take; i++)
                {
                    _buffer.Add(payload[cursor + i]);
                }

                cursor += take;

                if (!_expectedLength.HasValue && _buffer.Count >= 3)
                {
                    var sectionLength = ((_buffer[1] & 0x0F) << 8) | _buffer[2];
                    _expectedLength = 3 + sectionLength;
                    if (_expectedLength is < 3 or > 4096)
                    {
                        Reset();
                        yield break;
                    }
                }

                if (_expectedLength.HasValue && _buffer.Count >= _expectedLength.Value)
                {
                    yield return _buffer.Take(_expectedLength.Value).ToArray();
                    Reset();
                }

                if (!payloadUnitStart)
                {
                    continue;
                }

                if (cursor < end && payload[cursor] == 0xFF)
                {
                    yield break;
                }
            }
        }

        private void Reset()
        {
            _buffer.Clear();
            _expectedLength = null;
        }
    }
}

file static class StreamExtensions
{
    public static async Task<int> ReadExactlyOrLessAsync(this Stream stream, byte[] buffer, CancellationToken cancellationToken)
    {
        var total = 0;
        while (total < buffer.Length)
        {
            var read = await stream.ReadAsync(buffer.AsMemory(total, buffer.Length - total), cancellationToken);
            if (read == 0)
            {
                break;
            }

            total += read;
        }

        return total;
    }
}
