using DvbSatelliteTv.Transport;

namespace DvbSatelliteTv.Tests;

public sealed class TransportStreamParserTests
{
    [Fact]
    public async Task ParseFileAsync_extracts_services_from_multi_packet_sections()
    {
        var path = Path.Combine(Path.GetTempPath(), $"viewer-satprof-{Guid.NewGuid():N}.ts");
        await File.WriteAllBytesAsync(path, BuildTransportStream());

        try
        {
            var parser = new TransportStreamParser();
            var result = await parser.ParseFileAsync(path);

            var service = Assert.Single(result.Services);
            Assert.Equal(100, service.ServiceId);
            Assert.Equal("Demo TV", service.Name);
            Assert.Equal("Viewer", service.Provider);
            Assert.Equal(0x0100, service.PmtPid);
            Assert.Equal(0x0200, service.VideoPid);
            Assert.Contains(0x0201, service.AudioPids);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static byte[] BuildTransportStream()
    {
        var packets = new List<byte[]>();
        packets.AddRange(Packetize(0x0000, BuildPatSection()));
        packets.AddRange(Packetize(0x0100, BuildPmtSection()));
        packets.AddRange(Packetize(0x0011, BuildSdtSection()));
        return packets.SelectMany(x => x).ToArray();
    }

    private static byte[] BuildPatSection()
    {
        return WithCrc(
        [
            0x00, 0xB0, 0x0D,
            0x00, 0x01,
            0xC1,
            0x00,
            0x00,
            0x00, 0x64,
            0xE1, 0x00
        ]);
    }

    private static byte[] BuildPmtSection()
    {
        return WithCrc(
        [
            0x02, 0xB0, 0x17,
            0x00, 0x64,
            0xC1,
            0x00,
            0x00,
            0xE2, 0x00,
            0xF0, 0x00,
            0x1B, 0xE2, 0x00, 0xF0, 0x00,
            0x0F, 0xE2, 0x01, 0xF0, 0x00
        ]);
    }

    private static byte[] BuildSdtSection()
    {
        var provider = "Viewer"u8.ToArray();
        var name = "Demo TV"u8.ToArray();
        var descriptor = new List<byte> { 0x48, (byte)(3 + provider.Length + name.Length), 0x01, (byte)provider.Length };
        descriptor.AddRange(provider);
        descriptor.Add((byte)name.Length);
        descriptor.AddRange(name);

        var serviceLoopLength = descriptor.Count;
        var sectionLength = 16 + serviceLoopLength + 4;
        var section = new List<byte>
        {
            0x42,
            (byte)(0xF0 | ((sectionLength >> 8) & 0x0F)),
            (byte)(sectionLength & 0xFF),
            0x00, 0x01,
            0xC1,
            0x00,
            0x00,
            0x00, 0x01,
            0xFF,
            0x00, 0x64,
            0xFC,
            (byte)(0xF0 | ((serviceLoopLength >> 8) & 0x0F)),
            (byte)(serviceLoopLength & 0xFF)
        };
        section.AddRange(descriptor);
        return WithCrc(section);
    }

    private static IEnumerable<byte[]> Packetize(int pid, byte[] section)
    {
        var payload = new byte[1 + section.Length];
        payload[0] = 0x00;
        Array.Copy(section, 0, payload, 1, section.Length);

        var offset = 0;
        var continuity = 0;
        while (offset < payload.Length)
        {
            var packet = Enumerable.Repeat((byte)0xFF, 188).ToArray();
            packet[0] = 0x47;
            packet[1] = (byte)(((offset == 0 ? 0x40 : 0x00) | ((pid >> 8) & 0x1F)));
            packet[2] = (byte)(pid & 0xFF);
            packet[3] = (byte)(0x10 | (continuity++ & 0x0F));

            var take = Math.Min(184, payload.Length - offset);
            Array.Copy(payload, offset, packet, 4, take);
            offset += take;
            yield return packet;
        }
    }

    private static byte[] WithCrc(IEnumerable<byte> bytes)
    {
        return bytes.Concat(new byte[] { 0x00, 0x00, 0x00, 0x00 }).ToArray();
    }
}
