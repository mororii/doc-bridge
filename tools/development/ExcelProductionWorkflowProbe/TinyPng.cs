namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class TinyPng
{
    public static string WriteSample(string directory, string name, byte r, byte g, byte b)
    {
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, name);
        if (File.Exists(path)) return path;
        File.WriteAllBytes(path, Rgb(r, g, b));
        return path;
    }

    /// <summary>Minimal 1×1 truecolor PNG. Marked as a sample fixture, not site photography.</summary>
    public static byte[] Rgb(byte r, byte g, byte b)
    {
        static uint Crc(byte[] data)
        {
            var table = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                var c = i;
                for (var k = 0; k < 8; k++)
                    c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
                table[i] = c;
            }
            var crc = 0xFFFFFFFFu;
            foreach (var x in data) crc = table[(crc ^ x) & 0xFF] ^ (crc >> 8);
            return crc ^ 0xFFFFFFFFu;
        }

        static byte[] Be(uint v) =>
        [
            (byte)(v >> 24), (byte)(v >> 16), (byte)(v >> 8), (byte)v
        ];

        var raw = new byte[] { 0x00, r, g, b };
        using var ms = new MemoryStream();
        using (var deflate = new System.IO.Compression.ZLibStream(ms, System.IO.Compression.CompressionLevel.SmallestSize, leaveOpen: true))
            deflate.Write(raw);
        var compressed = ms.ToArray();

        byte[] Chunk(string type, byte[] data)
        {
            var typeBytes = Encoding.ASCII.GetBytes(type);
            var crcSrc = typeBytes.Concat(data).ToArray();
            return Be((uint)data.Length).Concat(typeBytes).Concat(data).Concat(Be(Crc(crcSrc))).ToArray();
        }

        var ihdr = new byte[]
        {
            0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0,
        };
        var sig = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        return sig.Concat(Chunk("IHDR", ihdr)).Concat(Chunk("IDAT", compressed)).Concat(Chunk("IEND", [])).ToArray();
    }
}
