using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;

internal static class Program
{
    private enum ProbeExitCode
    {
        NotProtected = 0,
        Encrypted = 1,
        WriteProtectedOnly = 2,
        PossiblyProtected = 3,
        Corrupt = 4,
        Unsupported = 5,
        Error = 6,
    }

    private sealed class ProbeResult
    {
        public string Path { get; set; } = string.Empty;
        public string Extension { get; set; } = string.Empty;
        public string State { get; set; } = "Error";
        public string Detail { get; set; } = string.Empty;
        public bool CanOpenSafely { get; set; }
        public bool CanConvertSafely { get; set; }
        public int ExitCode { get; set; }
    }

    private static int Main(string[] args)
    {
        try
        {
            var parsed = ParseArgs(args);
            if (string.IsNullOrWhiteSpace(parsed.Path))
            {
                return Write(new ProbeResult
                {
                    State = "Error",
                    Detail = "Missing --path",
                    ExitCode = (int)ProbeExitCode.Error,
                    CanOpenSafely = false,
                    CanConvertSafely = false,
                }, parsed.Json);
            }

            var path = parsed.Path!;
            if (!File.Exists(path))
            {
                return Write(new ProbeResult
                {
                    Path = path,
                    Extension = System.IO.Path.GetExtension(path),
                    State = "Error",
                    Detail = "File not found",
                    ExitCode = (int)ProbeExitCode.Error,
                    CanOpenSafely = false,
                    CanConvertSafely = false,
                }, parsed.Json);
            }

            var result = Probe(path);
            return Write(result, parsed.Json);
        }
        catch (Exception ex)
        {
            return Write(new ProbeResult
            {
                State = "Error",
                Detail = ex.Message,
                ExitCode = (int)ProbeExitCode.Error,
                CanOpenSafely = false,
                CanConvertSafely = false,
            }, json: true);
        }
    }

    private static ProbeResult Probe(string path)
    {
        var ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".docx" or ".xlsx" or ".pptx" => ProbeOpenXml(path),
            ".xls" => ProbeLegacyXls(path),
            ".doc" => ProbeLegacyDoc(path),
            ".ppt" => ProbeLegacyPpt(path),
            _ => new ProbeResult
            {
                Path = path,
                Extension = ext,
                State = "Unsupported",
                Detail = "Unsupported extension",
                ExitCode = (int)ProbeExitCode.Unsupported,
                CanOpenSafely = false,
                CanConvertSafely = false,
            }
        };
    }

    private static ProbeResult ProbeOpenXml(string path)
    {
        try
        {
            using var archive = ZipFile.OpenRead(path);
            var names = archive.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var encrypted = names.Contains("EncryptionInfo") || names.Contains("EncryptedPackage");
            return new ProbeResult
            {
                Path = path,
                Extension = System.IO.Path.GetExtension(path),
                State = encrypted ? "Encrypted" : "NotProtected",
                Detail = encrypted ? "OOXML encrypted package markers detected" : "No OOXML encryption markers detected",
                ExitCode = encrypted ? (int)ProbeExitCode.Encrypted : (int)ProbeExitCode.NotProtected,
                CanOpenSafely = !encrypted,
                CanConvertSafely = !encrypted,
            };
        }
        catch (InvalidDataException ex)
        {
            return Corrupt(path, ex.Message);
        }
        catch (Exception ex)
        {
            return Error(path, ex.Message);
        }
    }

    private static ProbeResult ProbeLegacyXls(string path)
    {
        try
        {
            using var doc = CompoundFile.Open(path);
            var workbook = doc.TryGetStream("Workbook") ?? doc.TryGetStream("Book");
            if (workbook == null)
            {
                return Corrupt(path, "Workbook stream not found");
            }

            using var ms = new MemoryStream(workbook, writable: false);
            using var br = new BinaryReader(ms);
            while (ms.Position + 4 <= ms.Length)
            {
                ushort sid = br.ReadUInt16();
                ushort size = br.ReadUInt16();
                if (sid == 0x002F)
                {
                    return new ProbeResult
                    {
                        Path = path,
                        Extension = ".xls",
                        State = "Encrypted",
                        Detail = "BIFF FilePass record detected",
                        ExitCode = (int)ProbeExitCode.Encrypted,
                        CanOpenSafely = false,
                        CanConvertSafely = false,
                    };
                }

                if (size > ms.Length - ms.Position)
                {
                    return Corrupt(path, "Invalid BIFF record size");
                }
                ms.Position += size;
            }

            return new ProbeResult
            {
                Path = path,
                Extension = ".xls",
                State = "NotProtected",
                Detail = "Workbook stream parsed; FilePass record not found",
                ExitCode = (int)ProbeExitCode.NotProtected,
                CanOpenSafely = true,
                CanConvertSafely = true,
            };
        }
        catch (CompoundFileException ex)
        {
            return Corrupt(path, ex.Message);
        }
        catch (Exception ex)
        {
            return Error(path, ex.Message);
        }
    }

    private static ProbeResult ProbeLegacyDoc(string path)
    {
        try
        {
            using var doc = CompoundFile.Open(path);
            var word = doc.TryGetStream("WordDocument");
            if (word == null || word.Length < 12)
            {
                return Corrupt(path, "WordDocument stream missing or too short");
            }

            ushort flags = BitConverter.ToUInt16(word, 10);
            bool fEncrypted = (flags & 0x0100) != 0;
            bool fWriteReservation = (flags & 0x0800) != 0;
            bool fReadOnlyRecommended = (flags & 0x0400) != 0;
            bool fObfuscated = (flags & 0x8000) != 0;

            if (fEncrypted)
            {
                return new ProbeResult
                {
                    Path = path,
                    Extension = ".doc",
                    State = "Encrypted",
                    Detail = fObfuscated
                        ? "WordDocument stream indicates XOR obfuscation/password protection"
                        : "WordDocument stream indicates encryption/password protection",
                    ExitCode = (int)ProbeExitCode.Encrypted,
                    CanOpenSafely = false,
                    CanConvertSafely = false,
                };
            }

            if (fWriteReservation || fReadOnlyRecommended)
            {
                return new ProbeResult
                {
                    Path = path,
                    Extension = ".doc",
                    State = "WriteProtectedOnly",
                    Detail = fWriteReservation
                        ? "Write-reservation password flag detected"
                        : "Read-only recommended flag detected",
                    ExitCode = (int)ProbeExitCode.WriteProtectedOnly,
                    CanOpenSafely = false,
                    CanConvertSafely = false,
                };
            }

            return new ProbeResult
            {
                Path = path,
                Extension = ".doc",
                State = "NotProtected",
                Detail = "WordDocument FIB indicates no encryption flag",
                ExitCode = (int)ProbeExitCode.NotProtected,
                CanOpenSafely = true,
                CanConvertSafely = true,
            };
        }
        catch (CompoundFileException ex)
        {
            return Corrupt(path, ex.Message);
        }
        catch (Exception ex)
        {
            return Error(path, ex.Message);
        }
    }

    private static ProbeResult ProbeLegacyPpt(string path)
    {
        try
        {
            using var doc = CompoundFile.Open(path);

            if (doc.HasStream("EncryptedSummary"))
            {
                return new ProbeResult
                {
                    Path = path,
                    Extension = ".ppt",
                    State = "Encrypted",
                    Detail = "EncryptedSummary stream detected",
                    ExitCode = (int)ProbeExitCode.Encrypted,
                    CanOpenSafely = false,
                    CanConvertSafely = false,
                };
            }

            var currentUser = doc.TryGetStream("Current User");
            if (currentUser != null && currentUser.Length >= 20)
            {
                uint headerToken = BitConverter.ToUInt32(currentUser, 12);
                if (headerToken == 0xF3D1C4DF)
                {
                    return new ProbeResult
                    {
                        Path = path,
                        Extension = ".ppt",
                        State = "PossiblyProtected",
                        Detail = "Current User stream header token matches encrypted-PPT marker",
                        ExitCode = (int)ProbeExitCode.PossiblyProtected,
                        CanOpenSafely = false,
                        CanConvertSafely = false,
                    };
                }
            }

            return new ProbeResult
            {
                Path = path,
                Extension = ".ppt",
                State = "NotProtected",
                Detail = "No PPT encryption marker detected",
                ExitCode = (int)ProbeExitCode.NotProtected,
                CanOpenSafely = true,
                CanConvertSafely = true,
            };
        }
        catch (CompoundFileException ex)
        {
            return Corrupt(path, ex.Message);
        }
        catch (Exception ex)
        {
            return Error(path, ex.Message);
        }
    }

    private static ProbeResult Corrupt(string path, string detail) => new()
    {
        Path = path,
        Extension = System.IO.Path.GetExtension(path),
        State = "Corrupt",
        Detail = detail,
        ExitCode = (int)ProbeExitCode.Corrupt,
        CanOpenSafely = false,
        CanConvertSafely = false,
    };

    private static ProbeResult Error(string path, string detail) => new()
    {
        Path = path,
        Extension = System.IO.Path.GetExtension(path),
        State = "Error",
        Detail = detail,
        ExitCode = (int)ProbeExitCode.Error,
        CanOpenSafely = false,
        CanConvertSafely = false,
    };

    private static int Write(ProbeResult result, bool json)
    {
        if (json)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine(JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = false }));
        }
        else
        {
            Console.WriteLine(result.State);
            if (!string.IsNullOrWhiteSpace(result.Detail))
            {
                Console.Error.WriteLine(result.Detail);
            }
        }
        return result.ExitCode;
    }

    private static (string? Path, bool Json) ParseArgs(string[] args)
    {
        string? path = null;
        bool json = args.Any(a => string.Equals(a, "--json", StringComparison.OrdinalIgnoreCase));

        for (int i = 0; i < args.Length; i++)
        {
            if (string.Equals(args[i], "--path", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                path = args[i + 1];
                i++;
                continue;
            }

            if (!args[i].StartsWith("-", StringComparison.Ordinal) && path == null)
            {
                path = args[i];
            }
        }

        return (path, json);
    }

    private sealed class CompoundFileException : Exception
    {
        public CompoundFileException(string message) : base(message) { }
    }

    private sealed class CompoundFile : IDisposable
    {
        private readonly FileStream _stream;
        private readonly BinaryReader _reader;
        private readonly int _sectorSize;
        private readonly int _miniSectorSize;
        private readonly int _miniStreamCutoff;
        private readonly int _firstMiniFatSector;
        private readonly int _miniFatSectorCount;
        private readonly int _firstDifatSector;
        private readonly int _difatSectorCount;
        private readonly uint[] _fat;
        private readonly uint[] _miniFat;
        private readonly byte[] _miniStream;
        private readonly List<DirEntry> _dirs;

        private const uint EndOfChain = 0xFFFFFFFE;
        private const uint FreeSector = 0xFFFFFFFF;

        private CompoundFile(FileStream stream, BinaryReader reader, int sectorSize, int miniSectorSize,
            int miniStreamCutoff, int firstMiniFatSector, int miniFatSectorCount, int firstDifatSector,
            int difatSectorCount, uint[] fat, uint[] miniFat, byte[] miniStream, List<DirEntry> dirs)
        {
            _stream = stream;
            _reader = reader;
            _sectorSize = sectorSize;
            _miniSectorSize = miniSectorSize;
            _miniStreamCutoff = miniStreamCutoff;
            _firstMiniFatSector = firstMiniFatSector;
            _miniFatSectorCount = miniFatSectorCount;
            _firstDifatSector = firstDifatSector;
            _difatSectorCount = difatSectorCount;
            _fat = fat;
            _miniFat = miniFat;
            _miniStream = miniStream;
            _dirs = dirs;
        }

        public static CompoundFile Open(string path)
        {
            var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            try
            {
                var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true);
                if (stream.Length < 512)
                {
                    throw new CompoundFileException("Compound file is too small");
                }

                var sig = reader.ReadBytes(8);
                var expected = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
                if (!sig.SequenceEqual(expected))
                {
                    throw new CompoundFileException("Not a valid OLE compound file");
                }

                stream.Position = 30;
                ushort sectorShift = reader.ReadUInt16();
                ushort miniSectorShift = reader.ReadUInt16();
                int sectorSize = 1 << sectorShift;
                int miniSectorSize = 1 << miniSectorShift;

                stream.Position = 44;
                uint fatSectorCount = reader.ReadUInt32();
                uint firstDirSector = reader.ReadUInt32();
                _ = reader.ReadUInt32();
                int miniStreamCutoff = (int)reader.ReadUInt32();
                int firstMiniFatSector = unchecked((int)reader.ReadUInt32());
                int miniFatSectorCount = unchecked((int)reader.ReadUInt32());
                int firstDifatSector = unchecked((int)reader.ReadUInt32());
                int difatSectorCount = unchecked((int)reader.ReadUInt32());

                var difat = new List<uint>(109 + Math.Max(0, difatSectorCount * (sectorSize / 4 - 1)));
                for (int i = 0; i < 109; i++)
                {
                    uint sector = reader.ReadUInt32();
                    if (sector != FreeSector)
                    {
                        difat.Add(sector);
                    }
                }

                uint nextDifat = (uint)firstDifatSector;
                for (int n = 0; n < difatSectorCount && nextDifat != EndOfChain && nextDifat != FreeSector; n++)
                {
                    byte[] sectorBytes = ReadSector(stream, sectorSize, nextDifat);
                    int entries = sectorSize / 4 - 1;
                    for (int i = 0; i < entries; i++)
                    {
                        uint sec = BitConverter.ToUInt32(sectorBytes, i * 4);
                        if (sec != FreeSector)
                        {
                            difat.Add(sec);
                        }
                    }
                    nextDifat = BitConverter.ToUInt32(sectorBytes, sectorSize - 4);
                }

                if (difat.Count < fatSectorCount)
                {
                    throw new CompoundFileException("Incomplete DIFAT/FAT chain");
                }

                var fatEntries = new List<uint>();
                foreach (uint fatSector in difat.Take((int)fatSectorCount))
                {
                    byte[] sectorBytes = ReadSector(stream, sectorSize, fatSector);
                    for (int i = 0; i < sectorSize; i += 4)
                    {
                        fatEntries.Add(BitConverter.ToUInt32(sectorBytes, i));
                    }
                }
                uint[] fat = fatEntries.ToArray();

                byte[] dirBytes = ReadChain(stream, sectorSize, fat, firstDirSector);
                var dirs = ParseDirectories(dirBytes);
                if (dirs.Count == 0)
                {
                    throw new CompoundFileException("Directory stream is empty");
                }

                var root = dirs[0];
                byte[] miniStream = Array.Empty<byte>();
                if (root.StartSector != EndOfChain && root.StreamSize > 0)
                {
                    miniStream = ReadChain(stream, sectorSize, fat, root.StartSector, checked((int)root.StreamSize));
                }

                uint[] miniFat = Array.Empty<uint>();
                if (firstMiniFatSector != unchecked((int)EndOfChain) && miniFatSectorCount > 0)
                {
                    byte[] miniFatBytes = ReadChain(stream, sectorSize, fat, (uint)firstMiniFatSector, miniFatSectorCount * sectorSize);
                    miniFat = new uint[miniFatBytes.Length / 4];
                    for (int i = 0; i < miniFat.Length; i++)
                    {
                        miniFat[i] = BitConverter.ToUInt32(miniFatBytes, i * 4);
                    }
                }

                return new CompoundFile(stream, reader, sectorSize, miniSectorSize, miniStreamCutoff,
                    firstMiniFatSector, miniFatSectorCount, firstDifatSector, difatSectorCount, fat, miniFat, miniStream, dirs);
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        public bool HasStream(string name) => _dirs.Any(d => d.Type == 2 && string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase));

        public byte[]? TryGetStream(string name)
        {
            var entry = _dirs.FirstOrDefault(d => d.Type == 2 && string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase));
            if (entry == null)
            {
                return null;
            }

            if (entry.StreamSize == 0)
            {
                return Array.Empty<byte>();
            }

            if (entry.StreamSize < (ulong)_miniStreamCutoff)
            {
                return ReadMiniChain(entry.StartSector, checked((int)entry.StreamSize));
            }

            return ReadChain(_stream, _sectorSize, _fat, entry.StartSector, checked((int)entry.StreamSize));
        }

        private byte[] ReadMiniChain(uint startSector, int expectedLength)
        {
            if (_miniStream.Length == 0 || _miniFat.Length == 0)
            {
                throw new CompoundFileException("Mini stream metadata missing");
            }

            using var ms = new MemoryStream();
            uint sector = startSector;
            int guard = 0;
            while (sector != EndOfChain)
            {
                guard++;
                if (guard > _miniFat.Length + 8)
                {
                    throw new CompoundFileException("Mini FAT chain loop detected");
                }
                int offset = checked((int)sector * _miniSectorSize);
                if (offset < 0 || offset + _miniSectorSize > _miniStream.Length)
                {
                    throw new CompoundFileException("Mini sector offset out of range");
                }
                ms.Write(_miniStream, offset, _miniSectorSize);
                if (sector >= _miniFat.Length)
                {
                    throw new CompoundFileException("Mini FAT index out of range");
                }
                sector = _miniFat[sector];
            }

            var data = ms.ToArray();
            if (expectedLength < data.Length)
            {
                Array.Resize(ref data, expectedLength);
            }
            return data;
        }

        private static List<DirEntry> ParseDirectories(byte[] bytes)
        {
            var dirs = new List<DirEntry>();
            for (int offset = 0; offset + 128 <= bytes.Length; offset += 128)
            {
                ushort nameLen = BitConverter.ToUInt16(bytes, offset + 64);
                string name = string.Empty;
                if (nameLen >= 2 && nameLen <= 64)
                {
                    name = Encoding.Unicode.GetString(bytes, offset, nameLen - 2).TrimEnd('\0');
                }
                byte type = bytes[offset + 66];
                uint startSector = BitConverter.ToUInt32(bytes, offset + 116);
                ulong size = BitConverter.ToUInt64(bytes, offset + 120);
                dirs.Add(new DirEntry(name, type, startSector, size));
            }
            return dirs;
        }

        private static byte[] ReadSector(FileStream stream, int sectorSize, uint sector)
        {
            long offset = 512L + ((long)sector * sectorSize);
            if (offset < 0 || offset + sectorSize > stream.Length)
            {
                throw new CompoundFileException("Sector offset out of range");
            }

            byte[] buffer = new byte[sectorSize];
            stream.Position = offset;
            int read = stream.Read(buffer, 0, buffer.Length);
            if (read != buffer.Length)
            {
                throw new CompoundFileException("Unable to read full sector");
            }
            return buffer;
        }

        private static byte[] ReadChain(FileStream stream, int sectorSize, uint[] fat, uint startSector, int? expectedLength = null)
        {
            using var ms = new MemoryStream();
            uint sector = startSector;
            int guard = 0;
            while (sector != EndOfChain)
            {
                guard++;
                if (guard > fat.Length + 8)
                {
                    throw new CompoundFileException("FAT chain loop detected");
                }
                ms.Write(ReadSector(stream, sectorSize, sector));
                if (sector >= fat.Length)
                {
                    throw new CompoundFileException("FAT index out of range");
                }
                sector = fat[sector];
            }
            var data = ms.ToArray();
            if (expectedLength.HasValue && expectedLength.Value < data.Length)
            {
                Array.Resize(ref data, expectedLength.Value);
            }
            return data;
        }

        public void Dispose()
        {
            _reader.Dispose();
            _stream.Dispose();
        }

        private sealed record DirEntry(string Name, byte Type, uint StartSector, ulong StreamSize);
    }
}
