
using System.IO.Compression;

namespace InStack.Excel.Builder;

public interface IZipStreamManager : IDisposable
{
    Stream CreateEntry(string name);
}

public sealed class ZipStreamManager: IZipStreamManager
{
    private ZipArchive _archive;
    public ZipStreamManager(Stream output)
    {
        _archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: false);
    }

    public Stream CreateEntry(string name)
    { 
        return
            _archive.CreateEntry(name).Open();
    }

    public void Dispose()
    {
        _archive.Dispose();
    }
}
