using InStack.Excel.Builder;

namespace Tests.Helpers
{
    internal class TestZipStreamManager : IZipStreamManager
    {
        public Dictionary<string, MemoryStream> Entries = new Dictionary<string, MemoryStream>();
        public Stream CreateEntry(string name)
        {
            var stream = new MemoryStream();
            Entries[name] = stream;

            return stream;
        }

        public void Dispose()
        {
        }
    }
}
