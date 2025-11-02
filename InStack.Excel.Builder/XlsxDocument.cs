using ExcelUtils.Builder.RowExtensions;
using System.Text;

namespace InStack.Excel.Builder;

internal record SheetInfo(string Name, string FileName, string RId, uint Id);

public class XlsxDocument : IDisposable
{
    private readonly List<SheetInfo> _sheets = [];
    private uint _nextSheetId = 1;
    private IZipStreamManager ZipStreamManager { get; } 

    public XlsxDocument(string filePath)
        : this(File.Create(filePath)) {}

    public XlsxDocument(Stream output)
        : this(new ZipStreamManager(output)){}

    public XlsxDocument(IZipStreamManager zipStreamManager)
    {
        ZipStreamManager = zipStreamManager;

        using (var entryStream = ZipStreamManager.CreateEntry("_rels/.rels"))
        {
            entryStream.Write(@"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
	<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml"" Id=""rId1"" />
</Relationships>"u8);
        }
    }

    public Stream CreateStyleEntry()
    {
        return ZipStreamManager.CreateEntry("xl/styles.xml");
    }

    public Sheet AddSheet(string sheetName, SheetConfig? config = null)
    {
        var sheetInfo = new SheetInfo(sheetName, $"sheet{_nextSheetId}.xml", $"rId{_nextSheetId}", _nextSheetId);
        _sheets.Add(sheetInfo);

        var entryStream = ZipStreamManager.CreateEntry($"xl/worksheets/{sheetInfo.FileName}");
        var sheet = new Sheet(entryStream, config ?? new SheetConfig());

        _nextSheetId++;

        return sheet;
    }

    public void Dispose()
    {
        using (var entryStream = ZipStreamManager.CreateEntry("xl/workbook.xml"))
        {
            entryStream.Write(@"<?xml version=""1.0"" encoding=""utf-8""?><workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheets>"u8);

            foreach (var sheetInfo in _sheets)
            {
                entryStream.Write(Encoding.UTF8.GetBytes(
                    @$"<sheet name=""{sheetInfo.Name}"" sheetId=""{sheetInfo.Id}"" r:id=""{sheetInfo.RId}""/>"));
            }

            entryStream.Write(@"</sheets></workbook>"u8);
        }

        using (var entryStream = ZipStreamManager.CreateEntry("xl/_rels/workbook.xml.rels"))
        {
            entryStream.Write(@"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">"u8);

            foreach (var sheetInfo in _sheets)
            {
                entryStream.Write(Encoding.UTF8.GetBytes(
                    $@"<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/{sheetInfo.FileName}"" Id=""{sheetInfo.RId}"" />"));
            }
    
            entryStream.Write(Encoding.UTF8.GetBytes($@"<Relationship Type =""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml"" Id=""rId{_nextSheetId}"" />"));
            entryStream.Write("</Relationships>"u8);
        }

        using (var entryStream = ZipStreamManager.CreateEntry("[Content_Types].xml"))
        {
            entryStream.Write(@"<?xml version=""1.0"" encoding=""utf-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""xml"" ContentType=""application/xml"" />
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml"" />
<Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"" />
<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>"u8);

            foreach (var sheetInfo in _sheets)
            {
                entryStream.Write(Encoding.UTF8.GetBytes(
                    $@"<Override PartName=""/xl/worksheets/{sheetInfo.FileName}"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"" />"));
            }
                
            entryStream.Write("</Types>"u8);
        }

        ZipStreamManager.Dispose();
    }

}