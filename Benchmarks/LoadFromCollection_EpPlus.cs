using BenchmarkDotNet.Attributes;
using Examples.CollectionExample;
using InStack.Excel.Builder;
using InStack.Excel.OpenXmlStyles;
using OfficeOpenXml;

namespace Benchmarks;


[MemoryDiagnoser]
public class BenchmarkEpPlusOpenXml
{
    private List<CollectionItem> _source;

    private MemoryStream _stream;

    [GlobalSetup]
    public void Setup()
    {
        _source = Collection.GetItems(100000);

        ExcelPackage.License.SetNonCommercialPersonal("<Your Name>");
    }

    [IterationSetup]
    public void IterationSetup()
    {
        _stream = new MemoryStream(1024 * 1024 * 40);
    }

    [Benchmark]
    public void Lib()
    {
        var styles = new CollectionExampleStyles();
        var stylesXml = Style.Compile(styles);

        using (var builder = new XlsxDocument(_stream))
        {
            using (var styleXmlWriter = builder.CreateStyleEntry())
            {
                using var xmlWriter = System.Xml.XmlWriter.Create(styleXmlWriter, new System.Xml.XmlWriterSettings { CloseOutput = true });
                stylesXml.WriteTo(xmlWriter);
            }

            builder.CollectionExampleSheet(_source, styles);
        }
    }


    [Benchmark]
    public void EpPlus()
    {
        using var pck = new ExcelPackage(_stream);
        var ws = pck.Workbook.Worksheets.Add(typeof(CollectionItem).Name);
        ws.Cells["A1"].LoadFromCollection(_source, true);

        ws.Column(4).Style.Numberformat.Format = "yyyy-mm-dd";

        pck.Save();
    }
}