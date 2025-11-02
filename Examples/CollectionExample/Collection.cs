namespace Examples.CollectionExample;

public static class Collection
{
    public static List<CollectionItem> GetItems(int total)
    {
        var source = new List<CollectionItem>();
        var rnd = new Random();

        for (var i = 0; i < total; i++)
        {
            source.Add(new CollectionItem
            {
                Id = Guid.NewGuid().ToString(),
                Name = $"John Doue {i}",
                NickName = "<Big> \"Boy\" & 'co'",
                Salary = rnd.Next(3000, 15000),
                BirthDate = DateTime.Now.AddYears(-rnd.Next(20, 50)),
                HasKids = i % 2 == 0
            });
        }

        return source;
    }
    public static async IAsyncEnumerable<CollectionItem> GetAsyncItems(this IEnumerable<CollectionItem> source)
    {
        foreach (var item in source)
        {
            yield return item;
            await Task.Yield();
        }
    }
}

public class CollectionItem
{
    public string Id { get; set; } = null!;
    public string Name { get; set; } = null!;
    public string NickName { get; set; } = null!;
    public decimal Salary { get; set; }
    public DateTime BirthDate { get; set; }
    public bool HasKids { get; set; }
}