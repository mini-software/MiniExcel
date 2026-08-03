namespace MiniExcelLib.OpenXml.Tests.Templates;

internal class TestIEnumerableTypePoco
{
    public string? @string { get; set; }
    public int? @int { get; set; }
    public decimal? @decimal { get; set; }
    public double? @double { get; set; }
    public DateTime? datetime { get; set; }
    public bool? @bool { get; set; }
    public Guid? Guid { get; set; }
}

internal class Employee
{
    public string? name { get; set; }
    public string? department { get; set; }
}

internal record struct Identity(int Type, string Id);

internal record NetValue(DateOnly Date, decimal Value)
{
    internal static List<NetValue> GenerateRandomValues(int fundType, DateOnly startDate)
    {
        var netValues = new List<NetValue>();
        var random = Random.Shared;

        for (int i = 0; i < 30; i++)
        {
            var value = fundType switch
            {
                1 => Math.Round(1.0000m + (decimal)random.NextDouble() * 0.0010m, 4),
                2 => Math.Round(1.2m + (decimal)random.NextDouble() * 0.8m, 4),
                3 => Math.Round(1.05m + (decimal)random.NextDouble() * 0.25m, 4),
                4 => Math.Round(1.1m + (decimal)random.NextDouble() * 0.7m, 4),
                5 => Math.Round(1.5m + (decimal)random.NextDouble() * 1.0m, 4),
                _ => 1.0000m
            };

            netValues.Add(new NetValue(startDate.AddDays(i), value));
        }

        return netValues;
    }
}

internal class Fund
{
    public int Id { get; set; }
    public string? Name { get; set; }
    public Identity Identity { get; set; }
    public DateOnly SetupDate { get; set; }

    public List<NetValue> NetValues { get; set; } = [];
}
