namespace MiniExcelLib.Tests.TemplateOptimization;

public class FileSystemEntry
{
    public string Path { get; set; }
    public SecurityIdentifier Owner { get; set; }
    public FileAttributes FileEntryAttributes { get; set; }
    public Dictionary<SecurityIdentifier, FileSystemRights> Acl { get; set; }
    public bool IsModified { get; set; }
    public string Error { get; set; }
}

public static class DummyFiller
{
    public static IEnumerable<FileSystemEntry> GenerateDummyFileSystemEntries(int n)
    {
        var allRights = Enum.GetValues<FileSystemRights>().ToArray();
        var random = new Random();
        var firstAcl = DummyAccessList(random, allRights);

        yield return new FileSystemEntry
        {
            Path = $@"C:\Users\DummyUser\Documents\",
            Owner = DummySecurityIdentifier(random),
            FileEntryAttributes = FileAttributes.Directory,
            Acl = firstAcl,
            IsModified = (random.Next() & 1) == 1,
            Error = firstAcl.Count == 0 ? "Access Denied" : string.Empty // 10% chance of error
        };

        var tempEnum = Enumerable.Range(1, n).Select(i =>
        {
            var acl = DummyAccessList(random, allRights);
            var randomAttributes = (FileAttributes)((1 << random.Next(12)) | (int)FileAttributes.Normal);
            randomAttributes &= ~FileAttributes.Directory; 
            return new FileSystemEntry
            {
                Path = $@"C:\Users\DummyUser\Documents\File {i++}.txt",
                Owner = DummySecurityIdentifier(random),
                FileEntryAttributes = randomAttributes,
                Acl = acl,
                IsModified = (random.Next() & 1) == 1,
                Error = acl.Count == 0 ? "Access Denied" : string.Empty // 10% chance of error
            };
        });
        foreach (var entry in tempEnum)
        {
            yield return entry;
        }
    }

    private static Dictionary<SecurityIdentifier, FileSystemRights> DummyAccessList(Random random, FileSystemRights[] allRights) =>
        Enumerable.Range(0, random.Next(0, 5))
            .Select(_ => new
            {
                Sid = DummySecurityIdentifier(random),
                AccessType = DummyFileSystemRights(random, allRights)
            })
            .GroupBy(o => o.Sid)
            .ToDictionary(
                g => g.Key,
                g => g.Select(x => x.AccessType).First()
            );

    private static SecurityIdentifier DummySecurityIdentifier(Random random) => new(
        $"S-1-5-21-100000000-100000000-100000000-{random.Next(999990000, 999999999)}");


    private static FileSystemRights DummyFileSystemRights(Random random, FileSystemRights[] allRights)
    {
        var rightsCount = random.Next(1, allRights.Length);
        return (FileSystemRights)random.NextInt64((1L << rightsCount) - 1);
    }
}