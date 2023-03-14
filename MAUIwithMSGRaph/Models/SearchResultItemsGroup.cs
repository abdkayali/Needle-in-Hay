namespace MAUIwithMSGRaph.Models
{
    public sealed class SearchResultItemsGroup : List<SearchResultItem>
    {
        public string Name { get; set; }
        public string HeaderText => $"{Name}: {Count}";
        public SearchResultItemsGroup(string name, List<SearchResultItem> items) : base(items)
        {
            Name = name;
        }
    }
}
