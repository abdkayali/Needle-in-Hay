using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Graph.Models;
using System.Globalization;
using System.Windows.Input;

namespace MAUIwithMSGRaph.Models
{
	public sealed class SearchResultItem : ObservableObject
	{
		private string _searchTerm;

		public string EntityType { get; set; }
		public string Summary { get; set; }

		public string SummaryHtml { 
			get
			{
				var boldedText = Summary.Replace(_searchTerm, $"<strong>{_searchTerm}</strong>", true, CultureInfo.InvariantCulture);
				return $"{boldedText}";
			}
		}

		public string ResourceId { get; set; }
		public string Uri { get; set; }
		public ICommand TapCommand => new AsyncRelayCommand<string>(async (url) => await Launcher.OpenAsync(url));

		public bool HasUri => Uri != null;

		public IDictionary<string, string> Dictionary { get; set; }

		public SearchResultItem(string searchTerm)
		{
			_searchTerm = searchTerm;
		}
	}
}
