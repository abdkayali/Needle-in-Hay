using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MAUIwithMSGRaph.Models;
using Microsoft.Graph.Models;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows.Input;

namespace MAUIwithMSGRaph.ViewModels
{
	public partial class MainViewModel : ObservableObject
    {
        [ObservableProperty]
        private string helloMessage = "Find the Needle in the Haystack!";

        [ObservableProperty]
        private string subtitleMessage = "Look for any term in the selected entity types";
        
        [ObservableProperty]
        private bool includeChatMessage;

        [ObservableProperty]
        private bool includeMessage;

        [ObservableProperty]
        private bool includeEvent;

        [ObservableProperty]
        private bool includeFilesSites;

        [ObservableProperty]
        private string searchTerms;

		[ObservableProperty]
		private string errorMessage = null;

		private GraphService _graphService;

		public bool ShowErrorMessage { get; set; } 

        public ICommand SearchForTheNeedleCommand { get; }

		public ObservableCollection<SearchResultItemsGroup> SearchResultItemsGroup { get; }

        public MainViewModel()
        {
			SearchResultItemsGroup= new ObservableCollection<SearchResultItemsGroup>();
            SearchForTheNeedleCommand = new AsyncRelayCommand(SearchForTheNeedleAsync);
		}

		private async Task SearchForTheNeedleAsync()
		{
			// validation here
			if (string.IsNullOrEmpty(SearchTerms))
			{
				ErrorMessage = "Please enter the term to search!"; 
				await Task.Delay(2 * 1000);
				ErrorMessage = null;

				return;
			}

			if (!(IncludeChatMessage || IncludeMessage || IncludeEvent || IncludeFilesSites))
			{
				ErrorMessage = "Please check at least one source for search!";
				await Task.Delay(2 * 1000);
				ErrorMessage = null;

				return;
			}

			SearchResultItemsGroup.Clear();

			if (_graphService == null)
			{
				_graphService = new GraphService();
			}

            if (IncludeChatMessage) AddToSearchResultRows(await GetResultsForEntityTypeAsync(EntityType.ChatMessage));
			if (IncludeMessage) AddToSearchResultRows(await GetResultsForEntityTypeAsync(EntityType.Message));
            if (IncludeEvent) AddToSearchResultRows(await GetResultsForEntityTypeAsync(EntityType.Event));
			if (IncludeFilesSites)
			{
				AddToSearchResultRows(await GetResultsForEntityTypeAsync(EntityType.Drive));
				AddToSearchResultRows(await GetResultsForEntityTypeAsync(EntityType.DriveItem));
				AddToSearchResultRows(await GetResultsForEntityTypeAsync(EntityType.List));
				AddToSearchResultRows(await GetResultsForEntityTypeAsync(EntityType.ListItem));
				AddToSearchResultRows(await GetResultsForEntityTypeAsync(EntityType.Site));
			}
		}

		private string GetEntityTypeString(EntityType entityType)
		{
			return entityType switch
			{
				EntityType.DriveItem => "Drive Item",
				EntityType.ListItem => "List Item",
				EntityType.ChatMessage => "Chat Message",
				EntityType.ExternalItem => "External Item",
				_ => entityType.ToString(),
			};
		}
		private void AddToSearchResultRows((EntityType, List<string>) tuple)
		{
			SearchResultItemsGroup.Add(new SearchResultItemsGroup(
				GetEntityTypeString(tuple.Item1), 
				tuple.Item2.Select(s => new SearchResultItem(this.SearchTerms) { Summary = s, EntityType = tuple.Item1 }).ToList()));
		}

		private async Task<(EntityType, List<string>)> GetResultsForEntityTypeAsync(EntityType entityType)
        {
			const int PageSize = 10;
			int fromParameter = 0;
			var result = new List<string>();
			
			try
			{
				while (true)
				{
					var requestBody = new Microsoft.Graph.Search.Query.QueryPostRequestBody
					{
						Requests = new List<SearchRequest>
							{
								new SearchRequest
								{
									EntityTypes = new List<EntityType?> { entityType },
									Query = new SearchQuery
									{
										QueryString = SearchTerms,
									},
									From = fromParameter,
									Size = PageSize
								},
							},
					};

					var queryResponse = await _graphService.Client.Search.Query.PostAsync(requestBody);
					if (queryResponse != null 
						&& queryResponse.Value.Any() 
						&& queryResponse.Value.First().HitsContainers.Any()) 
					{
						var hc = queryResponse.Value.First().HitsContainers.First();
						if (hc != null && hc.Hits != null && hc.Hits.Any())
						{
							result.AddRange(hc.Hits.Select(h => h.Summary));
							fromParameter += hc.Hits.Count;
						}

						if ((bool)!hc.MoreResultsAvailable || hc.Hits == null)
							break;

					}
				}				
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.Message);
			}

			return (entityType, result); 
		}
	}
}
