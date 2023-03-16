using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MAUIwithMSGRaph.Models;
using Microsoft.Graph.Models;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows.Input;
using CommunityToolkit.Maui.Alerts;
using CommunityToolkit.Maui.Core;
using static System.Net.Mime.MediaTypeNames;
using Newtonsoft.Json;

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
        private bool includePeople;

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

		public bool ShowErrorMessageLabel { get; set; } 

        public ICommand SearchForTheNeedleCommand { get; }

		//public ICommand TapCommand { get;  }

		public ObservableCollection<SearchResultItemsGroup> SearchResultItemsGroup { get; }

        public MainViewModel()
        {
			SearchResultItemsGroup= new ObservableCollection<SearchResultItemsGroup>();
            SearchForTheNeedleCommand = new AsyncRelayCommand(SearchForTheNeedleAsync);
			//TapCommand = new AsyncRelayCommand<Uri>(GoToWebsiteAsync);
			ShowErrorMessageLabel = DeviceInfo.Platform != DevicePlatform.Android;
		}

		//private async Task GoToWebsiteAsync(Uri uri)
		//{
		//	await Launcher.OpenAsync(uri);
		//}

		private async Task SearchForTheNeedleAsync()
		{
			// validation here
			if (string.IsNullOrEmpty(SearchTerms))
			{
				await ShowErrorAsync("Please enter the term to search for!");

				return;
			}

			if (!(IncludeChatMessage || IncludePeople || IncludeMessage || IncludeEvent || IncludeFilesSites))
			{
				await ShowErrorAsync("Please check at least one source for search!");

				return;
			}

			SearchResultItemsGroup.Clear();

			if (_graphService == null)
			{
				_graphService = new GraphService();
			}

            if (IncludeChatMessage) AddToSearchResultRows(await GetResultsForEntityTypeAsync(EntityType.ChatMessage));
			if (IncludePeople) AddToSearchResultRows(await GetResultsFromPeopleAsync());
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

		private async Task ShowErrorAsync(string message)
		{
			if (DeviceInfo.Platform != DevicePlatform.Android)
			{
				ErrorMessage = message;
				await Task.Delay(2 * 1000);
				ErrorMessage = null;

				return;
			}

			// use a toast
			var cancellationTokenSource = new CancellationTokenSource();
			ToastDuration duration = ToastDuration.Short;
			double fontSize = 14;

			var toast = Toast.Make(message, duration, fontSize);

			await toast.Show(cancellationTokenSource.Token);
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
		private void AddToSearchResultRows((string, List<SearchResultItem>) tuple)
		{
			SearchResultItemsGroup.Add(new SearchResultItemsGroup(
				tuple.Item1, 
				tuple.Item2));
		}

		private async Task<(string, List<SearchResultItem>)> GetResultsForEntityTypeAsync(EntityType entityType)
        {
			const int PageSize = 10;
			int fromParameter = 0;
			var result = new List<SearchResultItem>();
			
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
							foreach (var hit in hc.Hits)
							{
								string uri = null;
								if (entityType == EntityType.ChatMessage)
								{
									// use Channel Info endpoint to grab WebUrl for the message in the chat
									var channelInfoRaw = hit.Resource?.AdditionalData["channelIdentity"].ToString(); // {"channelId":"19:fa81aa6db3a84de4b42eb1832e8ad9a8@thread.tacv2","teamId":"4b72f057-4133-42e6-a529-92b6e74211aa"}
									var sampleChannelIdentityObject = new
									{
										channelId = string.Empty,
										teamId = string.Empty
									};
									var channelInfo = JsonConvert.DeserializeAnonymousType(channelInfoRaw, sampleChannelIdentityObject);
									uri = await GetWebUrlAsync(channelInfo.teamId, channelInfo.channelId);
								}
								result.Add(new SearchResultItem(SearchTerms)
								{
									Summary = hit.Summary,
									ResourceId = hit.Resource?.Id,
									Uri = uri,
									Dictionary = hit.Resource?.AdditionalData.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.ToString())
								});
							}
							//hc.Hits.ForEach(async h =>  
							//{
								
							//});
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

			return (GetEntityTypeString(entityType), result); 
		}

		private async Task<string> GetWebUrlAsync(string teamId, string channelId)
		{
			//var endpoint = $"https://graph.microsoft.com/v1.0/teams/{teamId}/channels/{channelId}";
			if (_graphService == null)
			{
				_graphService = new GraphService();
			}
			try
			{
				var channelItemRequestBuilder = _graphService.Client.Teams[teamId].Channels[channelId];
				//var requestInfo = channelItemRequestBuilder.ToGetRequestInformation();
				var result = await channelItemRequestBuilder.GetAsync();

				if (result != null)
				{
					return result.WebUrl;
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.Message);
			}
			

			return string.Empty;
		}

		private async Task<(string, List<SearchResultItem>)> GetResultsFromPeopleAsync()
		{
			var result = new List<SearchResultItem>();

			try
			{
				var response = await _graphService.Client.Me.People.GetAsync((requestConfig) =>
				{
					requestConfig.QueryParameters.Search = SearchTerms;
				});
				if (response != null && response.Value.Any())
				{
					result.AddRange(response.Value.Select(p => new SearchResultItem(SearchTerms)
					{ 
						Summary = p.DisplayName, 
						Uri = null,
						Dictionary = p.AdditionalData.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.ToString())
					}));
				}
			}
			catch (Exception ex)
			{
				await ShowErrorAsync("Error Occurred!");
				Debug.WriteLine(ex.Message);
			}

			return ("People", result);
		}

	}
}
