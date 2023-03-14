using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph.Models;

namespace MAUIwithMSGRaph.Models
{
	public sealed class SearchResultItem
	{
		private string _searchTerm;

		public EntityType EntityType { get; set; }
		public string Summary { get; set; }
		public string SummaryHtml { 
			get
			{
				var boldedText = Summary.Replace(_searchTerm.ToLower(), $"<strong>{_searchTerm}</strong>");
				return $"{boldedText}";
				//return $"<![CDATA[{boldedText}]]>";
			}

		}
		public string EntityTypeStr => EntityType.ToString();
		public SearchResultItem(string searchTerm)
		{
			_searchTerm = searchTerm;
		}
	}
}
