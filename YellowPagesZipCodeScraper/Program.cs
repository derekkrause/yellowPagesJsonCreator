using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace YellowPagesZipCodeScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter search term and press 'Enter': ");
            string searchTerm = Uri.EscapeDataString(Console.ReadLine());

            Console.WriteLine("Enter Zip Code and press 'Enter': ");
            string zipCode = Console.ReadLine();
            List<BusinessEntry> businessResults = new List<BusinessEntry>();

            int pageNumber = 1;
            bool _continue = true;
            string url;
            IHtmlDocument document;

            while (_continue)
            {
                getURL();
                getHtmlDocument();
                parseResults();
                processPagination();
            }

            Console.WriteLine();
            Console.WriteLine("Would you like to save the results to file? (Y/N)?");
            if (Console.ReadLine().ToLower() == "y" || Console.ReadLine().ToLower() == "yes")
            {
                saveToFile();
            }
            Console.WriteLine("Press 'Enter' to end.");
            Console.Read();

            void getURL()
            {
                url = Uri.EscapeUriString("https://www.yellowpages.com/search?search_terms=" + searchTerm + "&geo_location_terms=" + zipCode);
                if (pageNumber > 1)
                {
                    url = Uri.EscapeUriString(url += "&page=" + pageNumber);
                }
            }

            void getHtmlDocument()
            {
                var webClient = new WebClient();
                var html = webClient.DownloadString(url);
                var parser = new HtmlParser();
                document = parser.ParseDocument(html);
            }

            void parseResults()
            {
                var resultElements = document.QuerySelectorAll("div .result");

                foreach (var resultElement in resultElements)
                {
                    var business = new BusinessEntry()
                    {
                        Name = resultElement.QuerySelector(".business-name > span").TextContent,
                        Phone = resultElement.QuerySelector(".phones.phone.primary")?.TextContent ?? "not found",
                        StreetAddress = resultElement.QuerySelector(".street-address")?.TextContent ?? "not found",
                        CityStateZip = resultElement.QuerySelector(".locality")?.TextContent ?? "not found",
                        Description = resultElement.QuerySelector(".snippet > p > span")?.TextContent ?? "not found",
                        Website = resultElement.QuerySelector(".links > a")?.GetAttribute("href") ?? "not found"
                    };

                    businessResults.Add(business);
                    writeResultToConsole(business);
                }
            }

            void processPagination()
            {
                var pagination = document.QuerySelector("div .pagination");
                if (pagination.InnerHtml != "")
                {
                    var pageOptions = pagination.QuerySelectorAll("ul > li").ToList();
                    var lastPaginationButton = pageOptions[pageOptions.Count - 1].TextContent;
                    int lastPage;
                    if (lastPaginationButton.ToLower() == "next")
                    {
                        lastPage = int.Parse(pageOptions[pageOptions.Count - 2].TextContent);
                    }
                    else
                    {
                        lastPage = int.Parse(pageOptions[pageOptions.Count - 1].TextContent);
                    }

                    if (lastPage > pageNumber)
                    {
                        pageNumber += 1;
                    }
                    else
                    {
                        _continue = false;
                    }
                }
            }

            void saveToFile()
            {
                Console.WriteLine();
                Console.WriteLine("Choose how you'd like to save your file... Type A or B and press 'Enter'.");
                Console.WriteLine("A - .json");
                Console.WriteLine("B - .csv");
                char selection = Console.ReadLine().ToLower()[0];

                switch (selection)
                {
                    case 'a':
                        saveAsType("json");
                        break;
                    case 'b':
                        saveAsType("csv");
                        break;
                    default:
                        Console.WriteLine("Invalid Response. Please select either A or B and press 'Enter'.");
                        Console.Read();
                        break;
                }
            }

            void saveAsType(string fileType)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                string selectedPath = "";
                Thread t = new Thread((ThreadStart)(() =>
                {
                    saveFileDialog.Filter = "CSV file(*.csv) | *.csv";
                    saveFileDialog.Title = "Save Results to " + fileType;
                    if (fileType == "json")
                    {
                        saveFileDialog.Filter = "JSON Files (*.json)|*.json";
                    }

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        selectedPath = saveFileDialog.FileName;
                    };

                    if (selectedPath != "")
                    {
                        if (fileType == "json")
                        {
                            string jsonResults = JsonConvert.SerializeObject(businessResults);
                            File.WriteAllText(selectedPath, jsonResults);
                        }

                        if (fileType == "csv")
                        {
                            string csvResults = GenerateCSVReport(businessResults);
                            File.WriteAllText(selectedPath, csvResults);
                        }

                    }
                }));

                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                Console.WriteLine("File saved to " + $"{selectedPath}");
            }

            void writeResultToConsole(BusinessEntry business)
            {
                Console.WriteLine("Name: " + $"{business.Name}");
                Console.WriteLine("Phone: " + $"{business.Phone}");
                Console.WriteLine("Street: " + $"{business.StreetAddress}");
                Console.WriteLine("City: " + $"{business.CityStateZip}");
                Console.WriteLine("Description: " + $"{business.Description}");
                Console.WriteLine("Website: " + $"{business.Website}");
                Console.WriteLine();
                Console.WriteLine("=======================================");
                Console.WriteLine();
            }
        }
        public static string GenerateCSVReport(List<BusinessEntry> items)
        {
            var output = "";
            var delimiter = ",";
            var properties = typeof(BusinessEntry).GetProperties()
             .Where(n => n.PropertyType == typeof(string));

            using (var sw = new StringWriter())
            {
                var header = properties
                    .Select(n => n.Name)
                    .Aggregate((a, b) => a + delimiter + b);

                sw.WriteLine(header);

                foreach (var item in items)
                {
                    var row = properties
                    .Select(n => n.GetValue(item, null))
                    .Select(n => n == null ? "null" : n.ToString())
                    .Aggregate((a, b) => a + delimiter + b);
                    sw.WriteLine(row);
                }
                output = sw.ToString();
            }
            return output;
        }
    }

    public class BusinessEntry
    {
        public string Name { get; set; }
        public string Phone { get; set; }
        public string StreetAddress { get; set; }
        public string CityStateZip { get; set; }
        public string Description { get; set; }
        public string Website { get; set; }
    }
}
