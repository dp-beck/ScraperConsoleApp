using CsvHelper;
using CsvHelper.Configuration;
using HtmlAgilityPack;
using ScraperConsoleApp;

// the URL of the target Wikipedia page
string url = "https://en.wikipedia.org/wiki/List_of_Marvel_Cinematic_Universe_films";

var web = new HtmlWeb();
// downloading to the target page
// and parsing its HTML content
var document = web.Load(url);

var nodes = document
    .DocumentNode
    .SelectNodes("//*[@id='mw-content-text']/div[1]/table[position()>1 and position()<10]/tbody/tr[position()>1]");

// initializing the list of objects that will
// store the scraped data
List<Movie> movies = new List<Movie>();

// looping over the nodes
// and extract data from them
foreach (var node in nodes)
{
    // add a new Episode instance to
    // to the list of scraped data
    movies.Add(new Movie()
    {
        Title = HtmlEntity.DeEntitize(node.SelectSingleNode("th/i/a")?.InnerText),
        Released = HtmlEntity.DeEntitize(node.SelectSingleNode("td[2]")?.InnerText),
        Directors = HtmlEntity.DeEntitize(node.SelectSingleNode("td[3]")?.InnerText),
        Producers = HtmlEntity.DeEntitize(node.SelectSingleNode("td[4]")?.InnerText)
    });
}

CsvConfiguration csvHelper = new CsvConfiguration()
{
    CultureInfo = System.Globalization.CultureInfo.InvariantCulture,
    Delimiter = ",",
    HasHeaderRecord = false
};

using (var writer = new StreamWriter("output.csv"))
using (var csv = new CsvWriter(writer, csvHelper))
{
    // writing the records
    csv.WriteRecords(movies);
}


       