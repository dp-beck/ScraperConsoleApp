using System.ComponentModel.DataAnnotations;
using CsvHelper;
using CsvHelper.Configuration;
using HtmlAgilityPack;
using ScraperConsoleApp;
// TO DO:
// 3. Implement outputting to a DOCX file

// Get the URL of the target page

string userInput = "";
var web = new HtmlWeb();
HtmlDocument document = new HtmlDocument();
Company company = new Company();

do
{
    System.Console.WriteLine("Enter the URL of the Kentucky company profile page. Type 'exit' to quit.");
    userInput = Console.ReadLine()!;

    if (userInput.ToLower() == "exit")
    {
        System.Console.WriteLine("Exiting the program.");
        break;
    }

    try
    {
        // downloading to the target page
        // and parsing its HTML content
        document = web.Load(userInput);
    }
    catch (System.UriFormatException)
    {
        System.Console.WriteLine("Invalid URL format. Please try again.");
        userInput = "";
    }

} while (userInput == "");

if (userInput != "exit")
{
    var nodes = document.DocumentNode.SelectNodes("//*[@id='MainContent_pInfo']/div/div/div[position() > 0 and position() < 16]");

    foreach (var node in nodes)
    {
        // Extracting the property name and value
        var propertyName = HtmlEntity.DeEntitize(node.SelectSingleNode("div[1]")?.InnerText.Trim());
        var propertyValue = HtmlEntity.DeEntitize(node.SelectSingleNode("div[2]")?.InnerText.Trim());

        switch (propertyName)
        {
            case "Organization Number":
                company.OrganizationNumber = propertyValue;
                break;
            case "Name":
                company.Name = propertyValue;
                break;
            case "Profit or Non-Profit":
                company.ProfitOrNonprofit = propertyValue;
                break;
            case "Company Type":
                company.CompanyType = propertyValue;
                break;
            case "Industry":
                company.Industry = propertyValue;
                break;
            case "Number of Employees":
                company.NumberOfEmployees = propertyValue;
                break;
            case "Primary County":
                company.PrimaryCounty = propertyValue;
                break;
            case "Status":
                company.Status = propertyValue;
                break;
            case "Standing":
                company.Standing = propertyValue;
                break;
            case "State":
                company.State = propertyValue;
                break;
            case "File Date":
                company.FileDate = propertyValue;
                break;
            case "Authority Date":
                company.AuthorityDate = propertyValue;
                break;
            case "Last Annual Report":
                company.LastAnnualReport = propertyValue;
                break;
            case "Principal Office":
                company.PrincipalOffice = propertyValue;
                break;
            case "Registered Agent":
                company.RegisteredAgent = propertyValue;
                break;
            case "Country":
                company.Country = propertyValue;
                break;
            case "Managed By":
                company.ManagedBy = propertyValue;
                break;
        }
    }

    // Print out the company object to the console

    System.Console.WriteLine("Organization Number: " + company.OrganizationNumber);
    System.Console.WriteLine("Name: " + company.Name);
    System.Console.WriteLine("Profit or Nonprofit: " + company.ProfitOrNonprofit);
    System.Console.WriteLine("Company Type: " + company.CompanyType);
    System.Console.WriteLine("Industry: " + company.Industry);
    System.Console.WriteLine("Number of Employees: " + company.NumberOfEmployees);
    System.Console.WriteLine("Primary County: " + company.PrimaryCounty);
    System.Console.WriteLine("Status: " + company.Status);
    System.Console.WriteLine("Standing: " + company.Standing);
    System.Console.WriteLine("State: " + company.State);
    System.Console.WriteLine("File Date: " + company.FileDate);
    System.Console.WriteLine("Authority Date: " + company.AuthorityDate);
    System.Console.WriteLine("Last Annual Report: " + company.LastAnnualReport);
    System.Console.WriteLine("Principal Office: " + company.PrincipalOffice);
    System.Console.WriteLine("Registered Agent: " + company.RegisteredAgent);
    // Add more fields as needed

}
