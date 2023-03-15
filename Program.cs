using Azure.Identity;
using Microsoft.Graph;

var scopes = new[] { "User.Read", "Sites.ReadWrite.All" };
var interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions
{
    ClientId = "75e060cb-a37b-4bc1-8a28-836a1defa6c6"
};
var tokenCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);

var graphClient = new GraphServiceClient(tokenCredential, scopes);
var siteId = "stdntpartners.sharepoint.com,a4455d95-aa77-4cef-ac9c-8be2a570d4df";

var me = await graphClient.Me.GetAsync();

Console.WriteLine($">>>> Hello {me?.DisplayName}! <<<<");

//2.Search for Sharepoint site by name Kelcho
var searchResult = await graphClient.Sites
    .GetAsync(requestConfiguration =>
    {
        requestConfiguration.QueryParameters.Search = "kelcho";
    });

if (searchResult != null && searchResult.Value != null)
{
    foreach (var site in searchResult.Value)
    {
        Console.WriteLine(">>>> Kelcho Searched Site details <<<<");
        Console.WriteLine($"Site Name: {site.DisplayName} \t Created: {site.CreatedDateTime} \t Url: {site.WebUrl} \t Collection: {site.SiteCollection} \t Site-ID : {site.Id}");
    }
}

//3. Get a site by id
var siteByID = await graphClient.Sites[siteId].GetAsync();

if (siteByID != null)
{
    Console.WriteLine(">>>> Searched Site details Via ID <<<<");
    Console.WriteLine($"Site Name: {siteByID.DisplayName} \t Created: {siteByID.CreatedDateTime} \t Url: {siteByID.WebUrl} \t Collection: {siteByID.SiteCollection} \t Site-ID : {siteByID.Id}");
}

//4. enumerate site columns of the root site
var columnsResult = await graphClient.Sites[siteId].Columns.GetAsync();
if (columnsResult != null)
{
    Console.WriteLine(">>>> Enumerate site columns of the root site <<<<");
    foreach (var column in columnsResult.Value)
    {

        Console.WriteLine($"Column Name: {column.Name} \t  Column-ID : {column.Id}");
    }
}


//5. enumerate site columns of the root site
var result = await graphClient.Sites[siteId].Columns.GetAsync();
if (result != null)
{
    Console.WriteLine(">>>> Enumerate site columns of the root site <<<<");
    foreach (var column in result.Value)
    {
        Console.WriteLine($"DisplayName: {column.DisplayName} columnGroup: {column.ColumnGroup}\t description: {column.Description} \t  Column-ID : {column.Id}");
    }
}

