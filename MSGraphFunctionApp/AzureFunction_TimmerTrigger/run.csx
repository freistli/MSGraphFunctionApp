/*
# Disclaimer
# Sample Code is provided for the purpose of illustration only and is not intended
# to be used in a production environment. THIS SAMPLE CODE AND ANY RELATED INFORMATION
# ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. We grant You a nonexclusive, royalty-free right to use and
# modify the Sample Code and to reproduce and distribute the object code form of the
# Sample Code, provided that. You agree: (i) to not use Our name, logo, or trademarks 
# to market Your software product in which the Sample Code is embedded; (ii) to include
# a valid copyright notice on Your software product in which the Sample Code is embedded;
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against
# any claims or lawsuits, including attorneys’ fees, that arise or result from the use or
# distribution of the Sample Code
# Author: freistli@microsoft.com
# 06/12/2022
*/

#r "Newtonsoft.Json"


using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Collections.Generic;
using System.Text;
using System.Linq;


public class GroupActivities
{
    [JsonProperty("@odata.context")]
    public string odatacontext { get; set; }

    [JsonProperty("@odata.nextLink")]
    public string odatanextLink { get; set; }

    [JsonProperty("value")]
    public GroupActivity[] groupActivities { get; set; }
}

public class GroupActivity
{
    public string odatatype { get; set; }
    public string groupId { get; set; }
    public string reportRefreshDate { get; set; }
    public string groupDisplayName { get; set; }
    public bool isDeleted { get; set; }
    public string ownerPrincipalName { get; set; }
    public string lastActivityDate { get; set; }
    public string groupType { get; set; }
    public int memberCount { get; set; }
    public int externalMemberCount { get; set; }
    public object exchangeReceivedEmailCount { get; set; }
    public object sharePointActiveFileCount { get; set; }
    public object yammerPostedMessageCount { get; set; }
    public object yammerReadMessageCount { get; set; }
    public object yammerLikedMessageCount { get; set; }
    public int exchangeMailboxTotalItemCount { get; set; }
    public int exchangeMailboxStorageUsedInBytes { get; set; }
    public int sharePointTotalFileCount { get; set; }
    public int sharePointSiteStorageUsedInBytes { get; set; }
    public string reportPeriod { get; set; }
}

public class Groups
{
    [JsonProperty("@odata.context")]
    public string odatacontext { get; set; }
    [JsonProperty("@odata.nextLink")]
    public string odatanextLink { get; set; }
    [JsonProperty("value")]
    public Group[] groups { get; set; }
}

public class Group
{
    public string id { get; set; }
    public object deletedDateTime { get; set; }
    public object classification { get; set; }
    public DateTime createdDateTime { get; set; }
    public string[] creationOptions { get; set; }
    public string description { get; set; }
    public string displayName { get; set; }
    public object expirationDateTime { get; set; }
    public string[] groupTypes { get; set; }
    public object isAssignableToRole { get; set; }
    public string mail { get; set; }
    public bool mailEnabled { get; set; }
    public string mailNickname { get; set; }
    public object membershipRule { get; set; }
    public object membershipRuleProcessingState { get; set; }
    public object onPremisesDomainName { get; set; }
    public object onPremisesLastSyncDateTime { get; set; }
    public object onPremisesNetBiosName { get; set; }
    public object onPremisesSamAccountName { get; set; }
    public object onPremisesSecurityIdentifier { get; set; }
    public object onPremisesSyncEnabled { get; set; }
    public object preferredDataLocation { get; set; }
    public object preferredLanguage { get; set; }
    public string[] proxyAddresses { get; set; }
    public DateTime renewedDateTime { get; set; }
    public string[] resourceBehaviorOptions { get; set; }
    public object[] resourceProvisioningOptions { get; set; }
    public bool securityEnabled { get; set; }
    public string securityIdentifier { get; set; }
    public object theme { get; set; }
    public string visibility { get; set; }
    public object[] onPremisesProvisioningErrors { get; set; }
}

public class Owners
{
    [JsonProperty("@odata.context")]
    public string odatacontext { get; set; }
    [JsonProperty("value")]
    public User[] users { get; set; }
}

public class User
{
    public string odatatype { get; set; }
    public string id { get; set; }
    public string[] businessPhones { get; set; }
    public string displayName { get; set; }
    public string givenName { get; set; }
    public string jobTitle { get; set; }
    public string mail { get; set; }
    public string mobilePhone { get; set; }
    public string officeLocation { get; set; }
    public string preferredLanguage { get; set; }
    public string surname { get; set; }
    public string userPrincipalName { get; set; }
}

public class Report
{
    [JsonProperty("@odata.nextLink")]
    public string odatanextLink { get; set; }
    [JsonProperty("value")]
    public SPSite[] spSites { get; set; }
}

public class SPSite
{
    public string reportRefreshDate { get; set; }
    public string siteId { get; set; }
    public string siteUrl { get; set; }
    public string ownerDisplayName { get; set; }
    public string ownerPrincipalName { get; set; }
    public bool isDeleted { get; set; }
    public string lastActivityDate { get; set; }
    public string siteSensitivityLabelId { get; set; }
    public bool externalSharing { get; set; }
    public string unmanagedDevicePolicy { get; set; }
    public string geolocation { get; set; }
    public int fileCount { get; set; }
    public int activeFileCount { get; set; }
    public int pageViewCount { get; set; }
    public int visitedPageCount { get; set; }
    public int storageUsedInBytes { get; set; }
    public long storageAllocatedInBytes { get; set; }
    public int anonymousLinkCount { get; set; }
    public int companyLinkCount { get; set; }
    public int secureLinkForGuestCount { get; set; }
    public int secureLinkForMemberCount { get; set; }
    public string rootWebTemplate { get; set; }
    public string reportPeriod { get; set; }
}

class FunAppcs
{
    static readonly string clientID = Environment.GetEnvironmentVariable("ClientID");
    static readonly string clientSecret = Environment.GetEnvironmentVariable("ClientSecret");
    static readonly string tenantID = Environment.GetEnvironmentVariable("TenantID");
    static readonly string logicApp = Environment.GetEnvironmentVariable("LogicAppEndPoint");
    static readonly string TeamsQuerylogicApp = Environment.GetEnvironmentVariable("TeamsQueryLogicAppEndPoint");
    /// <summary>
    /// Get MS Graph Access Token as Application, use Client Secret OAuth Flow.
    /// Replace variables for client id, client secret, tenant id.
    /// </summary>
    /// <returns></returns>
    static async Task<string> GetAccessToken()
    {
        // request for access token.
        var parameters = new Dictionary<string, string>();
        parameters.Add("client_id", clientID);
        parameters.Add("client_secret", clientSecret);
        parameters.Add("scope", "https://graph.microsoft.com/.default");
        parameters.Add("grant_type", "client_credentials");

        var client = new HttpClient();
        client.BaseAddress = new Uri("https://login.microsoftonline.com");
        var request = new HttpRequestMessage(HttpMethod.Post, $"{tenantID}/oauth2/v2.0/token");

        request.Content = new FormUrlEncodedContent(parameters);
        var response = await client.SendAsync(request);
        var responseString = await response.Content.ReadAsStringAsync();
        dynamic data = JsonConvert.DeserializeObject(responseString);

        return data.access_token;
    }

    /// <summary>
    /// Post SharePoint report page to Logic App
    /// </summary>
    /// <param name="url"></param>
    /// <param name="report"></param>
    /// <returns></returns>
    static async Task<string> InvokeSharePointQueryLogicAppPostAsync(string url, Report report)
    {
        var reportString = JsonConvert.SerializeObject(report);
        using (var client = new HttpClient())
        {
            var content = new StringContent(reportString, Encoding.UTF8, "application/json");
            HttpResponseMessage result = await client.PostAsync(new Uri(url), content);
            return result.StatusCode.ToString();
        }
    }

    static async Task<string> InvokeTeamsQueryLogicAppPostAsync(string url, GroupActivities report)
    {
        var reportString = JsonConvert.SerializeObject(report);
        using (var client = new HttpClient())
        {
            var content = new StringContent(reportString, Encoding.UTF8, "application/json");
            HttpResponseMessage result = await client.PostAsync(new Uri(url), content);
            return result.StatusCode.ToString();
        }
    }

    /// <summary>
    /// Call Graph API with Get using HTTPClient
    /// </summary>
    /// <param name="url"></param>
    /// <param name="token"></param>
    /// <returns></returns>
    static async Task<string> GetGraphAsync(string url, string token)
    {
        using (var client = new HttpClient())
        {
            client.BaseAddress = new Uri(url);

            var contentType = new MediaTypeWithQualityHeaderValue("application/json");

            client.DefaultRequestHeaders.Accept.Add(contentType);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var response = await client.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                var stringData = await response.Content.ReadAsStringAsync();
                return stringData;
            }
            else
            {
                return $"Failed with {response.StatusCode}";
            }
        }
    }
    /// <summary>
    /// Query getSharePointSiteUsageDetail, return site usage report in JSON string. Support pages
    /// by checking odata.nextLink
    /// </summary>
    /// <param name="period">can fill D7, D90, etc</param>
    /// <param name="token">Access Token</param>
    /// <param name="nextLink">If uses Top and have more pages by following nextLink, then use this parameter</param>
    /// <returns></returns>
    static async Task<Report> SharePointUsageReportQuery(string period, string token, string top, string nextLink = null)
    {
        var url = !string.IsNullOrEmpty(nextLink) ?
            nextLink :
            $"https://graph.microsoft.com/beta/reports/getSharePointSiteUsageDetail(period='{period}')?$format=application/json&$top={top}";

        return JsonConvert.DeserializeObject<Report>(await GetGraphAsync(url, token));
    }
    /// <summary>
    /// Query getOffice365GroupsActivityDetail, return site usage report in JSON string. Support pages
    /// by checking odata.nextLink
    /// </summary>
    /// <param name="period"></param>
    /// <param name="token"></param>
    /// <param name="top"></param>
    /// <param name="nextLink"></param>
    /// <returns></returns>
    static async Task<GroupActivities> TeamsGroupUsageReportQuery(string period, string token, string top, string nextLink = null)
    {
        var url = !string.IsNullOrEmpty(nextLink) ?
            nextLink :
            $"https://graph.microsoft.com/beta/reports/getOffice365GroupsActivityDetail(period='{period}')?$format=application/json&$top={top}";

        return JsonConvert.DeserializeObject<GroupActivities>(await GetGraphAsync(url, token));
    }

    static async Task<List<Group>> GetGroupsInMemory(string token, string top)
    {
        var url = $"https://graph.microsoft.com/v1.0/groups?$select=id,mail&$top={top}";

        Groups groups = JsonConvert.DeserializeObject<Groups>(await GetGraphAsync(url, token));
        List<Group> groupList = new List<Group>(groups.groups);

        while (!string.IsNullOrEmpty(groups.odatanextLink))
        {
            log.LogInformation(groups.odatanextLink);
            groups = JsonConvert.DeserializeObject<Groups>(await GetGraphAsync(groups.odatanextLink, token));
            groupList.AddRange(new List<Group>(groups.groups));
        }

        return groupList;
    }
    /// <summary>
    /// Query Groups by group mail
    /// </summary>
    /// <returns></returns>
    static async Task<Groups> GetGroupsByMail(string groupMail, string token)
    {
        var url = $"https://graph.microsoft.com/v1.0/groups?$filter=mail eq '{groupMail}'&$select=id";

        return JsonConvert.DeserializeObject<Groups>(await GetGraphAsync(url, token));
    }

    static async Task<Owners> GetGroupOwners(string groupId, string token)
    {
        var url = $"https://graph.microsoft.com/v1.0/groups/{groupId}/owners?&$select=userPrincipalName";

        return JsonConvert.DeserializeObject<Owners>(await GetGraphAsync(url, token));
    }

    /// <summary>
    /// Check site group owners if the site owner is MS Group
    /// </summary>
    /// <param name="spSite"></param>
    /// <param name="token"></param>
    /// <returns></returns>
    static async Task<string> GetSiteOwners(SPSite spSite, string token)
    {
        //This Group type means the site is connected with MS 365 Group and owner email is group email
        if (spSite.rootWebTemplate == "Group")
        {
            Groups groups = await GetGroupsByMail(spSite.ownerPrincipalName, token);
            Owners owners = await GetGroupOwners(groups.groups[0].id, token);
            StringBuilder ownerList = new StringBuilder();
            foreach (var owner in owners.users)
            {
                ownerList.Append($"{owner.userPrincipalName};");
            }
            return string.IsNullOrEmpty(ownerList.ToString()) ? spSite.ownerPrincipalName : ownerList.ToString();
        }
        else
        {
            return spSite.ownerPrincipalName;
        }
    }

    /// <summary>
    /// Get full owners list for GroupActivity Item
    /// </summary>
    /// <param name="groupList"></param>
    /// <param name="groupActivity"></param>
    /// <param name="token"></param>
    /// <returns></returns>
    static async Task<string> GetGroupOwnersQuick(List<Group> groupList, GroupActivity groupActivity, string token)
    {
        Owners owners = await GetGroupOwners(groupActivity.groupId, token);
        StringBuilder ownerList = new StringBuilder();
        foreach (var owner in owners.users)
        {
            ownerList.Append($"{owner.userPrincipalName};");
        }
        return string.IsNullOrEmpty(ownerList.ToString()) ? groupActivity.ownerPrincipalName : ownerList.ToString();

    }
    /// <summary>
    /// Find matched group in memory groupList
    /// </summary>
    /// <param name="groupList"></param>
    /// <param name="spSite"></param>
    /// <param name="token"></param>
    /// <returns></returns>
    static async Task<string> GetSiteOwnersQuick(List<Group> groupList, SPSite spSite, string token)
    {
        if (spSite.rootWebTemplate == "Group")
        {
            IEnumerable<Group> groupsQuery = groupList.Where(group => group.mail == spSite.ownerPrincipalName);
            Owners owners = await GetGroupOwners(groupsQuery.FirstOrDefault().id, token);
            StringBuilder ownerList = new StringBuilder();
            foreach (var owner in owners.users)
            {
                ownerList.Append($"{owner.userPrincipalName};");
            }
            return string.IsNullOrEmpty(ownerList.ToString()) ? spSite.ownerPrincipalName : ownerList.ToString();
        }
        else
        {
            return spSite.ownerPrincipalName;
        }
    }


    public static async Task TeamsReportRunQuick(string period, string top)
    {
        log.LogInformation($"C# function (Teams Report) executed with parameters: Period {period} Top {top}");

        var watch = new System.Diagnostics.Stopwatch();
        watch.Start();
        int count = 0;
        log.LogInformation($"C# function executed at: {DateTime.Now}");
        var token = await GetAccessToken();

        GroupActivities report = await TeamsGroupUsageReportQuery(period, token, top);
        List<Group> groupList = await GetGroupsInMemory(token, top);

        foreach (var item in report.groupActivities)
        {
            item.ownerPrincipalName = await GetGroupOwnersQuick(groupList, item, token);
            log.LogInformation($"{item.groupDisplayName} - {item.ownerPrincipalName}");
        }

        var result = await InvokeTeamsQueryLogicAppPostAsync(TeamsQuerylogicApp, report);
        log.LogInformation(result);

        count += report.groupActivities.Length;

        while (!string.IsNullOrEmpty(report.odatanextLink))
        {
            report = await TeamsGroupUsageReportQuery(period, token, top, report.odatanextLink);
            log.LogInformation(report.odatanextLink);

            foreach (var item in report.groupActivities)
            {
                item.ownerPrincipalName = await GetGroupOwnersQuick(groupList, item, token);
                log.LogInformation($"{item.groupDisplayName} - {item.ownerPrincipalName}");
            }

            result = await InvokeTeamsQueryLogicAppPostAsync(TeamsQuerylogicApp, report);
            log.LogInformation(result);

            count += report.groupActivities.Length;
        }

        log.LogInformation($"C# function (Teams Report) finished at: {DateTime.Now}");
        log.LogInformation($"{count} items are processed.");
        watch.Stop();
        log.LogInformation($"Execution Time: {watch.ElapsedMilliseconds / 1000} s");
    }

    /// <summary>
    /// Load groups in memory to reduce MS Graph API calls
    /// </summary>
    public static async Task SharePointQueryRunQuick(string period, string top)
    {
        log.LogInformation($"C# function (SharePoint Report) executed with parameters: Period {period} Top {top}");

        var watch = new System.Diagnostics.Stopwatch();
        watch.Start();
        int count = 0;
        log.LogInformation($"C# function executed at: {DateTime.Now}");
        var token = await GetAccessToken();

        Report report = await SharePointUsageReportQuery(period, token, top);
        List<Group> groupList = await GetGroupsInMemory(token, top);

        foreach (var item in report.spSites)
        {
            item.ownerPrincipalName = await GetSiteOwnersQuick(groupList, item, token);
            log.LogInformation($"{item.siteUrl} {item.ownerDisplayName} {item.ownerPrincipalName}");
        }

        var result = await InvokeSharePointQueryLogicAppPostAsync(logicApp,
            report);
        log.LogInformation(result);

        count += report.spSites.Length;

        while (!string.IsNullOrEmpty(report.odatanextLink))
        {
            report = await SharePointUsageReportQuery(period, token, top, report.odatanextLink);
            log.LogInformation(report.odatanextLink);

            foreach (var item in report.spSites)
            {
                item.ownerPrincipalName = await GetSiteOwners(item, token);
                log.LogInformation($"{item.siteUrl} {item.ownerDisplayName} {item.ownerPrincipalName}");
            }

            result = await InvokeSharePointQueryLogicAppPostAsync(logicApp, report);
            log.LogInformation(result);

            count += report.spSites.Length;
        }

        log.LogInformation($"C# function (SharePoint Report) finished at: {DateTime.Now}");
        log.LogInformation($"{count} items are processed.");
        watch.Stop();
        log.LogInformation($"Execution Time: {watch.ElapsedMilliseconds / 1000} s");
    }

    /// <summary>
    /// Didn't store MS Groups in memory. More hits on MS Graph APIs
    /// </summary>
    /// <param name="period"></param>
    /// <param name="top"></param>
    /// <returns></returns>
    public static async Task SharePointQueryRun(string period, string top)
    {
        log.LogInformation($"C# function executed with parameters: Period {period} Top {top}");

        var watch = new System.Diagnostics.Stopwatch();
        watch.Start();
        int count = 0;
        log.LogInformation($"C# function executed at: {DateTime.Now}");
        var token = await GetAccessToken();

        Report report = await SharePointUsageReportQuery(period, token, top);

        foreach (var item in report.spSites)
        {
            item.ownerPrincipalName = await GetSiteOwners(item, token);
            log.LogInformation($"{item.siteUrl} {item.ownerDisplayName} {item.ownerPrincipalName}");
        }

        var result = await InvokeSharePointQueryLogicAppPostAsync(logicApp, report);
        log.LogInformation(result);

        count += report.spSites.Length;

        while (!string.IsNullOrEmpty(report.odatanextLink))
        {
            report = await SharePointUsageReportQuery(period, token, top, report.odatanextLink);
            log.LogInformation(report.odatanextLink);

            foreach (var item in report.spSites)
            {
                item.ownerPrincipalName = await GetSiteOwners(item, token);
                log.LogInformation($"{item.siteUrl} {item.ownerDisplayName} {item.ownerPrincipalName}");
            }

            result = await InvokeSharePointQueryLogicAppPostAsync(logicApp, report);
            log.LogInformation(result);

            count += report.spSites.Length;
        }

        log.LogInformation($"C# function finished at: {DateTime.Now}");
        log.LogInformation($"{count} items are processed.");
        watch.Stop();
        log.LogInformation($"Execution Time: {watch.ElapsedMilliseconds / 1000} s");
    }

}

static ILogger log;

public static async void Run(TimerInfo myTimer, ILogger ilog)
{
    log = ilog;
    log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
    //await FunAppcs.TeamsReportRunQuick("D7", "100");
    await FunAppcs.SharePointQueryRunQuick("D7", "100");

}
