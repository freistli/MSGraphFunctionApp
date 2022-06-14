# MSGraphFunctionApp

## Description

### Improve owner priciple name result when calling getSharePointSiteUsageDetail, default result for MS Group Owner only contain group email address
The .Net Core 3.1/5/6 project will get SPO site activity report. If the site owner of one report item is MS Group Mail, the code gets the ownership by querying MS Graph Groups, and replace the ownerPrincipalName with owners email list in the report. 
  
And then posts the report to one logic app to continuously operate the records, for example, save to sharepoint list. (the owner email text column should be multiple lines text type, otherwise single text line is 255 characters maximum, cannot accepte long owners email list)

To understand the logic, good start entry is:
https://github.com/freistli/MSGraphFunctionApp/blob/master/MSGraphFunctionApp/FunAppcs.cs#L452

### Improve owner priciple name result when calling getOffice365GroupsActivityDetail， default result only contain one email address
The .Net Core 3.1/5/6 project will get MS 365 group activity report. the code gets the ownership by querying all MS Graph Groups, and replace the ownerPrincipalName with complete owners email list in the report. 
  
And then posts the report to one logic app to continuously operate the records, for example, save to sharepoint list. (the owner email text column should be multiple lines text type, otherwise single text line is 255 characters maximum, cannot accepte long owners email list)

To understand the logic, good start entry is:
https://github.com/freistli/MSGraphFunctionApp/blob/master/MSGraphFunctionApp/FunAppcs.cs#L402

***AzureFunction_TimmerTrigger\run.csx*** use the same logic can work in Azure Function App Timer Trigger as well.

## Prepare: Set 5 environment variables on desktop:

1. ClientID

1. ClientSecret

1. TenantID

1. LogicAppEndPoint   (this is for SharePoint Query logic app)

1. TeamsQueryLogicAppEndPoint  (this is for Teams Query logic app)

<img src="https://user-images.githubusercontent.com/8623897/173288472-6fdb06c0-0169-4160-88bb-e0cceb90b916.png"  width="500"/>

***LogicAppEndPoint*** is a logic app can accept a post request body which contains a JSON array of SP Site activity items.

<details>
  <summary>
     Here is the body JSON Schema
  </summary>
  
  ```
{
    "properties": {
        "@@odata.nextLink": {
            "type": "string"
        },
        "value": {
            "items": {
                "properties": {
                    "activeFileCount": {
                        "type": "integer"
                    },
                    "anonymousLinkCount": {
                        "type": "integer"
                    },
                    "companyLinkCount": {
                        "type": "integer"
                    },
                    "externalSharing": {
                        "type": "boolean"
                    },
                    "fileCount": {
                        "type": "integer"
                    },
                    "geolocation": {
                        "type": "string"
                    },
                    "isDeleted": {
                        "type": "boolean"
                    },
                    "lastActivityDate": {},
                    "ownerDisplayName": {
                        "type": "string"
                    },
                    "ownerPrincipalName": {
                        "type": "string"
                    },
                    "pageViewCount": {
                        "type": "integer"
                    },
                    "reportPeriod": {
                        "type": "string"
                    },
                    "reportRefreshDate": {
                        "type": "string"
                    },
                    "rootWebTemplate": {
                        "type": "string"
                    },
                    "secureLinkForGuestCount": {
                        "type": "integer"
                    },
                    "secureLinkForMemberCount": {
                        "type": "integer"
                    },
                    "siteId": {
                        "type": "string"
                    },
                    "siteSensitivityLabelId": {
                        "type": "string"
                    },
                    "siteUrl": {
                        "type": "string"
                    },
                    "storageAllocatedInBytes": {
                        "type": "integer"
                    },
                    "storageUsedInBytes": {
                        "type": "integer"
                    },
                    "unmanagedDevicePolicy": {
                        "type": "string"
                    },
                    "visitedPageCount": {
                        "type": "integer"
                    }
                },
                "required": [
                    "reportRefreshDate",
                    "siteId",
                    "siteUrl",
                    "ownerDisplayName",
                    "ownerPrincipalName",
                    "isDeleted",
                    "lastActivityDate",
                    "siteSensitivityLabelId",
                    "externalSharing",
                    "unmanagedDevicePolicy",
                    "geolocation",
                    "fileCount",
                    "activeFileCount",
                    "pageViewCount",
                    "visitedPageCount",
                    "storageUsedInBytes",
                    "storageAllocatedInBytes",
                    "anonymousLinkCount",
                    "companyLinkCount",
                    "secureLinkForGuestCount",
                    "secureLinkForMemberCount",
                    "rootWebTemplate",
                    "reportPeriod"
                ],
                "type": "object"
            },
            "type": "array"
        }
    },
    "type": "object"
}
```

</details>

<img src="https://user-images.githubusercontent.com/8623897/173291914-1cb0d0a0-8c53-4ea9-a8e9-a171df982947.png"  width="500"/>

***TeamsQueryLogicAppEndPoint*** is a logic app can accept a post request body which contains a JSON array of MS 365 Group activity items.

<details>
  <summary>
     Here is the body JSON Schema
  </summary>
  
```
  {
    "properties": {
        "@@odata.context": {
            "type": "string"
        },
        "@@odata.nextLink": {
            "type": "string"
        },
        "value": {
            "items": {
                "properties": {
                    "@@odata.type": {
                        "type": "string"
                    },
                    "exchangeMailboxStorageUsedInBytes": {
                        "type": "integer"
                    },
                    "exchangeMailboxTotalItemCount": {
                        "type": "integer"
                    },
                    "exchangeReceivedEmailCount": {},
                    "externalMemberCount": {
                        "type": "integer"
                    },
                    "groupDisplayName": {
                        "type": "string"
                    },
                    "groupId": {
                        "type": "string"
                    },
                    "groupType": {
                        "type": "string"
                    },
                    "isDeleted": {
                        "type": "boolean"
                    },
                    "lastActivityDate": {},
                    "memberCount": {
                        "type": "integer"
                    },
                    "ownerPrincipalName": {
                        "type": "string"
                    },
                    "reportPeriod": {
                        "type": "string"
                    },
                    "reportRefreshDate": {
                        "type": "string"
                    },
                    "sharePointActiveFileCount": {},
                    "sharePointSiteStorageUsedInBytes": {
                        "type": "integer"
                    },
                    "sharePointTotalFileCount": {
                        "type": "integer"
                    },
                    "yammerLikedMessageCount": {},
                    "yammerPostedMessageCount": {},
                    "yammerReadMessageCount": {}
                },
                "required": [
                    "@@odata.type",
                    "groupId",
                    "reportRefreshDate",
                    "groupDisplayName",
                    "isDeleted",
                    "ownerPrincipalName",
                    "lastActivityDate",
                    "groupType",
                    "memberCount",
                    "externalMemberCount",
                    "exchangeReceivedEmailCount",
                    "sharePointActiveFileCount",
                    "yammerPostedMessageCount",
                    "yammerReadMessageCount",
                    "yammerLikedMessageCount",
                    "exchangeMailboxTotalItemCount",
                    "exchangeMailboxStorageUsedInBytes",
                    "sharePointTotalFileCount",
                    "sharePointSiteStorageUsedInBytes",
                    "reportPeriod"
                ],
                "type": "object"
            },
            "type": "array"
        }
    },
    "type": "object"
}
  ```
  </details>
  
## Build & Run

Call in ***FunAppcs.RunQuick*** or ***FunAppcs.Run*** function. RunQuick will load all groups im memory to reduce the frequency to MS Graph. 

Parameter ***Period*** is for SPSite usage report period. Parameter ***Top*** is for paging in MS Graph query.

```
namespace MSGraphFunctionApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //FunAppcs.SharePointQueryRunQuick("D7","200");
            //Console.ReadKey();
            FunAppcs.TeamsReportRunQuick("D7", "200");
            Console.ReadKey();
        }
    }
}
```

## Try Azure Function version

***AzureFunction_TimmerTrigger\run.csx*** can be used in Timer Trigger Azure Function. It uses the same logic as desktop client. Only difference is output information.

It requires to set 5 Environment variables in Azure AppSettings as well.

![image](https://user-images.githubusercontent.com/8623897/173336487-73e3e3b5-d5bf-4a8d-b94b-22473599a245.png)
![image](https://user-images.githubusercontent.com/8623897/173512266-a3ab7a6b-4450-4b7c-a44d-1690d3407433.png)

