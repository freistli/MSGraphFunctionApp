# MSGraphFunctionApp

## Description

The .Net Core 3.1/5/6 project will get SPO site activity report. If the site owner of one report item is MS Group Mail, the code gets the ownership by querying MS Graph Groups, and replace the ownerPrincipalName with owners email list in the report. 
  
And then posts the report to one logic app to continuous operate the records, for example, save to sharepoint list.

## Prepare: Set 4 environment variables on desktop:

1. ClientID

1. ClientSecret

1. TenantID

1. LogicAppEndPoint

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
            FunAppcs.RunQuick("D7","200");
            Console.ReadKey();
        }
    }
}
```

## Try Azure Function version

***AzureFunction_TimmerTrigger\run.csx*** can be used in Timer Trigger Azure Function. It uses the same logic as desktok client. Only difference is output information.

It requires to set 4 Environment variables in Azure AppSettings as well.

![image](https://user-images.githubusercontent.com/8623897/173336487-73e3e3b5-d5bf-4a8d-b94b-22473599a245.png)

![image](https://user-images.githubusercontent.com/8623897/173337220-67243e79-aca1-4d99-997b-77b65dc6e839.png)

