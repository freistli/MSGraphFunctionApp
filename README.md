# MSGraphFunctionApp

## Description

The .Net Core 3.1/5/6 project will get SPO site activity report. If the site owner of one report item is MS Group Mail, the code gets the ownership by querying MS Graph Groups, and replace the ownerPrincipalName with owners email list in the report. 
 
<img src="https://user-images.githubusercontent.com/8623897/173287501-8709c155-29c0-40b9-ac2b-4c2b8d6b17b0.png"  width="700"/>

And then posts the report to one logic app to continuous operate the records, for example, save to sharepoint list.

<img src="https://user-images.githubusercontent.com/8623897/173287649-5a8e0205-4f84-46f1-94d3-632896522ad4.png"  width="100%"/>

## Prepare: Set 4 environment variables on desktop:

1. ClientID

1. ClientSecret

1. TenantID

1. LogicAppEndPoint

<img src="https://user-images.githubusercontent.com/8623897/173288472-6fdb06c0-0169-4160-88bb-e0cceb90b916.png"  width="500"/>

<details>
  <summary>
    ***LogicAppEndPoint*** is a logic app can accept a post request body which contains a JSON array of SP Site activity items. Here is the body JSON Schema
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

<img src="https://user-images.githubusercontent.com/8623897/173285419-59a5adf3-59ac-4405-b68a-be13144ff776.png"  width="500"/>


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
