# MSGraphFunctionApp

## Set 4 environment variables on desktop:

ClientID

ClientSecret

TenantID

LogicAppEndPoint

***LogicAppEndPoint*** is a logic app can accept a post request body which contains a JSON array of SP Site activity items.

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
![image](https://user-images.githubusercontent.com/8623897/173281497-4d8e56a3-ead9-4d67-8d4f-5cac09987c8a.png)


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
