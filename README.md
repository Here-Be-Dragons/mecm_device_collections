# mecm_device_collections

## Overview 

Automation to create device collections and manage membership rules:
* Query Rules
* Direct Membership
* Include Collections
* Exclude Collections

**Note:** Query rules must match output from MECM exactly, else they will be deleted and re-created every execution.

## Usage

Example entry with all valid types:

```json
{
    "Name": "$Folder - Device - Random Devices",
    "LimitingCollection": "$Folder - Device - All Workstations",
    "Comment": "Finance Workstation-Class Devices",
    "Queries": [
        {
        "Name": "Client install is false or null",
        "Expression": "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from sms_r_system where Client = 0 or Client is null"
        }
    ],
    "DirectMembers": [
        "COMPUTER1",
        "COMPUTER2"
    ],
    "IncludeCollections": [
        "All Computers with SnagIt",
        "Software - Google Chrome"
    ],
    "ExcludeCollections": [
        "All Provisioning Devices",
        "All Unknown Computers"
    ]
}
```
## Other Notes

* Parent/Limiting collections must exist before child collections are created.  `.json` files are loaded in alphabetical order, so force the order by appending a number or otherwise making sure Limiting collections are created before they're needed.
* Variables **are** expanded from within the json files, so you can extend templatability by adding variables to things like OUs or other queries, which will be converted into words when the pipeline executes.
* Device Collections are never deleted by this script, however they will be renamed to facilitate manual removal from within the MECM console.