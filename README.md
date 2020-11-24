# mecm_device_collections

Automation to create device collections

Example entry with all valid types:

```powershell
$Collections +=
$DummyObject |
Select-Object @{Name="Name"; Expression={"$Ship - Device - Random Devices"}},
@{Name="Queries"; Expression={
    @{"Client install is false or null"="select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from sms_r_system 
    where Client = 0 or Client is null"}
}},
@{Name="DirectMembers"; Expression={@("SOMEHOSTNAME1","SOMEHOSTNAME2")}},
@{Name="IncludeCollections"; Expression={@("All Computers with SnagIt","Software - Google Chrome")}},
@{Name="ExcludeCollections"; Expression={@("All Provisioning Devices","All Unknown Computers")}},
@{Name="LimitingCollection"; Expression={"Device - All Workstations"}},
@{Name="Comment"; Expression={"This describes what the collection is for"}}
```
