﻿#-------------Start Input------------------
$ResourceGroupName= '*************'
$WorkspaceName= '*************'
$SubscriptionId='*************'
#-------------End Input------------------

$option = [System.StringSplitOptions]::RemoveEmptyEntries 

Login-AzureRmAccount -ErrorAction Stop
$Subscription = (Select-AzureRmSubscription -SubscriptionId $($SubscriptionId) -ErrorAction Stop)
$Subscription
$Workspace = (Set-AzureRMOperationalInsightsWorkspace -Name $($WorkspaceName) -ResourceGroupName $($ResourceGroupName) -ErrorAction Stop)
$WorkspaceLocation= $Workspace.Location
$WorkspaceLocationShort= $Workspace.PortalUrl.Split("/",$option)[1].Split(".",$option)[0]
$WorkspaceLocation

Function AdminConsent{

$domain='login.microsoftonline.com'
$mmsDomain = 'login.mms.microsoft.com'
$WorkspaceLocationShort= $Workspace.PortalUrl.Split("/",$option)[1].Split(".",$option)[0]
$WorkspaceLocationShort
$WorkspaceLocation
switch ($WorkspaceLocation) {
       "eastus"   {$OfficeAppClientId="d7eb65b0-8167-4b5d-b371-719a2e5e30cc"; break}
       "westeurope"   {$OfficeAppClientId="c9005da2-023d-40f1-a17a-2b7d91af4ede"; break}
       "southeastasia"   {$OfficeAppClientId="09c5b521-648d-4e29-81ff-7f3a71b27270"; break}
       "australiasoutheast"  {$OfficeAppClientId="f553e464-612b-480f-adb9-14fd8b6cbff8"; break}   
       "westcentralus"  {$OfficeAppClientId="98a2a546-84b4-49c0-88b8-11b011dc8c4e"; break}
       "japaneast"   {$OfficeAppClientId="b07d97d3-731b-4247-93d1-755b5dae91cb"; break}
       "uksouth"   {$OfficeAppClientId="f232cf9b-e7a9-4ebb-a143-be00850cd22a"; break}
       "centralindia"   {$OfficeAppClientId="ffbd6cf4-cba8-4bea-8b08-4fb5ee2a60bd"; break}
       "canadacentral"  {$OfficeAppClientId="c2d686db-f759-43c9-ade5-9d7aeec19455"; break}
       "eastus2"  {$OfficeAppClientId="7eb65b0-8167-4b5d-b371-719a2e5e30cc"; break}
       "westus2"  {$OfficeAppClientId="98a2a546-84b4-49c0-88b8-11b011dc8c4e"; break} #Need to check
       "usgovvirginia" {$OfficeAppClientId="c8b41a87-f8c5-4d10-98a4-f8c11c3933fe"; 
                         $domain='login.microsoftonline.us';
                         $mmsDomain = 'usbn1.login.oms.microsoft.us'; break} # US Gov Virginia
       default {$OfficeAppClientId="55b65fb5-b825-43b5-8972-c8b6875867c1";
                $domain='login.windows-ppe.net'; break} #Int
    }

    $OfficeAppRedirectUrl="https://$($WorkspaceLocationShort).$($mmsDomain)/Office365/Authorize"
    $OfficeAppRedirectUrl
    $domain
    Start-Process -FilePath  "https://$($domain)/common/adminconsent?client_id=$($OfficeAppClientId)&state=12345&redirect_uri=$($OfficeAppRedirectUrl)"
}

AdminConsent -ErrorAction Stop

