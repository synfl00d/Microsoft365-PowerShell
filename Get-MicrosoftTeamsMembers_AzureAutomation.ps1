<#
    .DESCRIPTION
        Gets Teams and Teams member informations from the Microsoft Graph Endpoint and writes the result into SQL Tables

    .NOTES
        AUTHOR: synfl00d
        LASTEDIT: June 15, 2022
#>

# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process
# Connect to Azure with system-assigned managed identity
$AzureContext = (Connect-AzAccount -Identity).context
# set and store context
$AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext

# Global Variables
$clientID = "{clientid}"
$clientSecret = Get-AzKeyVaultSecret -VaultName '{vaultname[' -Name '{secretname}' -AsPlainText
$tenantID = "{tenantid}"

#Database Connection
$PSCredential = (Get-AutomationPSCredential -Name '{credentialname}')
$Server = "{server}"
$ServerPort = 1433
$Database = "{database}"

#Truncate SQL Tables before filling them again
$PreExecuteCommand = "EXEC [graph].[sp_TruncateTeamsTables]"
try
{
	if ($PSCredential -eq $null) 
	{ 
		throw "Could not retrieve '$PSCredential' credential asset. Check that you created this first in the Automation service." 
	}   
	$SqlUsername = $PSCredential.UserName 
	$SqlPass = $PSCredential.GetNetworkCredential().Password 
	$Conn = New-Object System.Data.SqlClient.SqlConnection("Server=tcp:$Server,$ServerPort;Database=$Database;User ID=$SqlUsername;Password=$SqlPass;Trusted_Connection=False;Encrypt=True;Connection Timeout=30;") 
	$Conn.Open() 
	$Cmd=new-object system.Data.SqlClient.SqlCommand($PreExecuteCommand, $Conn)
	$Cmd.CommandTimeout=120
	$Cmd.ExecuteNonQuery();
	$Conn.Close()
}
catch
{
	$ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName           

	Write-Error $FailedItem
	Write-Error $ErrorMessage
	throw "Error while trying to connect to SQL Instance"
}

#Connect to GRAPH API
$tokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}
$connectUri = ("https://login.microsoftonline.com/{0}/oauth2/v2.0/token" -f $tenantID)
Write-Output ("Connect to Graph API with {0} ...." -f $connectUri)

$tokenResponse = Invoke-RestMethod -Uri $connectUri -Method POST -Body $tokenBody
$headers = @{
    "Authorization" = "Bearer $($tokenResponse.access_token)"
    "Content-type"  = "application/json"
}

try
{
    #Set GRAPH Endpoint
    $URLTeam = "https://graph.microsoft.com/v1.0/groups"

    #Get first batch of teams
	Write-Warning "[INFO] Getting first batch...."
    $Teams = Invoke-RestMethod -Headers $headers -Uri $URLTeam -Method GET

    $Teams = ""
    do 
    {
        #Get next batch of teams
		Write-Warning ("[INFO] Get next batch with URL: {0}" -f $Teams.'@odata.nextLink')
		$link = $Teams.'@odata.nextLink'
        $Teams = Invoke-RestMethod -Headers $headers -Uri $Teams.'@odata.nextLink' -Method GET

        #Check if last page is reached
		if($link -eq $Teams.'@odata.nextLink')
		{
			Write-Warning "[Info] End of page reached. Return"
			Return 0
			break;
		}
        
        #Iterate through Teams teams
        foreach ($line in $Teams.value)
        {
            #Create Teams Object
            $lineObject = [PSCustomObject]@{
                GroupId = $line.id
                DeletedDateTime = $line.deletedDateTime
                Classification = $line.classification
                CreatedDateTime = $line.createdDateTime | Out-String
                Description = $line.description
                DisplayName = $line.displayName
                ExpirationDateTime = $line.expirationDateTime
                GroupType = $line.groupTypes | Out-String
                ResourceProvisioningOptions = $line.resourceProvisioningOptions | Out-String
                ResourceBehaviorOptions = $line.ResourceBehaviorOptions | Out-String
                Mail = $line.mail
                MailEnabled = $line.mailEnabled
                MailNickname = $line.mailNickname
                MembershipRule = $line.membershipRule
                MembershipRuleProcessingState = $line.membershipRuleProcessingState
                ProxyAddresses = $line.proxyAddresses | Out-String
                RenewedDateTime = $line.renewedDateTime
                SecurityEnabled = $line.securityEnabled
                SecurityIdentifier = $line.securityIdentifier
                Visibility = $line.visibility
            }
            #Write Team to SQL
            $lineObject | Write-ObjectToSQL -Server $Server -Database $Database -TableName TeamsGroups -SchemaName graph -Credential $PSCredential
            Write-Output ("[INFO] Inserting: {0}" -f $lineObject.DisplayName)

            #Get Members
            $team = [string]::Format("https://graph.microsoft.com/v1.0/teams/{0}/members",$line.id)
            $members = Invoke-RestMethod -Headers $headers -Uri $team -Method GET

            #Write Members to SQL
            $membersObject = foreach ($member in $members.value)
            {
                [PSCustomObject]@{
                    Id = $member.id
                    TeamId = $line.id
                    Roles = $member.roles[0] | Out-String
                    DisplayName = $member.displayName
                    VisibleHistoryStartDateTime = $member.visibleHistoryStartDateTime
                    UserId = $member.userId
                    Email = $member.email
                    TenantId = $member.tenantId
                } 
            }
            #Write Team Members to SQL
            $membersObject | Write-ObjectToSQL -Server $Server -Database $Database -TableName TeamsGroupMembers -SchemaName graph -Credential $PSCredential
        }       
    } while ($Teams.'@odata.nextLink' -ne "")
}
catch
{
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName

	Write-Error $ErrorMessage
	Write-Error $FailedItem
	throw "Error while getting Team"
}