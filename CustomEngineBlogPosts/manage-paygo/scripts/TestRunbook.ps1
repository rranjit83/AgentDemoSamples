$sampleFile='./Webhooktestdata.json'

# Kick off an Azure Automation Runbook
$SubscriptionId = "8be5abeb-d89e-4b5e-a459-154ebc5a4601"
$AutomationAccountName = "RRANJITBillingPolicy"
$ResourceGroupName = "Azurevnetforpowerplatform"
$RunbookName = "UnlinkBillingPolicies"

az account set -s $SubscriptionId
az automation runbook start --name $RunbookName --resource-group $ResourceGroupName --automation-account-name $AutomationAccountName --parameters webhookData='@./Webhooktestdata.json'
 