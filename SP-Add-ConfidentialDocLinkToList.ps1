Add-PSSnapin "Microsoft.SharePoint.PowerShell"

$webApplicationUrl = "http://SPFarmUrl"
$listName = "Confidential List"

# Get all the documentss with specific custom column value.
# Only SPSite.Dispose() is called but not SPWeb.Dispose(). This is done because when SPSite is being disposed, 
# all the child objects are disposed as well, so we don't have unnessary overhead
#
# Returned object contains additional fields for the debugging purposes.
# For this specific task we can return only "Item URL" = $_.Url

function Get-AllConfidentialDocs ( [string] $webApplicationUrl ) {
  Get-SPWebApplication $webApplicationUrl | Get-SPSite -Limit All | foreach {
    $siteAssignment = Start-SPAssignment
      $_.AllWebs | foreach {  
        $_.Lists | ? { $_.BaseType -eq "DocumentLibrary" } | foreach { 
          $_.Items | ? { $_["Confidential"] -eq "Yes" } | foreach { 
            $data = @{
              "Item ID" = $_.ID
              "Item URL" = $_.Url
              "Item Title" = $_.Title
              "Item Name" = $_.Name
              "Item Created" = $_["Created"]
              "Item Modified" = $_["Modified"]
              "File Size" = $_.File.Length/1KB
            }
            New-Object PSObject -Property $data
          }
        }
      }
    Stop-SPAssignment $siteAssignment
    $_.Dispose()
  }
}

# Create new list with custom column
function Add-ConfidentialList ( [string] $webApplicationUrl, [string] $listName ) {
  $ErrorActionPreference = "Stop"
  try {
    $fieldName = "URL"
    $web = Get-SPWeb -Identity $webApplicationUrl
    $listDesciption = "List with links"
    $listTemplate = [Microsoft.SharePoint.SPListTemplateType]::GenericList
            
    #Check if List with specific name already exists
    if($web.Lists.TryGetList($listName) -eq $null) {
      $list = $web.Lists.Add($listName, $listDesciption, $listTemplate) 
      Write-Host "List Created Successfully!" -ForegroundColor Green

      if(!$list.Fields.ContainsField($fieldName)){     
        #Add columns to the List
        $list.Fields.Add($fieldName,[Microsoft.SharePoint.SPFieldType]::Text, $IsRequired)
        $list.Update()

        #Update the default view to include the new column
        $view = $list.DefaultView
        $view.ViewFields.Add($fieldName)
        $view.Update()
        Write-host "New Column '$fieldName' added to the List." -ForegroundColor Green
      }
    }  
    else {
      Write-Host "List with specific name already exists!" -ForegroundColor Red
    }
  }
  catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
  }
  finally {
    $ErrorActionPreference = "Continue"
  }
}

# Add data to the new list
function Add-ConfidentialDocLinkToList ( [string] $webApplicationUrl, [string] $listName ) {
  Add-ConfidentialList $webApplicationUrl $listName

  $ErrorActionPreference = "Stop"
  try {
    $web = Get-SPWeb -Identity $webApplicationUrl
    $list = $web.Lists.TryGetList($listName)
    if ($list) {
      Get-AllConfidentialDocs $webApplicationUrl | foreach {
        $newItem = $list.Items.Add()
        $newItem["URL"] = $_["Item URL"]
        $newItem.Update()
      }
    }
    $web.Dispose()
    Write-Host "Items were added to the list" -ForegroundColor Green
  catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
  }
  finally {
    $ErrorActionPreference = "Continue"
  }
}

# Timer to check the script execution time.
Write-Host "Script execution started."
$stopWatch = [Diagnostics.Stopwatch]::StartNew()

Add-ConfidentialDocLinkToList $webApplicationUrl $listName

$stopWatch.Stop()
$elapsed =  $stopWatch.Elapsed.TotalMilliseconds
Write-Host "Script execution time: " $elapsed