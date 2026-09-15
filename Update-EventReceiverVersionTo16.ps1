<#
  .SYNOPSIS
  Updates SharePoint list event receiver assembly references from version 15 to version 16.

  .EXAMPLE
  Get-SPSite -Limit All | .\Update-EventReceiverVersionTo16.ps1 -WhatIf
  Shows the event receiver assembly references for all piped site collections.

  .EXAMPLE
  Get-SPSite <URL> | .\Update-EventReceiverVersionTo16.ps1 -WhatIf
  Shows the event receiver assembly references for the specified site collection.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
  [Microsoft.SharePoint.SPSite]$Site
)

begin {
  # Ensure the SharePoint PowerShell Snapin is loaded
  if ($null -eq (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    # Required for SP2016/SP2019 environments
    Add-PSSnapin Microsoft.SharePoint.PowerShell
  }

  # Define the target assembly signatures
  $oldAssembly = "Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
  $newAssembly = "Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
}

process {
  Write-Host "Starting EventReceiver update process for site collection $($Site.Url)..." -ForegroundColor Cyan
  try {
    # Loop through all Subwebs (SPWeb) within the Site Collection
    $webUrls = @((Get-SPWeb -Site $Site -Limit All).Url)
    foreach ($webUrl in $webUrls | Where-Object { $null -ne $_ }) {
      Write-Host "Processing Web: $webUrl" -ForegroundColor Cyan
      $web = Get-SPWeb -Identity $webUrl

      try {
        # CRITICAL FIX: Force immediate evaluation using @(...) to create a static array copy.
        # This prevents the "Collection was modified" exception during .Update() execution.
        # Loop through all Lists within the Web
        foreach ($list in @($web.Lists)) {
          # CRITICAL FIX: Force immediate evaluation using @(...) to create a static array copy.
          # This prevents the "Collection was modified" exception during .Update() execution.
          $receiversToUpdate = @($list.EventReceivers | Where-Object { $_.Assembly -eq $oldAssembly })

          foreach ($receiver in $receiversToUpdate) {
            # Output the affected list URL and the specific EventReceiver Class Name
            Write-Host "Found in List: $($web.Url)/$($list.RootFolder.Url) | EventReceiver Class: $($receiver.Class)" -ForegroundColor Yellow

            # Handle -WhatIf and -Confirm parameters automatically via ShouldProcess
            if ($PSCmdlet.ShouldProcess("List: $($list.Title) ($($web.Url))", "Update EventReceiver Assembly from v15 to v16 for Class: $($receiver.Class)")) {

              # Update the assembly property and commit changes
              $receiver.Assembly = $newAssembly
              $receiver.Update()

              Write-Host "-> Successfully updated!" -ForegroundColor Green
            }
          }
        }
      }
      catch {
        Write-Error "Error processing Web instance ($($web.Url)): $_"
      }
      finally {
        # Properly dispose SPWeb objects to prevent memory leaks
        if ($null -ne $web) { $web.Dispose() }
      }
    }
  }
  catch {
    Write-Error "Error processing Site Collection ($($Site.Url)): $_"
  }
}
