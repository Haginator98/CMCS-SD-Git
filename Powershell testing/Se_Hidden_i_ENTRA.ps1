Import-Module microsoft.graph.authentication -ErrorAction SilentlyContinue
connect-Mggraph -scopes 'user.read.all','directory.read.all','User-LifeCycleInfo.Read.All' -ContextScope process

function ConvertTo-PSCustomObject {
  [CmdletBinding()]
  param (
    [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
    [System.Collections.Hashtable] $InputObject
  )
  Process {
    if ($InputObject) {
      $o = New-Object psobject
      foreach ($key in $InputObject.Keys) {
        $value = $InputObject[$key]
        if ($value -and $value.GetType().FullName -match 'System.Object\[\]') {
          if ($value.Count -gt 0 -and $value[0].GetType().FullName -match  'System.Collections.Hashtable') {
            $tempVal = $value | ConvertTo-PSCustomObject
            Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $tempVal
          }
        } elseif ($value -and $value.GetType().FullName -match 'System.Collections.Hashtable') {
          Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue (ConvertTo-PSCustomObject -InputObject $value)
        } else {
          Add-Member -InputObject $o -NotePropertyName $key -NotePropertyValue $value
        }
      }

      Write-Output $o
    }
  }
}

function igall {
    [CmdletBinding()]
    param (
      [string]$Uri,
      [int]$limit=1000
    )
    $nextUri = $uri
    $count=0
    do {
      $result = Invoke-MgGraphRequest -Method GET -uri $nextUri
      $nextUri = $result.'@odata.nextLink'
      if ($result.value) {
        $result.value | ConvertTo-PSCustomObject
      } elseif ($result) {
        $result | ConvertTo-PSCustomObject
      }
      $count +=1
    } while ($nextUri -and ($count -lt $limit))
}

function ig {
    [CmdletBinding()]
    param (
      [string]$Uri
    )
    $result = Invoke-MgGraphRequest -Method GET -uri $uri
    if ($result.value) {
      $result.value | ConvertTo-PSCustomObject
    } else {
      $result | ConvertTo-PSCustomObject
    }
}

function Get-RecoverAttributes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$employeeID
    )

    # Query user(s) by employeeID
    $users = igall "https://graph.microsoft.com/beta/users?`$filter=employeeID eq '$employeeID'"
    if (-not $users) { return }

    foreach ($u in $users) {
        $m = igall "https://graph.microsoft.com/beta/users/$($u.id)/manager" -ErrorAction SilentlyContinue
        $mupn = if ($m) { $m.userPrincipalName } else { 'no manager' }

        $r = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($u.id)?`$select=responsibilities"
        $bd = igall "https://graph.microsoft.com/beta/users/$($u.id)?`$select=Birthday"
        $orgdata = igall "https://graph.microsoft.com/beta/users/$($u.id)" | Select-Object -ExpandProperty employeeOrgData -ErrorAction SilentlyContinue

        $u | Add-Member -Force -NotePropertyName ManagerUpn -NotePropertyValue $mupn
        if ($orgdata) {
            $u | Add-Member -Force -NotePropertyName division -NotePropertyValue $orgdata.division
            $u | Add-Member -Force -NotePropertyName costcenter -NotePropertyValue $orgdata.costcenter
        }
        $u | Add-Member -Force -NotePropertyName birthday -NotePropertyValue ($bd.birthday)
        $u | Add-Member -Force -NotePropertyName responsibilities -NotePropertyValue ($r.responsibilities)

        $u | Select-Object ID,employeeID, userPrincipalName, DisplayName, GivenName, Surname, mail, accountEnabled, birthday, ManagerUpn, employeeHireDate, department, companyName, country, mobilephone, telephoneNumber, costcenter, division, extension_8bd11a3d16ec456c85d85fcffb1c3113_en_title, jobTitle, extension_8bd11a3d16ec456c85d85fcffb1c3113_gender, extension_8bd11a3d16ec456c85d85fcffb1c3113_middlename, extension_8bd11a3d16ec456c85d85fcffb1c3113_employmentStatus, extension_8bd11a3d16ec456c85d85fcffb1c3113_ReasonforLeaving, extension_8bd11a3d16ec456c85d85fcffb1c3113_LeaveDate, extension_8bd11a3d16ec456c85d85fcffb1c3113_usertype, extension_8bd11a3d16ec456c85d85fcffb1c3113_region, extension_8bd11a3d16ec456c85d85fcffb1c3113_area, streetAddress, City, postalcode, usageLocation, preferredLanguage, employeeLeaveDateTime, employeeType, responsibilities
    }
}

Write-Output "Merged script saved: $(Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)"
