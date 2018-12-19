#######################################################################################################
#
# Add-Diskspace
#
# Author: Paul Henry
#
# 
#
# Script contains 5 main cmdlets and 5 subsidiary cmdlets. Requires PowerCLI to be installed.
#
# Date: 12/6/2018
# 
# Version: 1.0.3
#
#
<#
  Cmdlets:

  "Initialize-Connection <ServerName>"
  Conducts Task 1. <ServerName> is the VSphere server you want to connect to.

  "Remove-StaleLogs <ServerName>"
  Conducts Task 2. <ServerName> is server you want to remove logs from.

  "Get-DiskSpaceUpgrade <ServerName>"
  Conducts Task 3-5. <ServerName> is the server you want to see upgrade options for.

  "Get-DiskSpaceResult <ServerName>"
  Conducts Task 3-6. <ServerName> is the server you want to see upgrade options for. Includes VSphere info. Requires Vsphere connection.

  "Add-DiskSpace <ServerName>"
  Conducts Tasks 1-11. <ServerName> is the server you want to add space to. Requires Vsphere connection.

 Tasks:

  1. Set up PowerCLI environment and Connect to Vsphere Server
  2. Clean out common logs over 8 days old on the C drive
  3. Scan all drives to determine freespace percentage.
  4. For drives with less than 30% of free space, calculates the amount to bring it to 35% free space.
  5. Reports the drives and changes needed to console.
  6. Correlates the partitions, volumes and Hard Disks in Vsphere. Adds this in to report given in Task 4.
  7. For each drive that needs space, adds space in VCenter, then extends the drives using a locally-executed diskpart script.
  8. Reruns the drive scan and reports to console for validation
  9. Fail if the servername includes MAIL or DFS
  10. Fail if resulting drive would be greater than 2 TB.
  11. Fail if the additions will bring the datastore's available space below 30%.

   
#>
#######################################################################################################

param(
$ComputerName,
$Notes,
$newip,
$oldip)

FUNCTION Initialize-Connection($ComputerName)
{

# This loads the Vsphere PowerCLI environment so you can talk to the VCenter server

IF(Test-path -Path 'C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1'){
& 'C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1'
}ELSE{
Write-Host -ForegroundColor Red "PowerCLI is not installed."
exit

}
# This actually connects you to the server
Connect-VIServer $ComputerName -Credential (Get-Credential)

}

Function Join-Object
{
    <#
    .SYNOPSIS
        Join data from two sets of objects based on a common value
    .DESCRIPTION
        Join data from two sets of objects based on a common value
        For more details, see the accompanying blog post:
            http://ramblingcookiemonster.github.io/Join-Object/
        For even more details,  see the original code and discussions that this borrows from:
            Dave Wyatt's Join-Object - http://powershell.org/wp/forums/topic/merging-very-large-collections
            Lucio Silveira's Join-Object - http://blogs.msdn.com/b/powershell/archive/2012/07/13/join-object.aspx
    .PARAMETER Left
        'Left' collection of objects to join.  You can use the pipeline for Left.
        The objects in this collection should be consistent.
        We look at the properties on the first object for a baseline.
    
    .PARAMETER Right
        'Right' collection of objects to join.
        The objects in this collection should be consistent.
        We look at the properties on the first object for a baseline.
    .PARAMETER LeftJoinProperty
        Property on Left collection objects that we match up with RightJoinProperty on the Right collection
    .PARAMETER RightJoinProperty
        Property on Right collection objects that we match up with LeftJoinProperty on the Left collection
    .PARAMETER LeftProperties
        One or more properties to keep from Left.  Default is to keep all Left properties (*).
        Each property can:
            - Be a plain property name like "Name"
            - Contain wildcards like "*"
            - Be a hashtable like @{Name="Product Name";Expression={$_.Name}}.
                 Name is the output property name
                 Expression is the property value ($_ as the current object)
                
                 Alternatively, use the Suffix or Prefix parameter to avoid collisions
                 Each property using this hashtable syntax will be excluded from suffixes and prefixes
    .PARAMETER RightProperties
        One or more properties to keep from Right.  Default is to keep all Right properties (*).
        Each property can:
            - Be a plain property name like "Name"
            - Contain wildcards like "*"
            - Be a hashtable like @{Name="Product Name";Expression={$_.Name}}.
                 Name is the output property name
                 Expression is the property value ($_ as the current object)
                
                 Alternatively, use the Suffix or Prefix parameter to avoid collisions
                 Each property using this hashtable syntax will be excluded from suffixes and prefixes
    .PARAMETER Prefix
        If specified, prepend Right object property names with this prefix to avoid collisions
        Example:
            Property Name                   = 'Name'
            Suffix                          = 'j_'
            Resulting Joined Property Name  = 'j_Name'
    .PARAMETER Suffix
        If specified, append Right object property names with this suffix to avoid collisions
        Example:
            Property Name                   = 'Name'
            Suffix                          = '_j'
            Resulting Joined Property Name  = 'Name_j'
    .PARAMETER Type
        Type of join.  Default is AllInLeft.
        AllInLeft will have all elements from Left at least once in the output, and might appear more than once
          if the where clause is true for more than one element in right, Left elements with matches in Right are
          preceded by elements with no matches.
          SQL equivalent: outer left join (or simply left join)
        AllInRight is similar to AllInLeft.
        
        OnlyIfInBoth will cause all elements from Left to be placed in the output, only if there is at least one
          match in Right.
          SQL equivalent: inner join (or simply join)
         
        AllInBoth will have all entries in right and left in the output. Specifically, it will have all entries
          in right with at least one match in left, followed by all entries in Right with no matches in left, 
          followed by all entries in Left with no matches in Right.
          SQL equivalent: full join
    .EXAMPLE
        #
        #Define some input data.
        $l = 1..5 | Foreach-Object {
            [pscustomobject]@{
                Name = "jsmith$_"
                Birthday = (Get-Date).adddays(-1)
            }
        }
        $r = 4..7 | Foreach-Object{
            [pscustomobject]@{
                Department = "Department $_"
                Name = "Department $_"
                Manager = "jsmith$_"
            }
        }
        #We have a name and Birthday for each manager, how do we find their department, using an inner join?
        Join-Object -Left $l -Right $r -LeftJoinProperty Name -RightJoinProperty Manager -Type OnlyIfInBoth -RightProperties Department
            # Name    Birthday             Department  
            # ----    --------             ----------  
            # jsmith4 4/14/2015 3:27:22 PM Department 4
            # jsmith5 4/14/2015 3:27:22 PM Department 5
    .EXAMPLE  
        #
        #Define some input data.
        $l = 1..5 | Foreach-Object {
            [pscustomobject]@{
                Name = "jsmith$_"
                Birthday = (Get-Date).adddays(-1)
            }
        }
        $r = 4..7 | Foreach-Object{
            [pscustomobject]@{
                Department = "Department $_"
                Name = "Department $_"
                Manager = "jsmith$_"
            }
        }
        #We have a name and Birthday for each manager, how do we find all related department data, even if there are conflicting properties?
        $l | Join-Object -Right $r -LeftJoinProperty Name -RightJoinProperty Manager -Type AllInLeft -Prefix j_
            # Name    Birthday             j_Department j_Name       j_Manager
            # ----    --------             ------------ ------       ---------
            # jsmith1 4/14/2015 3:27:22 PM                                    
            # jsmith2 4/14/2015 3:27:22 PM                                    
            # jsmith3 4/14/2015 3:27:22 PM                                    
            # jsmith4 4/14/2015 3:27:22 PM Department 4 Department 4 jsmith4  
            # jsmith5 4/14/2015 3:27:22 PM Department 5 Department 5 jsmith5  
    .EXAMPLE
        #
        #Hey!  You know how to script right?  Can you merge these two CSVs, where Path1's IP is equal to Path2's IP_ADDRESS?
        
        #Get CSV data
        $s1 = Import-CSV $Path1
        $s2 = Import-CSV $Path2
        #Merge the data, using a full outer join to avoid omitting anything, and export it
        Join-Object -Left $s1 -Right $s2 -LeftJoinProperty IP_ADDRESS -RightJoinProperty IP -Prefix 'j_' -Type AllInBoth |
            Export-CSV $MergePath -NoTypeInformation
    .EXAMPLE
        #
        # "Hey Warren, we need to match up SSNs to Active Directory users, and check if they are enabled or not.
        #  I'll e-mail you an unencrypted CSV with all the SSNs from gmail, what could go wrong?"
        
        # Import some SSNs. 
        $SSNs = Import-CSV -Path D:\SSNs.csv
        #Get AD users, and match up by a common value, samaccountname in this case:
        Get-ADUser -Filter "samaccountname -like 'wframe*'" |
            Join-Object -LeftJoinProperty samaccountname -Right $SSNs `
                        -RightJoinProperty samaccountname -RightProperties ssn `
                        -LeftProperties samaccountname, enabled, objectclass
    .NOTES
        This borrows from:
            Dave Wyatt's Join-Object - http://powershell.org/wp/forums/topic/merging-very-large-collections/
            Lucio Silveira's Join-Object - http://blogs.msdn.com/b/powershell/archive/2012/07/13/join-object.aspx
        Changes:
            Always display full set of properties
            Display properties in order (left first, right second)
            If specified, add suffix or prefix to right object property names to avoid collisions
            Use a hashtable rather than ordereddictionary (avoid case sensitivity)
    .LINK
        http://ramblingcookiemonster.github.io/Join-Object/
    .FUNCTIONALITY
        PowerShell Language
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine = $true)]
        [object[]] $Left,

        # List to join with $Left
        [Parameter(Mandatory=$true)]
        [object[]] $Right,

        [Parameter(Mandatory = $true)]
        [string] $LeftJoinProperty,

        [Parameter(Mandatory = $true)]
        [string] $RightJoinProperty,

        [object[]]$LeftProperties = '*',

        # Properties from $Right we want in the output.
        # Like LeftProperties, each can be a plain name, wildcard or hashtable. See the LeftProperties comments.
        [object[]]$RightProperties = '*',

        [validateset( 'AllInLeft', 'OnlyIfInBoth', 'AllInBoth', 'AllInRight')]
        [Parameter(Mandatory=$false)]
        [string]$Type = 'AllInLeft',

        [string]$Prefix,
        [string]$Suffix
    )
    Begin
    {
        function AddItemProperties($item, $properties, $hash)
        {
            if ($null -eq $item)
            {
                return
            }

            foreach($property in $properties)
            {
                $propertyHash = $property -as [hashtable]
                if($null -ne $propertyHash)
                {
                    $hashName = $propertyHash["name"] -as [string]         
                    $expression = $propertyHash["expression"] -as [scriptblock]

                    $expressionValue = $expression.Invoke($item)[0]
            
                    $hash[$hashName] = $expressionValue
                }
                else
                {
                    foreach($itemProperty in $item.psobject.Properties)
                    {
                        if ($itemProperty.Name -like $property)
                        {
                            $hash[$itemProperty.Name] = $itemProperty.Value
                        }
                    }
                }
            }
        }

        function TranslateProperties
        {
            [cmdletbinding()]
            param(
                [object[]]$Properties,
                [psobject]$RealObject,
                [string]$Side)

            foreach($Prop in $Properties)
            {
                $propertyHash = $Prop -as [hashtable]
                if($null -ne $propertyHash)
                {
                    $hashName = $propertyHash["name"] -as [string]         
                    $expression = $propertyHash["expression"] -as [scriptblock]

                    $ScriptString = $expression.tostring()
                    if($ScriptString -notmatch 'param\(')
                    {
                        Write-Verbose "Property '$HashName'`: Adding param(`$_) to scriptblock '$ScriptString'"
                        $Expression = [ScriptBlock]::Create("param(`$_)`n $ScriptString")
                    }
                
                    $Output = @{Name =$HashName; Expression = $Expression }
                    Write-Verbose "Found $Side property hash with name $($Output.Name), expression:`n$($Output.Expression | out-string)"
                    $Output
                }
                else
                {
                    foreach($ThisProp in $RealObject.psobject.Properties)
                    {
                        if ($ThisProp.Name -like $Prop)
                        {
                            Write-Verbose "Found $Side property '$($ThisProp.Name)'"
                            $ThisProp.Name
                        }
                    }
                }
            }
        }

        function WriteJoinObjectOutput($leftItem, $rightItem, $leftProperties, $rightProperties)
        {
            $properties = @{}

            AddItemProperties $leftItem $leftProperties $properties
            AddItemProperties $rightItem $rightProperties $properties

            New-Object psobject -Property $properties
        }

        #Translate variations on calculated properties.  Doing this once shouldn't affect perf too much.
        foreach($Prop in @($LeftProperties + $RightProperties))
        {
            if($Prop -as [hashtable])
            {
                foreach($variation in ('n','label','l'))
                {
                    if(-not $Prop.ContainsKey('Name') )
                    {
                        if($Prop.ContainsKey($variation) )
                        {
                            $Prop.Add('Name',$Prop[$Variation])
                        }
                    }
                }
                if(-not $Prop.ContainsKey('Name') -or $Prop['Name'] -like $null )
                {
                    Throw "Property is missing a name`n. This should be in calculated property format, with a Name and an Expression:`n@{Name='Something';Expression={`$_.Something}}`nAffected property:`n$($Prop | out-string)"
                }


                if(-not $Prop.ContainsKey('Expression') )
                {
                    if($Prop.ContainsKey('E') )
                    {
                        $Prop.Add('Expression',$Prop['E'])
                    }
                }
            
                if(-not $Prop.ContainsKey('Expression') -or $Prop['Expression'] -like $null )
                {
                    Throw "Property is missing an expression`n. This should be in calculated property format, with a Name and an Expression:`n@{Name='Something';Expression={`$_.Something}}`nAffected property:`n$($Prop | out-string)"
                }
            }        
        }

        $leftHash = @{}
        $rightHash = @{}

        # Hashtable keys can't be null; we'll use any old object reference as a placeholder if needed.
        $nullKey = New-Object psobject
        
        $bound = $PSBoundParameters.keys -contains "InputObject"
        if(-not $bound)
        {
            [System.Collections.ArrayList]$LeftData = @()
        }
    }
    Process
    {
        #We pull all the data for comparison later, no streaming
        if($bound)
        {
            $LeftData = $Left
        }
        Else
        {
            foreach($Object in $Left)
            {
                [void]$LeftData.add($Object)
            }
        }
    }
    End
    {
        foreach ($item in $Right)
        {
            $key = $item.$RightJoinProperty

            if ($null -eq $key)
            {
                $key = $nullKey
            }

            $bucket = $rightHash[$key]

            if ($null -eq $bucket)
            {
                $bucket = New-Object System.Collections.ArrayList
                $rightHash.Add($key, $bucket)
            }

            $null = $bucket.Add($item)
        }

        foreach ($item in $LeftData)
        {
            $key = $item.$LeftJoinProperty

            if ($null -eq $key)
            {
                $key = $nullKey
            }

            $bucket = $leftHash[$key]

            if ($null -eq $bucket)
            {
                $bucket = New-Object System.Collections.ArrayList
                $leftHash.Add($key, $bucket)
            }

            $null = $bucket.Add($item)
        }

        $LeftProperties = TranslateProperties -Properties $LeftProperties -Side 'Left' -RealObject $LeftData[0]
        $RightProperties = TranslateProperties -Properties $RightProperties -Side 'Right' -RealObject $Right[0]

        #I prefer ordered output. Left properties first.
        [string[]]$AllProps = $LeftProperties

        #Handle prefixes, suffixes, and building AllProps with Name only
        $RightProperties = foreach($RightProp in $RightProperties)
        {
            if(-not ($RightProp -as [Hashtable]))
            {
                Write-Verbose "Transforming property $RightProp to $Prefix$RightProp$Suffix"
                @{
                    Name="$Prefix$RightProp$Suffix"
                    Expression=[scriptblock]::create("param(`$_) `$_.'$RightProp'")
                }
                $AllProps += "$Prefix$RightProp$Suffix"
            }
            else
            {
                Write-Verbose "Skipping transformation of calculated property with name $($RightProp.Name), expression:`n$($RightProp.Expression | out-string)"
                $AllProps += [string]$RightProp["Name"]
                $RightProp
            }
        }

        $AllProps = $AllProps | Select -Unique

        Write-Verbose "Combined set of properties: $($AllProps -join ', ')"

        foreach ( $entry in $leftHash.GetEnumerator() )
        {
            $key = $entry.Key
            $leftBucket = $entry.Value

            $rightBucket = $rightHash[$key]

            if ($null -eq $rightBucket)
            {
                if ($Type -eq 'AllInLeft' -or $Type -eq 'AllInBoth')
                {
                    foreach ($leftItem in $leftBucket)
                    {
                        WriteJoinObjectOutput $leftItem $null $LeftProperties $RightProperties | Select $AllProps
                    }
                }
            }
            else
            {
                foreach ($leftItem in $leftBucket)
                {
                    foreach ($rightItem in $rightBucket)
                    {
                        WriteJoinObjectOutput $leftItem $rightItem $LeftProperties $RightProperties | Select $AllProps
                    }
                }
            }
        }

        if ($Type -eq 'AllInRight' -or $Type -eq 'AllInBoth')
        {
            foreach ($entry in $rightHash.GetEnumerator())
            {
                $key = $entry.Key
                $rightBucket = $entry.Value

                $leftBucket = $leftHash[$key]

                if ($null -eq $leftBucket)
                {
                    foreach ($rightItem in $rightBucket)
                    {
                        WriteJoinObjectOutput $null $rightItem $LeftProperties $RightProperties | Select $AllProps
                    }
                }
            }
        }
    }
}

Function Get-DiskSpaceUpgrade($ComputerName)
{

# Query target Server for all Partition with non-zero space
If ((Test-Connection -ComputerName $ComputerName -count 1)){
$AllPartitions = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName | where {$_.size -ne $null} | sort DeviceId
}
$AllPartitionsResult = @()

IF($allPartitions){foreach ($EachPartition in $AllPartitions)
    {

    # Go through each Partition and see if it needs have space added
    $AddAmount = $null
    $PartitionCap = [math]::Round((($EachPartition.size)/1024/1024/1024))
    $CheckToAdd = [math]::Round((((($EachPartition.size)*30)-(($EachPartition.freespace)*100))/70)/1024/1024/1024)
    $AddAmount = [math]::Round((((($EachPartition.size)*35)-(($EachPartition.freespace)*100))/65)/1024/1024/1024)

    # Construct the Reporting Object
    $PartitionResult = New-Object -TypeName PSObject
        If($CheckToAdd -gt 0){
            IF($AddAmount % 5){

            $AddAmount = $AddAmount + (5 - $AddAmount % 5)}
            $PartitionCapAndAdded = [math]::Round(($AddAmount + $PartitionCap))
            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $ComputerName
            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'DriveLetter' -Value $EachPartition.DeviceID
            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'CurrentSize' -Value $PartitionCap
            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'AmountToAddGB' -Value $AddAmount
            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'ResultSizeGB' -Value $PartitionCapAndAdded

            } ELSE {

            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'ServerName' -Value $ComputerName
            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'DriveLetter' -Value $EachPartition.DeviceID
            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'CurrentSize' -Value $PartitionCap
            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'AmountToAddGB' -Value 0
            $PartitionResult | Add-Member -MemberType NoteProperty -Name 'ResultSizeGB' -Value $PartitionCap

            }

            $AllPartitionsResult += $PartitionResult




            
    }
    }

RETURN $AllPartitionsResult
}

Function Get-VMwareToDiskMatch($ComputerName)
{
    # Query Vsphere server to get View
    $VM = Get-VM $ComputerName
    $VMSummaries = @()
    $DiskMatches = @()
    $VMView = $VM | Get-View

    # Go through each Virtual SCSI Controller and Disk to get their IDs
    ForEach ($VirtualSCSIController in ($VMView.Config.Hardware.Device | Where {$_.DeviceInfo.Label -match "SCSI Controller"}))
        {
        ForEach ($VirtualDiskDevice in ($VMView.Config.Hardware.Device | Where {$_.ControllerKey -eq $VirtualSCSIController.Key}))
            {

            # Construct Result Object for View
            $VMSummary = "" | Select VM, HostName, PowerState, DiskFile, DiskName, DiskSize, SCSIController, SCSITarget
            $VMSummary.VM = $VM.Name
            $VMSummary.HostName = $VMView.Guest.HostName
            $VMSummary.PowerState = $VM.PowerState
            $VMSummary.DiskFile = $VirtualDiskDevice.Backing.FileName
            $VMSummary.DiskName = $VirtualDiskDevice.DeviceInfo.Label
            $VMSummary.DiskSize = $VirtualDiskDevice.CapacityInKB * 1KB
            $VMSummary.SCSIController = $VirtualSCSIController.BusNumber
            $VMSummary.SCSITarget = $VirtualDiskDevice.UnitNumber

            # Add Result view for this disk to Main Result Object

            $VMSummaries += $VMSummary

            }

        }
    $Disks = Get-WmiObject -Class Win32_DiskDrive -ComputerName $VM.Name
    $Difference = $Disks.SCSIPort | sort-object -Descending | Select -last 1 
            foreach ($device in $VMSummaries)
        {
            $Disks | % {if((($_.SCSIPort - $Difference) -eq $device.SCSIController) -and ($_.SCSITargetID -eq $device.SCSITarget))
                {
                    $DiskMatch = "" | Select VMWareDisk, VMWareDiskSize, WindowsDeviceID, WindowsDiskSize 
                    $DiskMatch.VMWareDisk = $device.DiskName
                    $DiskMatch.WindowsDeviceID = $_.DeviceID.Substring(4)
                    $DiskMatch.VMWareDiskSize = $device.DiskSize/1gb
                    $DiskMatch.WindowsDiskSize =  [decimal]::round($_.Size/1gb)
                    $DiskMatches+=$DiskMatch
             
                    }
                }   
        }




RETURN $DiskMatches
}

Function Get-DiskToPartitionMatch($ComputerName)
{
    $VM = Get-VM $ComputerName
    $WinDevIDs = Get-VMwareToDiskMatch $ComputerName
    $DiskDrivesToDiskPartition = Get-WmiObject -Class Win32_DiskDriveToDiskPartition -ComputerName $ComputerName
    $WinDevsToDrives = @()

    foreach($ID in $WinDevIDs){
        $PreRes = $null
        $PreRes = $DiskDrivesToDiskPartition.__RELPATH -match $ID.WindowsDeviceID
            
        for($i=0;$i -lt $PreRes.Count;$i++){
            $matches =$null
            $WinDev = "" | Select PhysicalDrive, DiskAndPart
            $PreRes[$i] -match '.*(Disk\s#\d+\,\sPartition\s#\d+).*' |out-null
            $WinDev.PhysicalDrive = $ID.WindowsDeviceID
            $WinDev.DiskAndPart = $matches[1] 
            $WinDevsToDrives+=$WinDev
            }
        }

    $LogicalDiskToPartition = Get-WmiObject -Class Win32_LogicalDiskToPartition -ComputerName $vm.name
    IF(($LogicalDiskToPartition | measure).count -eq 1){$singledrive = $true}
        $final = @()
        foreach($drive in $WinDevsToDrives){
            $matches =$null
            $WinDevVol = "" | Select PhysicalDrive, DiskAndPart, VolumeLabel
            $WinDevVol.PhysicalDrive = $drive.PhysicalDrive
            $WinDevVol.DiskAndPart = $drive.DiskAndPart

            $Res = $LogicalDiskToPartition.__RELPATH -match $drive.DiskAndPart


            $Res[0] -match '.*Win32_LogicalDisk.DeviceID=\\"([A-Z]\:).*' 
            if($matches){
                $WinDevVol.VolumeLabel = $matches[1]
                }
            if($singledrive){
            if($res){
            $intermediate = $null
            $end = $end

            $intermediate = $LogicalDiskToPartition.__RELPATH.substring(0,($LogicalDiskToPartition.__RELPATH.length - 3))
            $end = $intermediate.substring($LogicalDiskToPartition.__RELPATH.length - 5)
            
            $windevvol.VolumeLabel = $end}}
            $final+=$WinDevVol
            
        }

RETURN $final
}

Function Get-DriveSpaceAdditionSum($ComputerName)
{
    $DriveAdditions = Get-DiskSpaceUpgrade $ComputerName
    
    ForEach($addition in $DriveAdditions.amounttoaddgb)
    {

    $DriveAdditionSum += $addition

    }

    ForEach($result in $DriveAdditions.resultsizegb)
    {

    IF($result -gt 2047)
        {

        $errorCode = "Resulting drive will be larger than 2 TB."

        }

    }

RETURN $DriveAdditionSum,$errorcode
}

Function Get-DatastorePercent($ComputerName)
{
$Datastore = Get-datastore -vm $ComputerName
$DriveAdditionSum = Get-DriveSpaceAdditionSum $ComputerName | select -first 1

$DatastorePercentage = ((($datastore.FreeSpaceGB)-$DriveAdditionSum) / $datastore.CapacityGB)

$DatastoreHasSpace = $true

$FailureReason = $null

If($DatastorePercentage -le 0.20){
    $DatastoreHasSpace = $false
    $FailureReason = "Datastore lacks space."
    }    

IF($ComputerName -like "*MAILBOX*"){
    IF(!($failurereason)){
        $DatastoreHasSpace = $false
        $FailureReason = "Name contains Mailbox."
        
        }

        ELSE

        {
        $DatastoreHasSpace = $false
        $FailureReason += "`r`nName contains Mailbox."

        }

}

IF($ComputerName -like "*DFS*"){
    IF(!($failurereason)){
        $DatastoreHasSpace = $false
        $FailureReason = "Name contains DFS."
        
        }

        ELSE

        {
        $DatastoreHasSpace = $false
        $FailureReason += "`r`nName contains DFS."

        }

}
IF(Get-Snapshot -vm $ComputerName){
    IF(!($failurereason)){
        $DatastoreHasSpace = $false
        $FailureReason = "Snapshot exists."
        
        }

        ELSE

        {
        $DatastoreHasSpace = $false
        $FailureReason += "`r`nSnapshot exists."

        }

}

RETURN $DatastoreHasSpace,$failureReason
}

Function Get-DiskSpaceResult($ComputerName)
{

$AccessFailed = $false
IF(!($global:DefaultVIServer))
    {
    Write-Host -ForegroundColor Red "Not connected to a VSphere server"
    Initialize-Connection (Read-Host "Please enter the Vsphere server you wish to connect to.")

    }
    ELSE
    {
    $VISERVER = $global:DefaultVIServer.name
    Write-Host -ForegroundColor Green "Already connected to $VISERVER"

    }



Write-Host "Getting Disk Info for $ComputerName"
$Upgrade = Get-DiskSpaceUpgrade $ComputerName
#$upgrade
Write-Host "Getting VMware Info for $ComputerName"
$Vmware = Get-VMwareToDiskMatch $ComputerName
#$vmware
Write-Host "Getting Disk/Partition Info for $ComputerName"
$DiskPart = Get-DiskToPartitionMatch $ComputerName
#$diskpart

Write-Host "Getting DriveSpace/DataStore Info for $ComputerName"
$Percent = Get-DatastorePercent $ComputerName
$DriveAddSum =  Get-DriveSpaceAdditionSum $ComputerName

IF($DriveAddSum[1])
    {
    IF($percent[1])
        {

        $percent[1] = $driveaddSum[1] + "`r`n" + $percent[1]

        }
        ELSE
        {
        $percent[0] = $false
        $percent[1] = $driveaddSum[1]

        }


    }


#Write-Progress -Activity Updating -Status 'Calculating Results...' -PercentComplete ((4/5)*100) -CurrentOperation DiskSpaceResult
Write-Host "Calculating Results for $computername"

IF($DiskPart | where {$_.volumelabel -ne $null}){IF($upgrade){$Adds = Join-Object -Left $upgrade -LeftJoinProperty DriveLetter -right ($diskpart | where {$_.volumelabel -ne $null}) -RightJoinProperty VolumeLabel -Type AllInLeft }}
IF($DiskPart | where {$_.volumelabel -ne $null}){IF($upgrade){$FinalResult = Join-Object -Left $Adds -LeftJoinProperty PhysicalDrive -Right $Vmware -RightJoinProperty WindowsDeviceId -LeftProperties * -RightProperties VMwareDisk -Type AllInLeft}}

IF($FinalResult){$FinalResult = $FinalResult | select ServerName,@{l='TotalAmountToAdd';e={$DriveAddSum[0]}},DriveLetter,@{l='DiskNumber';e={(($_.diskandpart).substring(6))[-15..-15]}},@{l='Partition';e={[int](($_.diskandpart).substring(20))+1}},VMWareDisk,AmountToAddGB,ResultSizeGB,@{l="Datastore";e={(Get-datastore -vm $ComputerName).name}},@{l="GoCondition";e={$Percent[0]}},CurrentSize,@{l="noGoCode";e={$Percent[1]}} -Unique
}


RETURN $FinalResult | sort servername,driveletter
}

Function Add-DiskSpace($ComputerName)
{

Remove-StaleLogs $ComputerName

IF(!($global:DefaultVIServer))
    {
    Write-Host -ForegroundColor Red "Not connected to a VSphere server"
    Initialize-Connection (Read-Host "Please enter the Vsphere server you wish to connect to.")

    }
    ELSE
    {
    $VISERVER = $global:DefaultVIServer.name
    Write-Host -ForegroundColor Green "Already connected to $VISERVER"

    }

$drives = Get-DiskSpaceResult $ComputerName |where {$_.AmountToAddGB -ne '0'}
Write-Host -ForegroundColor Green "Displaying needed changes"
$drives | ft -autosize


IF($drives.GoCondition | select -first 1){

Write-Host -ForegroundColor Green "No issues detected."

foreach ($Drive in $drives)
    {
    $driveletter = $drive.DriveLetter

    
    
    
    
    Write-Host "Attempting to add space to $driveletter on $ComputerName in VMware..."
    Get-Harddisk $ComputerName | where {$_.name -eq $drive.VMwareDisk} | Set-HardDisk -CapacityGB $drive.ResultSizeGB -Confirm:$false

    # Construct diskpart script

    $disk = $drive.disknumber
    $partition = $drive.partition
    Write-Host "Attempting to extend disk $disk partition $partition to fill added space for $driveletter"
    Invoke-Command -ComputerName $ComputerName {param($disk,$partition)


        IF(!(Test-Path -Path "C:\temp"))
            {
            New-Item -Name temp -Path C: -ItemType Directory -force
            }
        New-Item -Name extend.txt -path c:\temp\ -ItemType file -Force

        Add-Content -Path c:\temp\extend.txt "rescan"
        Add-Content -Path c:\temp\extend.txt "select disk $disk"
        Add-Content -Path c:\temp\extend.txt "select partition $partition"
        Add-Content -Path c:\temp\extend.txt "extend"

        
        diskpart /s c:\temp\extend.txt 
        Remove-Item -path c:\temp\extend.txt

        } -ArgumentList $disk,$partition
    }

    $gutcheck = Get-DiskSpaceUpgrade $ComputerName | ft -AutoSize
    
    } ELSE {
    Write-Host -ForegroundColor Red $drives.NoGoCode
    $nogo = $drives.NoGoCode
    $gutcheck = Write-host -ForegroundColor Red "Script did not complete. $nogo"
    }
    
RETURN $gutcheck
}

FUNCTION Remove-StaleLogs($ComputerName)
{
# Get date minus 8 days
$limit = (Get-Date).AddDays(-8)

# Get date minus 20 days 
$limit2 = (Get-Date).AddDays(-20)

# Set Error Action Preference
$ErrorActionPreference = 'silentlycontinue'

#######################################
# Define Paths of Log files to delete #
#######################################

# Delete files from path
$path1name = 'Windows Temp Files'
$path1 = "\\$ComputerName\c$\Windows\Temp"

# Delete files from path2
$path2name = 'Default Web Site 1 Logs'
$path2 = "\\$ComputerName\c$\logs\LogFiles\W3svc1"

# Delete files from path3
$path3name = 'Exchange Pop3 Logs'
$path3 = "\\$ComputerName\c$\Program Files\Microsoft\Exchange Server\V14\Logging\Pop3"

# Delete files from path4
$path4name = 'IIS HTTPerr Logs'
$path4 = "\\$ComputerName\c$\Windows\System32\LogFiles\Httperr"

# Delete files from path5
$path5name = 'Lucas Session Logs'
$path5 = "\\$ComputerName\c$\Lucas_Log\SessionLogs"

# Delete files from path6
$path6name = 'IIS Logs for Default WebSite'
$path6 = "\\$ComputerName\c$\inetpub\logs\LogFiles\W3SVC1"

# Delete files from path7
$path7name = 'Windows Internal Database Logs'
$path7 = "\\$ComputerName\c$\Windows\Wid\Log"


$paths = @{}
$paths.add("$path1name","$path1")
$paths.add("$path2name","$path2")
$paths.add("$path3name","$path3")
$paths.add("$path4name","$path4")
$paths.add("$path5name","$path5")
$paths.add("$path6name","$path6")
$paths.add("$path7name","$path7")

####################################
# Clean out logs older than 8 days #
####################################

$resultant = @()

$before = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='C:'" 

$beforesize =$before.size

foreach($path in $paths.keys){
    
    $completed = $false
    Write-Host "Attempting to remove $path..."

    # Test if path exists and delete files older than the $limit for $path.

    $beforefreespace = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='C:'" | select -ExpandProperty freespace
	If (Test-Path -path $paths.$path) {
    
    Get-ChildItem -Path $paths.$path -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force;

    # Delete any empty directories left behind after deleting the old files for path.

    Get-ChildItem -Path $paths.$path -Recurse -Force | Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } | Remove-Item -Force -Recurse
    $completed = 'True'

    
    }

# If path doesnt exist log path clean up is skipped

    Else {$completed = 'DNE'}

    $afterfreespace = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='C:'" | select -ExpandProperty freespace

# Create Object for logging
    $FreespaceChange = (($beforefreespace - $afterfreespace)/1024/1024/-1024)
    $FreespacePercentageChange = [math]::Round((($beforefreespace - $afterfreespace)/$beforesize)*100,2)
    $TempObject1 = New-Object -TypeName PSObject -Prop (@{
                                                    'DateOfRun'=(Get-Date);
                                                    'TargetServer'=$ComputerName;
                                                    'FreeSpacePercentageChange'=$FreespacePercentageChange;
                                                    'FreeSpaceChange'=$FreespaceChange;
                                                    'Capacity'=[math]::Round((($beforesize)/1024/1024/1024));
                                                    'PathName'=$path;
                                                    'Completed?'=$completed
                                                    }) 
    #$completed
   
    $resultant += ($tempobject1 | select DateOfRun,TargetServer,FreeSpacePercentageChange,FreeSpaceChange,Capacity,PathName,Completed?)
}
$resultant = $resultant | ft -autosize
RETURN $resultant
}

