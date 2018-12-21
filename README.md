# Add-Diskspace
Script to  expand a Windows VM's drive if below 30%

The purpose of this script is to make the process of expanding a virtual machine's disk drive more efficient.

It introduces 10 cmdlets/functions, 5 of which are intermediary and 5 that are intended for actual use.

1. Initialize-Connection <ComputerName> : This function loads the PowerCLI environment and connects to the VSphere server indicated by <ComputerName>.
  
2. Remove-Stalelogs <ComputerName> : This function removes logs on <Computername> that are older than 8 days from the below common log repositories. It then shows a report in the console of the change in space achieved.
  
    a. Windows Temp Files.
  
    b. Default Web Site 1 Logs.
  
    c. Exchange POP3 Logs.
  
    d. IIS HTTPerr Logs.
  
    e. Lucasware Logs.
  
    f. IIS logs for default website.
  
    g. WID logs.
  
3. Get-DiskSpaceUpgrade <ComputerName> : This function queries the server indicated in <Computername> via WMI for freespace and capacity of attached drives and, if there is less than 30% freespace, calculates the amount of space required to bring the drive to 35% freespace.
  
4. Get-DiskSpaceResult <ComputerName> : This function does the same thing as Get-DiskSpaceUpgrade except it queries the Vsphere server as well to correlate the drives/partitions to specific virtual disk files in Vsphere. It will also determine if the server should not be automatically expanded, given certain conditions.
  
5. Add-Diskspace <ComputerName> : This function uses every other cmdlet in the following order:
  
    a. If not connected to a Vsphere server, executes Initialize-connection.
  
    b. Execute Remove-StaleLogs.
  
    c. Execute Get-DiskSpaceResult and display needed changes, if any.
  
    d. For each of the drives that need additional space, expand the vmdk in VSphere. Then construct a diskpart bat script to extend the drive, move it to the remote computer, and then execute it.
  
    e. After all drives are completed, re-executes Get-DiskSpaceResult to validate the changes and displays the result in console.
  
    f. Failure reasons include:
  
      i. Resulting drive would exceed 2 TB.
    
      ii. Hostname contains "DFS" or "Mailbox" in the name.
    
      iii. There is an existing snapshot.
    
      iv. Adding the requested amount to the server would result in the datastore being brought below 20% available space.
