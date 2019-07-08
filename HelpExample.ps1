<# 
.Synopsis 
Generates a list of installed programs on a computer 
 
.DESCRIPTION 
This function generates a list by querying the registry and returning the installed programs of a local or remote computer. 
 
.NOTES    
Name       : Get-RemoteProgram 
Author     : Jaap Brasser 
Version    : 1.4.1 
DateCreated: 2013-08-23 
DateUpdated: 2018-04-09 
Blog       : https://www.jaapbrasser.com 
 
.LINK 
https://www.jaapbrasser.com 
 
.PARAMETER ComputerName 
The computer to which connectivity will be checked 
 
.PARAMETER Property 
Additional values to be loaded from the registry. Can contain a string or an array of string that will be attempted to retrieve from the registry for each program entry 
 
.PARAMETER IncludeProgram 
This will include the Programs matching that are specified as argument in this parameter. Wildcards are allowed. Both Include- and ExcludeProgram can be specified, where IncludeProgram will be matched first 
 
.PARAMETER ExcludeProgram 
This will exclude the Programs matching that are specified as argument in this parameter. Wildcards are allowed. Both Include- and ExcludeProgram can be specified, where IncludeProgram will be matched first 
 
.PARAMETER ProgramRegExMatch 
This parameter will change the default behaviour of IncludeProgram and ExcludeProgram from -like operator to -match operator. This allows for more complex matching if required. 
 
.PARAMETER LastAccessTime 
Estimates the last time the program was executed by looking in the installation folder, if it exists, and retrieves the most recent LastAccessTime attribute of any .exe in that folder. This increases execution time of this script as it requires (remotely) querying the file system to retrieve this information. 
 
.PARAMETER ExcludeSimilar 
This will filter out similar programnames, the default value is to filter on the first 3 words in a program name. If a program only consists of less words it is excluded and it will not be filtered. For example if you Visual Studio 2015 installed it will list all the components individually, using -ExcludeSimilar will only display the first entry. 
 
.PARAMETER SimilarWord 
This parameter only works when ExcludeSimilar is specified, it changes the default of first 3 words to any desired value. 
 
.EXAMPLE 
Get-RemoteProgram 
 
Description: 
Will generate a list of installed programs on local machine 
 
.EXAMPLE 
Get-RemoteProgram -ComputerName server01,server02 
 
Description: 
Will generate a list of installed programs on server01 and server02 
 
.EXAMPLE 
Get-RemoteProgram -ComputerName Server01 -Property DisplayVersion,VersionMajor 
 
Description: 
Will gather the list of programs from Server01 and attempts to retrieve the displayversion and versionmajor subkeys from the registry for each installed program 
 
.EXAMPLE 
'server01','server02' | Get-RemoteProgram -Property Uninstallstring 
 
Description 
Will retrieve the installed programs on server01/02 that are passed on to the function through the pipeline and also retrieves the uninstall string for each program 
 
.EXAMPLE 
'server01','server02' | Get-RemoteProgram -Property Uninstallstring -ExcludeSimilar -SimilarWord 4 
 
Description 
Will retrieve the installed programs on server01/02 that are passed on to the function through the pipeline and also retrieves the uninstall string for each program. Will only display a single entry of a program of which the first four words are identical. 
 
.EXAMPLE 
Get-RemoteProgram -Property installdate,uninstallstring,installlocation -LastAccessTime | Where-Object {$_.installlocation} 
 
Description 
Will gather the list of programs from Server01 and retrieves the InstallDate,UninstallString and InstallLocation properties. Then filters out all products that do not have a installlocation set and displays the LastAccessTime when it can be resolved. 
 
.EXAMPLE 
Get-RemoteProgram -Property installdate -IncludeProgram *office* 
 
Description 
Will retrieve the InstallDate of all components that match the wildcard pattern of *office* 
 
.EXAMPLE 
Get-RemoteProgram -Property installdate -IncludeProgram 'Microsoft Office Access','Microsoft SQL Server 2014' 
 
Description 
Will retrieve the InstallDate of all components that exactly match Microsoft Office Access & Microsoft SQL Server 2014 
 
.EXAMPLE 
Get-RemoteProgram -IncludeProgram ^Office -ProgramRegExMatch 
 
Description 
Will retrieve the InstallDate of all components that match the regex pattern of ^Office.*, which means any ProgramName starting with the word Office 
#>