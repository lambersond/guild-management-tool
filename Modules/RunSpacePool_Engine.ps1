<########################################################################################
Title: RunSpacePool Assistant
Arthur(s): David Lamberson & Jeffrey Wills
Created: 22 August 2016
Last Edited: 23 August 2016
Company(s): Phantom Eagle & Jacobs
Version: 1.0
#########################################################################################
Description: Provides functions to create the RunsSacePool, add jobs/threads to the Pool, and
             Remove completed jobs from the Pool.  Also, provides three Syncronized hash tables
             That can be accessed from all runspaces.
               
            HASHTABLES:
            RunspacemanagementHash stores all Jobs/threads as well as the .Iterator,
            .RunPoolActive, .Script, .PoolCleaner.
            .Jobx: is a Powershell Object type with three properties
                    Powershell: stores the command thread passed to the runspace
                    Runspace:   stores provides the handle to close the job and capture 
                                the runspace output if specified
                    Name:       Name of the Job and will be used to create a hash key
                                with the runspace output if the -CaptureOutput switched is passed
                                to the Add-RPThread function
                    CaptureOut: boolean if set to true will trigger runspace output to be
                                stored.

            IOHash: provied for the user to move data between runspaces.  Any key name can be
                    created.  Typicaly these keys will be defined in the scriptblock that you
                    pass to the RunspacePool via the Add-Thread function
            
            FormHasH: Similar to IOhash but for used to logicaly organize and keep all form resources
                        in one Hash.
            
            FUNCTIONS:
            Open-RunspacPool:  Opens and cofigures the RunspacePool typicaly used once at start of script

            Add-RPThread: This is the high use function that will be used to add jobs/threads to the runspace
                          
            Run-PoolAutoClean: Automatically removes old jobs/threads from RunSpacePool and 
                               RunSpaceManagemtHash.  It will store the runspace output in the IOHash if the 
                               CaptureOut flag was set to true.  the IOHash key will be named after the jobs
                               Name.

            Clear-Runspace:  Used by the PoolCleaner to remove jobs from the runspace and RunspaceManagmentHash
                             Also, reads the CaptureOut flag and if set to true store the the runspace output to
                             the IOHash with key named after the $RunspceManagementHash.JobXX.Name
                             It can be call by user to Clear the Runspace if the Pool Cleaner is not running.
                             You will never use this function if the PoolCleaner is running 
               
Version  Date             User              Updates
1.0.0    22Aug2016     Wills/Lamberson      -Initial Release
1.0.1    24Aug2016     Wills                -Prevented Run-PoolAutoClean auto clean from running more than one
                                             PoolAutoClean routine 
                                            -Capitalized J in $Runspacemanagment.JobXX key 


TODO:
[x] import modules to Intial SessionState
[x] import functions to Intial SessionState
[ ] Add pipeline support Add-RPThread
[x] make a Hash for Jobs
[x] prevent job monitor from running multiple times


#>

$Global:rpoolscript = $MyInvocation.MyCommand.Definition

Function Open-RunspacePool
{
  <#

.SYNOPSIS
Opens and cofigures the RunspacePool and several Syncronized Hashtables to move data between runspaces

.DESCRIPTION
   Opens and cofigures the RunspacePool also creates the following hashtables
   Can be invoked without parameters.  The Default settings are as follows.
        sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault(),
        NumberOfRunspaces = 10
        ApartmentState = STA
        ThreadOptions = ReuseThread
        
   HASHTABLES:
            RunspacemanagementHash stores all Jobs/threads as well as the .Iterator,
            .RunPoolActive, .Script, .PoolCleaner.
            .Jobx: is a Powershell Object type with three properties
                    Powershell: stores the command thread passed to the runspace
                    Runspace:   stores provides the handle to close the job and capture 
                                the runspace output if specified
                    Name:       Name of the Job and will be used to create a hash key
                                with the runspace output if the -CaptureOutput switched is passed
                                to the Add-RPThread function
                    CaptureOut: boolean if set to true will trigger runspace output to be
                                stored.

            IOHash: provied for the user to move data between runspaces.  Any key name can be
                    created.  Typicaly these keys will be defined in the scriptblock that you
                    pass to the RunspacePool via the Add-Thread function
            
            FormHasH: Similar to IOhash but for used to logicaly organize and keep all form resources
                        in one Hash.

          The Following Hashkeys are set:
          $Global:RunSpaceManagementHash.RunPoolActive = $true           : flag for user to see the the Pool is open

          $Global:RunSpaceManagementHash.Iterator = 0 :                  : Used to number jobs and create the key name
                                                                           for the job

          $Global:RunSpaceManagementHash.Script = $Global:rpoolscript    : Stores the location of this PS1 file so that
                                                                           the pool cleaner has access to the Clear-Runspace
                                                                           function.
          

.PARAMETER SessionState
Represents a session state configuration that is used when a runspace is opened. The elements specified here, 
such as different types of commands, providers, and variables, are accessible each time a runspace that uses 
this configuration is opened. This class is introduced in Windows PowerShell 2.0.

.PARAMETER ThreadOptions 
In Windows PowerShell, each line—each command—is started in its own thread, or process. When ThreadOptions 
are set to Reuse Thread, each command is run in the same thread.  Reuse Thread improves the utilization of memory 
in Windows PowerShell and reduces the likelihood of memory leaks.

.PARAMETER AppartmentState
Can be set to STA, MTA or NA
STA is recommended
for more info on Appartments https://blogs.msdn.microsoft.com/cbrumme/2004/02/02/apartments-and-pumping-in-the-clr/

.PARAMETER NumberofRunspaces
Sets the maximum number of runspaces.  Default is 10

.PARAMETER AddFunctions
Adds Functions to the RunSpacePool IntialSessionState making them available to all runspaces that are spawned.

.PARAMETER AddModules
Adds Modules to the RunSpacePool IntialSessionState making them available to all runspaces that are spawned.

.EXAMPLE
Open-RunSpacePool -NumberOfRunspaces 5

.NOTES

.LINK


#>
  param(
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault(),
        $NumberOfRunspaces = 4,
        [string]$ApartmentState = "STA",
        $AddFunctions,
        $AddModules,
        [string]$ThreadOptions = "ReuseThread"
        )
  if($AddFunctions)
    {
       foreach($function in $AddFunctions)
        {
          $Definition = Get-Content Function:\$function
          $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "$function", $Definition
          $RunSpacePool.InitialSessionState.Commands.Add($SessionStateFunction)
        }
    }
  if($AddModule)
    {
      foreach($module in $AddModules) 
        {
          $sessionstate.ImportPSModule($Module) 
        }
    }
  Echo "Creating FormHash"
  if($Global:FormHash){Write-host FormHash Already Exists} else{$Global:FormHash = [Hashtable]::Synchronized(@{})}
  Echo "Creating IOHash"
  if($Global:IOHash){Write-host IOHash Already Exists} else{$Global:IOHash = [Hashtable]::Synchronized(@{})}
  if($Global:RPoolJobsHash){Write-host IOHash Already Exists} else{$Global:RPoolJobsHash = [Hashtable]::Synchronized(@{})}
  Echo "Creating and Configuring RunspaceManagementHash and RunSpacePool"
  if($Global:RunSpaceManagementHash.RunPoolActive -eq $false -or $Global:RunSpaceManagementHash.RunPoolActive -eq $null)
    {
      
      $Global:RunSpaceManagementHash = [Hashtable]::Synchronized(@{})
      $Global:RunSpaceManagementHash.RunPoolActive = $true
      $Global:RunSpaceManagementHash.Iterator = 0
      #$Global:RunSpaceManagementHash.Script = $Global:rpoolscript
      $Global:RunSpaceManagementHash.PoolCleaner = 'Not Started'
      $Global:RunSpaceManagementHash.MaximumRunspaces = $NumberOfRunspaces
      $Global:RunSpacePool = [runspacefactory]::CreateRunspacePool(1,$NumberOfRunspaces, $sessionstate, $Host)
      $Global:RunSpacePool.ApartmentState = $ApartmentState
      $Global:RunSpacePool.ThreadOptions = $ThreadOptions
      $Global:RunSpacePool.Open()
      Echo "RunspaceManagementHash and RunSpacePool Created"
    }
  Else{Echo "RunSpacePool Already Active"}
}



Function Clear-Runspace
{ 
<#

.SYNOPSIS
Used with the Open-RunspacePool command.  The Clear-Runspace command removes all completd jobs from the Runpool, 
and retrieves data if CaptureOut flag is set

.DESCRIPTION
Used with the Open-RunspacePool command.  The Clear-Runspace command removes all completd jobs from the Runpool, 
and retrieves data if CaptureOut flag is set

This routine checks job objects found in the $RunspacemanagmentHash.  If the job is completed it
issues an endInvoke() and the disposes of the powershell runspace Job/Thread.  If the CaptureOutput
Flag was set when creating the Job with Add-RPThread -CaptureOutput the data in the runspace will be
written to the IOHash with the key being the JobName .e.g $IOHash.Job1.  It is important to name the
Job if you want to CaptureOutPut, it will make it easier to find the data in the $IOHash


.EXAMPLE

.NOTES
The Open-RunspacePool Command must be run prior to turning on the PoolCleaner

.LINK


#>  
    Foreach($job in $RPoolJobsHash.keys) 
       {
           If ($RPoolJobsHash.("$job").Runspace.IsCompleted)
            {
                if($RPoolJobsHash.("$job").CaptureOut -eq $true)
                   {$IOHash.($($RPoolJobsHash.("$job").Name)) = $RPoolJobsHash.("$job").Powershell.EndInvoke($RPoolJobsHash.("$job").Runspace)}
                Else{$RPoolJobsHash.("$job").Powershell.EndInvoke($RPoolJobsHash.("$job").Runspace)}
                $RPoolJobsHash.("$job").Powershell.Dispose()
                $RPoolJobsHash.("$job").Runspace = $null
                $RPoolJobsHash.("$job").Powershell = $null
            }   
        }

    $temphash = $RPoolJobsHash.clone()
    foreach($i in $temphash.keys)
       {
         if($RPoolJobsHash.($i).Runspace -eq $null)
            {
                $RPoolJobsHash.Remove($i)
            }
       }
        $temphash = $null  
 } 
 


  function Add-RPThread
{
<#

.SYNOPSIS
Adds a Scriptblock or command to the RunspacePool

.DESCRIPTION
This function is  where all jobs are entered into the RunPool, it also specifies if the console data
from the runspace should be kept.  In addition to submitting the job to the runpool it also
modifies the $RunSpaceManagementHash to contain the information about the job.  While you can view
$RunSpaceManagmentHash this is primarily for the RunPoolengine to manage the RunPool.  If the Poolcleaner 
is Running and it is a short running job the Job entry in the $RunSpaceManagementHash maybe removed 
before you can view it.           

.PARAMETER ScriptBlock
A standard Scriptblock is the rimary method of putting commands/job into the RunPool.
 

Syntactically, a script block is a statement list in braces, as shown in 
the following syntax:

    $scriptblock = {echo hello world}

    {
        param ([type]$parameter1 [,[type]$parameter2])
        <statement list>
    }

for more information Get-Help about_Script_Blocks

The -Scriptblock parameter can also accept a command passed as a string beteween single quotes
 -ScriptBlock 'foreach($i in 1..10){$IOHash.key += $i}'


.PARAMETER JobName
This will be added to the $RunspaceManagementHash.jobXX PSObject Name Property.  This can be used
to see which exact jobs are running in the RunPool.  Also, if the -CaptureOut is specified this will be 
used for the Key name in the $IOHash that stores the runspace output. 

.PARAMETER AddArguments
Adds Arguments into the runspace.  It is how jobs in the runspace can be aware of variables outside the
runspace such as the $Global:IOHash Synchash.  If you have no param() section in your script block this
parameter will insert a param() section in your scriptblock with any variables you have supplied.  
AddArgument accepts comma delimted lists but all items must be strings. 

-AddArgument '$IOHASH',`$Formhash. 
 
Using enclosing single quotes around a string or a backtick ` infront of a special character will prevent 
the shell from expanding variables and force them to be interprted as a string.  Double quotes still allow
some expansion and therefore shouldn't be used.

.PARAMETER AddModules
Adds Modules to the runspace.  

-AddModules 'C:\Path\to\moudle\MyModule.psm1'

.PARAMETER AddFunctions
Adds custom functions from the primary runspace to runspace in the runspacepool.  It will only be available
to the runspace that the Job/Thread is run in.  If the same set off functions  and or modules are required
in all runspaces it is best to add them to the $intialSessionState of the RunspacePool
 
-AddFunctions 'Custom-Function','Custom-Funtion2'

.PARAMETER CaptureOut
A boolean flag set to $false by default.  If set to $true it will trigger the Clear-Runspace function to
capture the data that is output when the JobXX.Powershell.EndInvoke(Runspace) method is passed to the Runspace.
This is typicaly data that would be written to the console but can't be seen during execution as the runspace
can't write to the host.

The best method for moving data in and out of the runspaces is to use the $IOHash, or any 
$YourHash = [Hashtable]::Synchronized(@{}).  Just be sure to add it to the AddArguments and the
scriptblock params if you have a params section.  Otherwise leave it out the params section of the scriptblock
and the AddArguments will add it to your scriptblock at runtime 


.EXAMPLE
Add-RPThread -scriptBlock $runClean -AddArguments '$RunSpaceManagementHash','$IOHash' -JobName PoolCleaner

.EXAMPLE
Add-RPThread -scriptBlock '$formhash.form1.showdialog()' -AddArguments `$FormHash -JobName Form1

.EXAMPLE
Add-RPThread -scriptBlock 'dir c:\' -JobName Get-ChildItem -CaptureOut
.NOTES

.LINK


#>
    param
       (
          $ScriptBlock,
          $JobName = $('Job' + $RunSpaceManagementHash.iterator),
          $AddFunctions,
          $AddModules,
          $RunPool = $Global:RunSpacePool,
          [switch]$CaptureOut,
          [array]$AddArguments
       )
    
    #Build Function String
    if($AddFunctions)
       {
         foreach( $i in $AddFunctions)
            {
               $Definition = Get-Content Function:\$i
               [string]$function = "Function $i `{$Definition`}"
               [array]$Functions += "`r`n$function"
            }
       }
    #Build Import-Module String
    if($AddModule)
       {
          foreach( $i in $AddModule)
          {
             [string]$Modules += "`r`nImport-Module $i" 
          }
       }
    

    #build Param string
    If(!($ScriptBlock.Tostring() -match [regex]"^param"))
    {
       [string]$paramblock = 'param('
       for($i = 0 ; $i -lt $AddArguments.Count ; $i++ )
       {$paramblock += "$($AddArguments[$i])" 
          if($AddArguments[$i] -ne $AddArguments[-1])
          {$paramblock += ','}
       }
       
       $paramblock += ")`r`n"
       #[string]$scriptString = $paramblock + $ScriptBlock.ToString()
    }
    
    [string]$scriptString = $paramblock + $Modules + $Functions + $ScriptBlock.ToString()
    

    $string = '[powershell]::Create()'
    $string += ".AddScript(`{$scriptString`})"
    foreach($i in $AddArguments)
    {$string += ".AddArgument($i)"}
    $sb = [scriptblock]::Create($string)
    $job = (&$sb)
    $job.RunspacePool = $RunPool
    
    if(!($RPoolJobsHash.($JobName)))
      {
        $RPoolJobsHash.$JobName = New-Object psobject @{
           Powershell =  $job
           Runspace = $job.BeginInvoke()
           Name = $JobName
           CaptureOut = $CaptureOut
           JobNumber = $RunSpaceManagementHash.Iterator}
        $RunSpaceManagementHash.iterator++
      }
    Else
      {
        $RPoolJobsHash.($JobName + $RunSpaceManagementHash.Iterator) = New-Object psobject @{
           Powershell =  $job
           Runspace = $job.BeginInvoke()
           Name = ($JobName + $RunSpaceManagementHash.Iterator)
           CaptureOut = $CaptureOut
           JobNumber = $RunSpaceManagementHash.Iterator}
        $RunSpaceManagementHash.iterator++
      }
    #return $string

}

Function Run-PoolAutoClean
{
<#

.SYNOPSIS
Used with the Open-RunspacePool command.  The PoolCleaner routine removes completd jobs from the Runpool, 
and retrieves data if CaptureOut flag is set

.DESCRIPTION
Routine that runs in is own runspace in the runspace pool.  If you have specified your RunSpacePool
to be a maximum of 10 and you enable the Run-PoolAutoClean, 1 of these 10 will be used for the Cleaner
Leaving you with 9 runspaces available in the Pool.

This routine checks job objects found in the $RunspacemanagmentHash.  If the job is completed it
issues an endInvoke() and the disposes of the powershell runspace Job/Thread.  If the CaptureOutput
Flag was set when creating the Job with Add-RPThread -CaptureOutput the data in the runspace will be
written to the IOHash with the key being the JobName .e.g $IOHash.Job1.  It is important to name the
Job if you want to CaptureOutPut, it will make it easier to find the data in the $IOHash

.PARAMETER Start
Starts the Autocleaner.  However it is not required and the PoolClear will start if command is 
passed with no parameters i.e Run-PoolAutoClean

.PARAMETER Pause
The PoolCleaner routine is still running however it will not Clean the RunSpacePool or move data out
until passed the -Resume param

.PARAMETER Resume
Will return the PoolCleaner to active cleaning from a Paused State

.PARAMETER Off
Exits the PoolAutclean Routine.  In Order to restart the PoolCleaner the Run-PoolAutoClean or
Run-PoolAutoClean -Start command will need to be passed.

.EXAMPLE

.NOTES
The Open-RunspacePool Command must be run prior to turning on the PoolCleaner

.LINK


#>    
    param
       (
        [switch]$Start,
        [switch]$Off,
        [switch]$Resume,
        [switch]$Pause
       )
       
            if($Off)
              {
                 $RunSpaceManagementHash.PoolCleaner = 'ShutDown'
              }
            Elseif($Pause){$RunSpaceManagementHash.PoolCleaner = 'Paused'}
            Elseif($Resume){$RunSpaceManagementHash.PoolCleaner = 'Running'}
            
            
            Elseif($RunSpaceManagementHash.PoolCleaner -ne 'Running'`
                   -and $RunSpaceManagementHash.PoolCleaner -ne 'Paused'`
                   -or $RunSpaceManagementHash.PoolCleaner -eq 'Shutdown'`
                   -or $RunSpaceManagementHash.PoolCleaner -eq 'Not Started'`
                   -or $Start)
               {
                   $RunSpaceManagementHash.PoolCleaner = 'Running'
                   $runClean =
                     {
                        $RunSpaceManagementHash.poolcleanON = $true
                        Do
                            {
                                if($RunSpaceManagementHash.PoolCleaner -eq 'Paused')
                                {Start-Sleep 1}
                                Else
                                {Clear-Runspace
                                Start-Sleep 1}
                            }
                        while($Global:RunSpaceManagementHash.PoolCleaner -ne 'Shutdown')
                      }
              
                   Add-RPThread -scriptBlock $runClean -AddArguments '$RunSpaceManagementHash','$RPoolJobsHash','$IOHash' -JobName PoolCleaner-Hidden  -AddFunctions 'Clear-Runspace'
               }
            Else{Echo "Pool Cleaner Already Started"}
          
  } 
  
  Function Close-RunSpacepool
  {
    <#

.SYNOPSIS
Used with the Open-RunspacePool command.  The Close-RunSpacepool command clears and closes the RunPool

.DESCRIPTION
After are RunSpacePool has been created with the 

.EXAMPLE
Close-RunSpacepool

.NOTES


.LINK


#>
    Run-PoolAutoClean -Off
    Clear-Runspace
    $RunSpacePool.Dispose()
    $RunSpacePool = $null
    $RunSpaceManagementHash.RunPoolActive = $false
  }            
     
     
 function Sort-RPoolJobs
{
$JobNumbers = $null
$bar = $null

foreach($i in $RPoolJobsHash.keys)
  {
     if($i -match 'Hidden'){continue}

     else{ foreach($j in $RPoolJobsHash.$i.jobnumber)
           {
             [array]$JobNumbers += [int]$j
           }
        }
  }

$JobNumbers = ($JobNumbers | Sort-Object) 


foreach($i in $JobNumbers)
{
   foreach($j in $RPoolJobsHash.keys){if($RPoolJobsHash.$j.JobNumber -eq $i) {[array]$bar += $RPoolJobsHash.$j.Name}}
}
$bar
}
 
 $MonitorJobs = { #Monitors non-hidden Jobs
 
 if($RunSpaceManagementHash.JobMonitor -ne 'On')
   {
     $RunSpaceManagementHash.JobMonitor = 'On'
     Do{
           [array]$Joblist = Sort-RPoolJobs
           if($JobList -ne $null){$RunSpaceManagementHash.JobsInProgress = $True}
           else{$RunSpaceManagementHash.JobsInProgress = $False}
           Start-Sleep -Milliseconds 181
     } 
     while($RunSpaceManagementHash.JobMonitor -eq 'On')
  }
}

function Monitor-RPoolJobs  #Monitors non Hidden Jobs
{
   param([switch]$On,[switch]$Off)

   if($Off){$RunSpaceManagementHash.JobMonitor = "Off"}
   elseif($On -or $true -AND $RPoolJobsHash.Keys -notcontains 'RPoolJobMonitor-Hidden' ){
      Add-RPThread -ScriptBlock $MonitorJobs -JobName RPoolJobMonitor-Hidden -AddArguments `$RPoolJobsHash,`$RunSpaceManagementHash -AddFunctions 'Sort-RPoolJobs'
      }
}           
                      