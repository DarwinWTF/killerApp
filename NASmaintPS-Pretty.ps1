# =================================================================
# Author: 
#     NASmaintPS by K.Hall 2023/08/01
#
# Summary Description:
#   Clean workbin files from NAS.
#   
#
# Dependancy list: 
#   Log4Net DLL v 4
#   AppConfig.xml defined and available
#   <folder structure> 
#    \temp
#    \logs
#    \data
#    \jobs
#    \archive
#
# History:
#   Version 1.0
#
# Notes: 
#   Module must be registered with the MS EVENT log to log events before the script will log to 
#     Event View\Application and Services Log\MDMS\<scriptname>
#     Open CMD as admin, navigate to run folder, type: Powershell.exe .\NASmaintPS.ps1 -RegisterModule
#         -RegisterModule switch needs to run under ADMIN credentials. 
#     Any logs or data will get created with elevated privs.
#
# Usage: Instructions:  
#   Run job via Task Scheduler with Runas CORP\sv-mdms & Highest privelges
#         Script "E:\apps\NASmaintPS\jobs\RunJob.bat "
#         Add Arguements "NASmaintPSjob"
#         start in root folder
#  
#    - "Debug" will create verbose logging when necassary
#    - "Info" is default 
#    -  Set "All"or"Debug" -LogLevel in AppConfig.xml
#    
# =================================================================

[CmdletBinding()]
param(
	[Parameter(Mandatory = $False)]
	[switch]$RegisterModule
)

#****************************************************************
# enforce strict code, flow, and variable behaviors 
#****************************************************************
Set-StrictMode -Version Latest
#Set-StrictMode -Off
$ErrorActionPreference = "Stop"
#****************************************************************
#  Note : (Error.Clear) will clear the exception buffers from prior ISE editor errors 
#         ISE likes to hold the error history ex: $Error{}
#****************************************************************
$Error.Clear()

#****************************************************************
# implement default startup code and Log4Net logger
#****************************************************************
try {
	#****************************************************************
	##### Define default Variables #####
	#****************************************************************
	$exitcode = 0
	$scriptName = $MyInvocation.MyCommand.Name
	$scriptID = [System.IO.Path]::GetFileNameWithoutExtension($scriptName)
	$date = (Get-Date).ToString('yyyyMMdd')
	$dateTime = (Get-Date).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
	$logStore = $PSScriptRoot + "\logs\"
	$tempStore = $PSScriptRoot + "\temp\"
	$dataStore = $PSScriptRoot + "\data\"
	$EventLogName = "MDMS"


	if ($RegisterModule) {
		# Note : -RegisterModule switch needs to run under ADMIN credentials. Any logs or data will get created with elevated privs.
		# It is important to not lock up the application or job logs under ADMIN privs. This section will be executed as an 
		# installer and create seperate ADMIN install logs.  

		try {
			$RegisterLog = "$logStore$scriptID" + "Install.log"
			if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator"))
			{
				$exitcode = -1
				(Get-Date).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") + " -RegisterModule needs to runas (administrator). Start an Administrator instance and rerun RegisterModule script" |
				Out-File -Append -FilePath $RegisterLog -Encoding ASCII
				exit $exitcode
			}

			$SourceExists = [System.Diagnostics.Eventlog]::SourceExists("$scriptID")
			(Get-Date).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") + " RegisterModule to Windows event log request : $EventLogName\$scriptID" |
			Out-File -Append -FilePath $RegisterLog -Encoding ASCII

			if ($SourceExists -eq $false) {
				[System.Diagnostics.EventLog]::CreateEventSource("$scriptID","MDMS")
				(Get-Date).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") + " New Event log created : $EventLogName\$scriptID" |
				Out-File -Append -FilePath $RegisterLog -Encoding ASCII
			}
			else {
				(Get-Date).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") + " NO action required - RegisterModule to Windows event log request is already set: $EventLogName\$scriptID" |
				Out-File -Append -FilePath $RegisterLog -Encoding ASCII
			}
			(Get-Date).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") + " RegisterModule with Windows event log completed : $EventLogName\$scriptID" |
			Out-File -Append -FilePath $RegisterLog -Encoding ASCII

			exit $exitcode
		}
		catch {

			"*Process Exception trapped on $scriptName at line number " + $_.InvocationInfo.ScriptLineNumber | Out-File -Append -FilePath $RegisterLog -Encoding ASCII
			"ScriptStackTrace  : " + $_.ScriptStackTrace | Out-File -Append -FilePath $RegisterLog -Encoding ASCII
			"Line              : " + $_.InvocationInfo.Line | Out-File -Append -FilePath $RegisterLog -Encoding ASCII
			"Stack Trace       : " + $_.Exception.StackTrace | Out-File -Append -FilePath $RegisterLog -Encoding ASCII
			"Exception         : " + $_.Exception | Out-File -Append -FilePath $RegisterLog -Encoding ASCII
			"Exception Message : " + $_.Exception.Message | Out-File -Append -FilePath $RegisterLog -Encoding ASCII

			exit $exitcode
		}
	}

	#***************************************************************
	# read app.config file
	#***************************************************************
	$AppConfigFilePath = $PSScriptRoot + "\AppConfig.xml"
	if (!(Test-Path -Path $AppConfigFilePath)) {
		throw "AppConfigFilePath $AppConfigFilePath is not valid. Please provide a valid path to AppConfig.xml"
	}
	[xml]$AppConfigData = Get-Content $AppConfigFilePath

	#****************************************************************
	# get log4net dll to start the logger.
	# determine input log level - commandline input will Supersede AppConfig setting if it was applied as an input parma
	#****************************************************************
	$Log4NetDllFile = ($AppConfigData.configuration.appSettings.Add | Where-Object { $_.key -eq "Log4NetDllFile" }).Value
	if ($null -eq $Log4NetDllFile) {
		throw "Log4NetDllPath is not defined"
	}
	if (!(Test-Path -Path $Log4NetDllFile)) {
		throw "Log4NetDllPath $Log4NetDllFile is not valid. Please provide a valid dll and path to AppConfig.xml"
	}

	$LogLevel = ($AppConfigData.configuration.appSettings.Add | Where-Object { $_.key -eq "LogLevel" }).Value
	$Environment = ($AppConfigData.configuration.appSettings.Add | Where-Object { $_.key -eq "Environment" }).Value
	$ReportName = ($AppConfigData.configuration.appSettings.Add | Where-Object { $_.key -eq "ReportName" }).Value

	#****************************************************************
	# scriptLog           :log4net_template.ps1.log
	# Log4NetLogFileName  :log4net_template.log
	# use command line debug to write to scriptLog to trap startup issues
	#****************************************************************
	$Log4NetLogFileName = $logStore + $scriptID + ".log"
	#****************************************************************

	if ($LogLevel -eq "Debug") {

		Write-Output "================================================"
		Write-Output "runtime : $dateTime"
		Write-Output "**********startup configurations *****************"
		Write-Output "scriptName         $scriptName"
		Write-Output "scriptID           $scriptID"
		Write-Output "date               $date"
		Write-Output "logStore           $logStore"
		Write-Output "tempStore          $tempStore"
		Write-Output "dataStore          $dataStore"
		Write-Output "LogLevel           $LogLevel"
		Write-Output "Log4NetDllFile     $Log4NetDllFile"
		Write-Output "Log4NetLogFileName $Log4NetLogFileName"

	}

	#****************************************************************
	# Log4Net RollingAppender
	#   Define Values for $RollingAppender and message formats 
	#****************************************************************
	$Log4NetPattern = '%date{yyyy-MM-dd HH:mm:ss.fff} [%level] %message%n'


	if ($LogLevel -eq "Debug") {

		Write-Output "$dateTime log4net configurations *****************"
		Write-Output "Log4NetLogFile $Log4NetLogFileName"
		Write-Output "Log4NetPattern $Log4NetPattern"
		Write-Output "Log4NetDllFile $Log4NetDllFile"
	}

	#****************************************************************
	# Load log4net.dll
	#****************************************************************
	[void][Reflection.Assembly]::LoadFile($Log4NetDllFile)

	#****************************************************************
	# set log4net log pattern
	#****************************************************************
	$Log4NetPatternLayout = [log4net.Layout.ILayout](New-Object log4net.Layout.PatternLayout ($Log4NetPattern))

	#****************************************************************
	# This is a useful alternate command syntax ..
	#   $Log4NetLogPath = ([System.IO.Directory]::GetParent($MyInvocation.MyCommand.Path)).FullName
	# This is useful alternative if NOT using rolling appender - create a new log every day
	#   $Log4NetLogFileName = Join-Path -Path $Log4NetLogPath -ChildPath $('LogFile_{0:yyyy-MM-dd}.log' -f (Get-Date))
	#****************************************************************
	# Simple Appender configuration
	#   Load FileAppender Configuration - it is better to use the rolling appender...but this is an option as well
	# $Log4NetAppendToFile = $True
	# $Log4NetFileAppender = new-object log4net.Appender.FileAppender($Log4NetPatternLayout,$Log4NetLogFile,$Log4NetAppendToFile);
	#****************************************************************

	#****************************************************************
	# RollingAppender configuration - this is the preferred logger method
	#****************************************************************
	$Log4NetRollingAppender = New-Object log4net.Appender.RollingFileAppender;

	#****************************************************************
	# this is useful syntax to discover types and members
	# $Log4NetRollingAppender.GetType()
	# $Log4NetRollingAppender | Get-Member
	#****************************************************************
	$Log4NetRollingAppender.Layout = $Log4NetPatternLayout;
	$Log4NetRollingAppender.File = $Log4NetLogFileName
	$Log4NetRollingAppender.StaticLogFileName = $True
	$Log4NetRollingAppender.AppendToFile = $True
	$Log4NetRollingAppender.RollingStyle = [log4net.Appender.RollingFileAppender+RollingMode]::Size
	$Log4NetRollingAppender.MaxSizeRollBackups = 10

	#****************************************************************
	#MaximumFileSize valid domain list "KB", "MB" or "GB"
	#****************************************************************
	$Log4NetRollingAppender.MaximumFileSize = "10MB"
	$Log4NetRollingAppender.PreserveLogFileNameExtension = $True

	#****************************************************************
	# note: this filters away all log messages that falls below your threshold - "All" is the best use case 
	# since it gets complicated between threshold and level settings
	#****************************************************************
	$Log4NetRollingAppender.Threshold = [log4net.Core.Level]::All
	[log4net.Config.BasicConfigurator]::Configure($Log4NetRollingAppender)
	$Log4NetRollingAppender.ActivateOptions()

	#****************************************************************
	# Note: rfa -> rolling file appender
	#****************************************************************
	$Log4NetRollingAppender.Name = "rfa"
	$Log4NetRollingAppender.ImmediateFlush = $True

	#****************************************************************
	# it is fine to use root...but the name is more precise
	#   $Log4NetLog = [log4net.LogManager]::GetLogger("root")
	#****************************************************************
	$Log4NetLog = [log4net.LogManager]::GetLogger($Log4NetRollingAppender.Name.ToString())
	#****************************************************************
	# note: level indicates what log statements will actually will be generated. It works in conjunction with threshold
	$Log4NetLog.Logger.Level = [log4net.Core.Level]::$LogLevel

	#****************************************************************
	#Write-Host "RollingAppender.Name " $Log4NetRollingAppender.Name.ToString()
	#****************************************************************
	$Log4NetLog.Debug(‘********** log4net level **********’)
	$Log4NetLog.Debug("IsInfoEnabled  " + $Log4NetLog.IsInfoEnabled.ToString())
	$Log4NetLog.Debug("IsWarnEnabled  " + $Log4NetLog.IsWarnEnabled.ToString())
	$Log4NetLog.Debug("IsErrorEnabled " + $Log4NetLog.IsErrorEnabled.ToString())
	$Log4NetLog.Debug("IsFatalEnabled " + $Log4NetLog.IsFatalEnabled.ToString())
	$Log4NetLog.Debug("IsDebugEnabled " + $Log4NetLog.IsDebugEnabled.ToString())

}
catch {

	Write-Output ("$dateTime Startup Exception trapped on $scriptName at line number " + $_.InvocationInfo.ScriptLineNumber)
	Write-Output $_.InvocationInfo.Line
	Write-Output $_.Exception.StackTrace
	Write-Output $_.Exception.Message

	$PSCmdlet.ThrowTerminatingError($PSItem)
}
finally {

}
#****************************************************************
# main() 
#****************************************************************

# Set the path to the configuration CSV file
$ConfigFilePath = "E:\apps\NASmaintPS\config2.csv"


try {
	$Log4NetLog.Info("Process start runtime : $dateTime")
	$Log4NetLog.Info("Powershell version:$($PSVersionTable.PSVersion)")

	# Read the configuration CSV file
	$Config = Import-Csv $ConfigFilePath
	$Log4NetLog.Info("Total Input Rows: " + $Config.Count)
	# Loop through each record in the configuration and perform the specified file operation
	foreach ($Record in $Config) {
		$Operation = $Record.OPERATION.ToUpper().Trim()
		$Description = $Record.DESCRIPTION
		$SourcePath = $Record.SOURCE.Trim()
		$Destination = $Record.DESTINATION.Trim()
		$NDays = $Record.NDAYS
		$Filter = $Record.FILTER.Trim()

		$Log4NetLog.Info("===================================================")
		$Log4NetLog.Info("$Record")
		$Log4NetLog.Info("===================================================")
		$Log4NetLog.Info("Operation:$Operation")
		$Log4NetLog.Info("Descriptoin:$Description")
		$Log4NetLog.Info("SourcePath:$SourcePath")
		$Log4NetLog.Info("Destination:$Destination")
		$Log4NetLog.Info("Ndays:$NDays")
		$Log4NetLog.Info("Filter:$Filter")

		# Perform the file operation based on the specified operation

		switch ($Operation) {
			"DELETE" {
				# Get files older than NDays that match the filter
				$Log4NetLog.Info('Starting DELETE Switch')
				$FilesToDelete = Get-ChildItem -Path $SourcePath -File -Recurse | Where-Object {
					$_.LastWriteTime -lt (Get-Date).AddDays(- $NDays) -and $_.Name -like $Filter
				}

				$Log4NetLog.Info("Total Files to Delete: " + ($FilesToDelete | Measure-Object).Count)
				if (($FilesToDelete | Measure-Object).Count -ne 0) {
					foreach ($File in $FilesToDelete) {
						# Write null to the file
						$nullStream = [System.IO.StreamWriter]::new($File.FullName)
						$nullStream.Write($null)
						$nullStream.Flush()
						$nullStream.Close()


						# Delete the file
						$Log4NetLog.Info("Deleting: $($File.FullName)")
						Remove-Item -Path $file.FullName -Force
						$Log4NetLog.Info("Deleted:  $($File.FullName)")
					}
				}

				else {
					$Log4NetLog.Info("No files found to delete in $SourcePath")
				}
			}
			"COPY" {
				# Do SecureMove for 'COPY' operation
				# Step 1: Get Source Files List and Copy files from source to destination
				$Log4NetLog.Info('Starting COPY Switch')
				if (-not (Test-Path $Destination -PathType Container)) {
					$Log4NetLog.Error("Destination Folder Doesn't Exist")
					return
				}
				$FilesToCopy = Get-ChildItem -Path $SourcePath -File -Recurse | Where-Object {
					$_.LastWriteTime -lt (Get-Date).AddDays(- $NDays) -and $_.Name -like $Filter
				}

				$FilesToCopy | ForEach-Object {


					$Log4NetLog.Info("Source.FullFile: $($_.FullName)")
					$Log4NetLog.Info("Destination:     $Destination")
					$Log4NetLog.Info("SourceFileOnly:  $($_.Name)")

					# Perform the copy operation
					Copy-Item -Path $_.FullName -Destination $Destination -Force
					$Log4NetLog.Info("Completed 'COPY' operation")

					# Calculate checksums for source and destination files
					$SourceChecksum = Get-FileHash -Path $_.FullName -Algorithm SHA256 | Select-Object -ExpandProperty Hash
					$DestinationChecksum = Get-FileHash -Path $Destination\$_ -Algorithm SHA256 | Select-Object -ExpandProperty Hash

					# Compare checksums
					if ($SourceChecksum -eq $DestinationChecksum) {
						$Log4NetLog.Info("Source and destination files MATCH. Proceeding to null write and delete.")
						# Continue with null write and delete operations
						# ...
					} else {
						$Log4NetLog.Error("Source and destination files do NOT Match. Aborting script.")
						# Handle the situation when files do not match, e.g., delete the copied file if necessary
						# ...
						exit
					}
				}


				# Step 3: Write null then Delete the source files
				if ($FilesToCopy.Count -gt 0) {
					foreach ($File in $FilesToCopy) {
						# Write null to the file
						$Log4NetLog.Info("Deleting Souce after NullWrite: $($File.FullName)")
						$nullStream = [System.IO.StreamWriter]::new($File.FullName)
						$nullStream.Write($null)
						$nullStream.Flush()
						$nullStream.Close()


						# Delete the file
						$Log4NetLog.Info("Deleting Souce after NullWrite: $($File.FullName)")
						Remove-Item -Path $file.FullName -Force
					}
				}


				$Log4NetLog.Info("Total Files Removed:" + $FilesToCopy.Count)
			}

			"NOP" {
				$Log4NetLog.Info('Starting NOP Switch')
				$Log4NetLog.Info("No action required for 'NOP' operation...Moving on")
			}

			Default {
				$Log4NetLog.Error("$Operation switch not defined")
				$Log4NetLog.Info("No action Completed for $Operation ...Moving on")
			}

		}

	}
}
catch {
	$exitcode = 1
	$Log4NetLog.Error("*Process Exception trapped on $scriptName at line number " + $_.InvocationInfo.ScriptLineNumber)
	$Log4NetLog.Error("ScriptStackTrace  : " + $_.ScriptStackTrace)
	$Log4NetLog.Error("Line              : " + $_.InvocationInfo.Line)
	$Log4NetLog.Error("Stack Trace       : " + $_.Exception.StackTrace)
	$Log4NetLog.Error("Exception         : " + $_.Exception)
	$Log4NetLog.Error("Exception Message : " + $_.Exception.Message)
}

finally {
	if ($exitcode -ge 0) {
		if ($exitcode -eq 0) {
			Write-EventLog –LogName $EventLogName –Source $scriptID –EntryType Information –EventID $exitcode –Message “Process Complete:Success.”
		}
		else {
			Write-EventLog –LogName $EventLogName –Source $scriptID –EntryType Error –EventID $exitcode –Message “Process failed with non-zero return code. System Administrator review required.”
		}
	}
	[log4net.LogManager]::ResetConfiguration()
	exit $exitcode
}
