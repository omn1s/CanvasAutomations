[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Config = @{}

Function Log-ResponseObject {
	Param(
		[Parameter(Mandatory=$True,Position=0)]
		$ResponseObject
	)
	Out-File -FilePath $Config.LogFile -InputObject "Logging response object" -append
	Out-File -FilePath $Config.LogFile -InputObject $ResponseObject -append
}


Function Invoke-WebRequestWithRetry {
	Param(
		[Parameter(Mandatory=$True,Position=0)]
		$RequestParameters,
		
		[Parameter(Mandatory=$True,Position=1)]
		[Int]
		$MaximumRetries,
		
		[Parameter(Mandatory=$True,Position=2)]
		[Int]
		$SecondsWaitBetweenTries
	)
	
	$IsSuccessfulResponse = $False
	$RetryCount = -1
	while (!$IsSuccessfulResponse -and ($RetryCount -le $MaximumRetries))
	{
		If ($Request -ne $Null) {
			Clear-Variable $Request
		}
		
		$RetryCount += 1
		
		if ($RetryCount -ge 1) {
			Start-Sleep -Seconds $SecondsWaitBetweenTries
		}
		
		try
		{ 
			$Request = Invoke-WebRequest @RequestParameters
			
			Log-ResponseObject $Request
		}
		
		catch [System.Net.WebException] 
		{
			$LogParams = @{
				Level = "error"
				Message = "A WebException was caught during an HTTP request: $($_.Exception.Message)"
			}
			Log-CanvasAutomations @LogParams
			
			Log-ResponseObject $_.Exception
			
			If ($_.Exception.Response -ne $Null) 
			{
				Log-ResponseObject $_.Exception.Response
			}
		}
		
		if ($Request.StatusCode -eq 200) {
			$IsSuccessfulResponse = $True
		}
	}
	
	If ($IsSuccessfulResponse) 
	{
		return $Request
	}
	Else {
		$LogParams = @{
				Level = "fatal"
				Message = "The maximum number of retries for this request was exceeded. This indicates network failure or other systemic issue. Terminating automation session..."
		}
		Log-CanvasAutomations @LogParams
		Throw $LogParams.Message
	}
}

Function Get-LoglineTimeStamp {
	"$(Get-Date -Format `"MM.dd.yyyy HH.mm.ss`")"
}

Function Log-CanvasAutomations {
	Param(
		[Parameter(Mandatory=$True,Position=0)]
		[ValidateSet("debug", "info", "warning", "error", "fatal")]
		[String]
		$Level,
		
		[Parameter(Mandatory=$True,Position=1)]
		[String]
		$Message
	)
	
	"$(Get-LoglineTimeStamp) | $($Level.ToUpper()) | $($Message)" |
	Out-File $Config.LogFile -append
}

Function Configure-Application {
	$SessionConfiguration = @{}
	
	If (Test-Path "./Config.xml") {
		[xml]$ConfigXml = Get-Content "./Config.xml"
	} Else {
		Throw "Config.xml does not exist in the current directory."
	}
	
	$Headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
	$Content = "application/json"
	$Headers.add("Accept", $Content)
	$Token = ""
	
	If ($ConfigXml.Configuration.ApplicationLevel -eq "test") {
	
		$Token = Get-Content $ConfigXml.Configuration.TestKeyFile
		$SessionConfiguration.BaseUrl = $ConfigXml.Configuration.TestBaseUrl
		$SessionConfiguration.BatchTermsFile = $ConfigXml.Configuration.TestBatchTermsFile
		
	} ElseIf ($ConfigXml.Configuration.ApplicationLevel -eq "prod") {
	
		$Token = Get-Content $ConfigXml.Configuration.ProdKeyFile
		$SessionConfiguration.BaseUrl = $ConfigXml.Configuration.ProdBaseUrl
		$SessionConfiguration.BatchTermsFile = $ConfigXml.Configuration.ProdBatchTermsFile
		
	} Else {
		Throw "The application level provided was not test or prod. Valid configurations only exist for test or prod."
	}

	$Authorization = "Bearer $($Token)"
	$Headers.add("Authorization", $Authorization)
	$SessionConfiguration.Headers = $Headers
	$FilenameTimeStamp = Get-LoglineTimeStamp
	$SessionConfiguration.LogFile = "Canvas_Automations_$($FilenameTimeStamp).log"
	$SessionConfiguration.ApplicationLevel = $ConfigXml.Configuration.ApplicationLevel
	$SessionConfiguration.SISDataDirectory = $ConfigXml.Configuration.SISDataDirectory
	
	return $SessionConfiguration
}

Function Assert-HttpStatus ($StatusCode, $StatusDescription) {
	if ($StatusCode -ne 200) {
		$LogParams = @{
			Level = "error"
			Message = "HTTP request failed with status code $($StatusCode) $($StatusDescription)."
		}
		Log-CanvasAutomations @LogParams
		return $False
	} else {
		return $True
	}
}

Function Validate-SISData {
	$SISDataDirectory = $Config.SISDataDirectory
	
	If (Test-Path $SISDataDirectory) {
		$Files = "users.csv", "terms.csv", "courses.csv", "sections.csv", "enrollments.csv"
		$Files | ForEach-Object {
			$Path = "$($SISDataDirectory)$($_)"
			If ((Test-Path $Path) -ne $True) {
				$LogParams = @{
					Level = "fatal"
					Message = "File Path: `"$($Path)`" not accessible."
				}
				Log-CanvasAutomations @LogParams
				Throw $LogParams.Message
			}
		}
	} Else {
		$LogParams = @{
			Level = "fatal"
			Message = "SIS Data Directory: `"$($SISDataDirectory)`" not accessible."
		}
		Log-CanvasAutomations @LogParams
		Throw $LogParams.Message
	}
	
	$LogParams = @{
		Level = "info"
		Message = "SIS data file-level validation completed successfully."
	}
	Log-CanvasAutomations @LogParams
}

Function Copy-SISData {
	$SISDataDirectory = $Config.SISDataDirectory
	Validate-SISData
	$Files = "users.csv", "terms.csv", "courses.csv", "sections.csv", "enrollments.csv"
	$Files | ForEach-Object {
		If (Test-Path $_) {
			Remove-Item $_
		}
		Copy-Item "$($SISDataDirectory)$($_)" .
	}
}

Function Import-CanvasSISData {
	Param(
		[Parameter(Mandatory=$True,Position=0)]
		[ValidateScript({Test-Path "$($_).csv"})]
		[String]
		$ImportType,
		
		[Parameter(Mandatory=$False)]
		[ValidateScript({$ImportType -eq "enrollments"})]
		[Switch]
		$IsBatch,
		
		[Parameter(Mandatory=$False)]
		[ValidateScript({($ImportType -eq "enrollments") -and $IsBatch})]
		[Int]
		$TermId
	)
	
	$ImportUrl = $Config.BaseUrl + "sis_imports?extension=csv"
	$LogBatchIndication = "(non-batch)"
	
	if ($IsBatch) {
		if ($TermId -eq $Null) {
			$LogParams = @{
				Level = "fatal"
				Message = "-IsBatch was specified but a null termId was provided. Terminating..."
			}
			Log-CanvasAutomations @LogParams
			Throw "-IsBatch was specified but a null termId was provided."
		}
		
		if ($TermId -eq 1) {
			$LogParams = @{
				Level = "fatal"
				Message = "-IsBatch was specified but the termId matched the default term, this behavior is not allowed. Terminating..."
			}
			Log-CanvasAutomations @LogParams
			Throw "-IsBatch was specified but the termId matched the default term, this behavior is not allowed. Terminating..."
		}
		
		$LogBatchIndication = "(batch, term: $($TermId))"
		$ImportUrl = $ImportUrl + "&batch_mode=1&batch_mode_term_id=$($TermId)"
		$LogParams = @{
			Level = "info"
			Message = "Starting $($ImportType) batch import. batch_mode_term_id: $($TermId)"
		}
		Log-CanvasAutomations @LogParams
		
	} else {
		$LogParams = @{
			Level = "info"
			Message = "Starting $($ImportType) general import $($LogBatchIndication)"
		}
		Log-CanvasAutomations @LogParams
	}
	
	$ImportRequestParameters = @{
		Uri = $ImportUrl
		Headers = $Config.Headers
		InFile = "./$($ImportType).csv"
		ContentType = "text/csv"
		Method = "POST"
	}
	$ImportRequest = Invoke-WebRequestWithRetry -RequestParameters $ImportRequestParameters -MaximumRetries 5 -SecondsWaitBetweenTries 2
	
	$Message = "Requested $($ImportType) import $($LogBatchIndication) HTTP Request Status: $($ImportRequest.StatusCode) $($ImportRequest.StatusDescription)."
	$LogParams = @{
		Level = "info"
		Message = $Message
	}
	Log-CanvasAutomations @LogParams
	
	$ImportId = (ConvertFrom-Json $ImportRequest.Content).id
	$Message = "Monitoring $($ImportType) import $($LogBatchIndication). Import Id: $($ImportId)"
	$LogParams = @{
		Level = "info"
		Message = $Message
	}
	Log-CanvasAutomations @LogParams
	
	$updateCheckCount = 1
	while($True)
	{
		$ImportStatusUri = $Config.BaseUrl + "sis_imports/$($ImportId)"
		$ImportStatusRequestParamaters = @{
			Uri = $ImportStatusUri
			Headers = $Config.Headers
			Method = "GET"
		}
			
		$ImportStatusRequest = Invoke-WebRequestWithRetry -RequestParameters $ImportStatusRequestParamaters -MaximumRetries 2 -SecondsWaitBetweenTries 1
		
		$ImportStatus = ConvertFrom-Json $ImportStatusRequest.Content
		$IsFailureResult = $False
		$IsCompletedResult = $False
		$FailureStates = "aborted", "failed_with_messages", "failed", "restoring", "partially_restored", "restored"
		switch ($ImportStatus.workflow_state)
		{
			"initializing" {
				$Message = "$($ImportType) import $($LogBatchIndication) workflow initializing. Pausing execution and checking again."
				$LogParams = @{
					Level = "info"
					Message = $Message
				}
				Log-CanvasAutomations @LogParams
				
				$updateCheckCount += 1
				Start-Sleep 5
				break
			}
			
			"created" {
				$Message = "$($ImportType) import $($LogBatchIndication) workflow created. Pausing execution and checking again."
				$LogParams = @{
					Level = "info"
					Message = $Message
				}
				Log-CanvasAutomations @LogParams
				
				$updateCheckCount += 1
				Start-Sleep 10
				break
			}
			
			"importing"  {
				$Message = "$($ImportType) import $($LogBatchIndication) workflow in process. Pausing execution and checking again." 
				$LogParams = @{
					Level = "info"
					Message = $Message
				}
				Log-CanvasAutomations @LogParams
				$updateCheckCount += 1
				Start-Sleep 60
				break
			}
			
			"cleanup_batch" {
				$Message = "$($ImportType) import $($LogBatchIndication) workflow is in cleanup_batch mode. Pausing execution and checking again."
				$LogParams = @{
					Level = "info"
					Message = $Message
				}
				Log-CanvasAutomations @LogParams
				Start-Sleep 5
				break
			}
			
			"imported" {
				$Message = "$($ImportType) import $($LogBatchIndication) workflow is successful with no warnings. Exiting $($ImportType) import. Total checks made: $($updateCheckCount)"
				$LogParams = @{
					Level = "info"
					Message = $Message
				}
				Log-CanvasAutomations @LogParams
				$IsCompletedResult = $True
				break
			}
			
			"imported_with_messages" {
				$Message = "$($ImportType) import $($LogBatchIndication) workflow is successful but with warnings. Exiting $($ImportType) import. Total checks made: $($updateCheckCount)"
				$LogParams = @{
					Level = "warning"
					Message = $Message
				}
				Log-CanvasAutomations @LogParams
				$IsCompletedResult = $True
				break
			}			
			
			{ @($FailureStates) -contains $_ } {
				$Message = "Current $($ImportType) import $($LogBatchIndication) workflow state, $($ImportStatus.workflow_state), represents import failure. Terminating..."
				$LogParams = @{
					Level = "error"
					Message = $Message
				}
				Log-CanvasAutomations @LogParams
				$IsFailureResult = $True
				break
			}
			
			default {
				$Message = "Current $($ImportType) import $($LogBatchIndication) workflow state, $($ImportStatus.workflow_state), is an unexpected workflow state. Terminating..."
				$LogParams = @{
					Level = "fatal"
					Message = $Message
				}
				Log-CanvasAutomations @LogParams
				$IsFailureResult = $True
				break
			}
		}
		if ($IsFailureResult) {
			$Message = "The import, $($ImportId), resulted in a failure state ($($ImportStatus.workflow_state)). Investigate."
			$LogParams = @{
				Level = "error"
				Message = $Message
			}
			Log-CanvasAutomations @LogParams
			break
		}
		if ($IsCompletedResult) { break }
	}	
	return $ImportId
}

Function Get-CanvasImportStatus {
	Param(
		[Parameter(Mandatory=$True,Position=0)]
		[String]
		$ImportId
	)
	$ImportUrl = "$($Config.BaseUrl)sis_imports/$($ImportId)"

	$ImportStatusRequestParameters = @{
		Uri = $ImportUrl
		Headers = $Config.Headers
		Method = "GET"
	}
	Invoke-WebRequestWithRetry -RequestParameters $ImportStatusRequestParameters -MaximumRetries 2 -SecondsWaitBetweenTries 1
}

Function Get-CanvasImportsList {
	$importsUrl = $Config.BaseUrl + "sis_imports"
	$importsRequestParameters = @{
		Uri = $importsUrl
		Headers = $Config.Headers
		Method = "GET"
	}
	Invoke-WebRequestWithRetry -RequestParameters $ImportsRequestParameters -MaximumRetries 2 -SecondsWaitBetweenTries 1
}

Function Get-CanvasTerms {
	$BaseTermsUrl = $Config.BaseUrl + "terms"

	$PerPage = 10
	$LastPage = 12

	$EnrollmentTerms = @()
	
	For ($i = 1; $i -le $LastPage; $i++) {
		$TermsPageUrl = $BaseTermsUrl + "?page=$($i)&per_page=$($PerPage)"
		$TermsRequestParameters = @{
			Uri = $TermsPageUrl
			Headers = $Config.Headers
			Method = "GET"
		}
		$TermsPageRequest = Invoke-WebRequestWithRetry -RequestParameters $TermsRequestParameters -MaximumRetries 2 -SecondsWaitBetweenTries 1
		$TermsPageContent = ConvertFrom-Json $TermsPageRequest.Content
		$EnrollmentTerms += $TermsPageContent.enrollment_terms
	}
	
	$EnrollmentTerms
}

Function Start-CanvasGeneralImports {
	$GeneralImportIds = @()
	$GeneralImportIds += Import-CanvasSISData "users"
	$GeneralImportIds += Import-CanvasSISData "terms"
	$GeneralImportIds += Import-CanvasSISData "courses"
	$GeneralImportIds += Import-CanvasSISData "sections"
	$GeneralImportIds += Import-CanvasSISData "enrollments"
	$GeneralImportIds
}

Function Start-CanvasBatchImports {
	$BatchImportIds = @()
	$BatchTerms = Import-Csv $Config.BatchTermsFile
	$BatchTerms | ForEach-Object {
		$BatchImportIds += Import-CanvasSISData -ImportType enrollments -IsBatch -TermId $_.Id
	}
	$BatchImportIds
}


$Config = Configure-Application
$LogParams = @{
	Level = "Info"
	Message = "Automation Session Configured. Current ApplicationLevel: $($Config.ApplicationLevel)"
}
Log-CanvasAutomations @LogParams

Function Start-MyCanvasImports {
	$GeneralImportIds = Start-CanvasGeneralImports
	$BatchImportIds = Start-CanvasBatchImports
	
	$AllImportIds = $GeneralImportIds + $BatchImportIds
	
	[String[]]$EmailMessageBody = @()
	[String[]]$EmailAttachments = @()
	
	$MessageHeaderRow = "Id`tStarted At`t`t`tCompleted At`t`t`tWorkflowState"
	$MessageFancyLineRow = "==`t==========`t`t`t============`t`t`t============="
	
	$EmailMessageBody += $MessageHeaderRow
	$EmailMessageBody += $MessageFancyLineRow
	
	[String[]]$TempFilesToDelete = @()
	
	$AllImportIds | ForEach-Object {
		$ImportStatus = ConvertFrom-Json (Get-CanvasImportStatus $_).Content
		$MessageLine = "$($ImportStatus.id)`t$($ImportStatus.started_at)`t`t$($ImportStatus.ended_at)`t`t$($ImportStatus.workflow_state)"
		
		If ($ImportStatus.errors_attachment -ne $Null) {
			$ErrorsDownloadUrl =  $ImportStatus.errors_attachment.url
			$File = "$($ImportStatus.id)_errors.csv"
			$ErrorsDownloadRequestParameters = @{
				Uri = $ErrorsDownloadUrl
				OutFile = $File
			}
			Invoke-WebRequestWithRetry -RequestParameters $ErrorsDownloadRequestParameters -MaximumRetries 2 -SecondsWaitBetweenTries 1
			$EmailAttachments += $File
			$TempFilesToDelete += $File
		}
		
		$EmailMessageBody += $MessageLine	
	}
	
	$EmailAttachments += $Config.LogFile
	
	$MessageParams = @{
		From = "blank@blank.com"
		To = "blank@blank.com"
		Attachments = $EmailAttachments
		Body = $EmailMessageBody -join "`n"
		Subject = $Config.LogFile
		SmtpServer = "x.x.x.x"
	}
	
	Send-MailMessage @MessageParams
	
	$TempFilesToDelete | Remove-Item
}
