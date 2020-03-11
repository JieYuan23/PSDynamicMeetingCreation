
$graphUrl = 'https://graph.microsoft.com'
$tenantId = "42d4a46d-9bc5-454b-821c-b1610ac5de9b"

# Add required assemblies
#Function GetToken
#{
	# Azure AD OAuth User Token for Graph API
	# Get OAuth token for a AAD User (returned as $token)
	Add-Type -AssemblyName System.Web, PresentationFramework, PresentationCore
	# Application (client) ID, tenant ID and redirect URI
	$clientId = "0ba51b74-369c-4494-8d48-0b4041a96f0c"
	$clientSecret = 'H8I=A4We5:zKkXm0GslbH@YprlT-?eVY'
	$redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

	# Scope - Needs to include all permisions required separated with a space
	$scope = "User.Read.All Group.Read.All" # This is just an example set of permissions

	# Random State - state is included in response, if you want to verify response is valid
	$state = Get-Random

	# Encode scope to fit inside query string 
	$scopeEncoded = [System.Web.HttpUtility]::UrlEncode($scope)

	# Redirect URI (encode it to fit inside query string)
	$redirectUriEncoded = [System.Web.HttpUtility]::UrlEncode($redirectUri)

	# Construct URI
	$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize?client_id=$clientId&response_type=code&redirect_uri=$redirectUriEncoded&response_mode=query&scope=$scopeEncoded&state=$state"

	# Create Window for User Sign-In
	$windowProperty = @{
	    Width  = 500
	    Height = 700
	}

	$signInWindow = New-Object System.Windows.Window -Property $windowProperty

	# Create WebBrowser for Window
	$browserProperty = @{
	    Width  = 480
	    Height = 680
	}

	$signInBrowser = New-Object System.Windows.Controls.WebBrowser -Property $browserProperty

	# Navigate Browser to sign-in page
	$signInBrowser.navigate($uri)

	# Create a condition to check after each page load
	$pageLoaded = {

	    # Once a URL contains "code=*", close the Window
	    if ($signInBrowser.Source -match "code=[^&]*") {

		# With the form closed and complete with the code, parse the query string

		$urlQueryString = [System.Uri]($signInBrowser.Source).Query
		$script:urlQueryValues = [System.Web.HttpUtility]::ParseQueryString($urlQueryString)

		$signInWindow.Close()

	    }
	}

	# Add condition to document completed
	$signInBrowser.Add_LoadCompleted($pageLoaded)

	# Show Window
	$signInWindow.AddChild($signInBrowser)
	$signInWindow.ShowDialog()

	# Extract code from query string
	$authCode = $script:urlQueryValues.GetValues(($script:urlQueryValues.keys | Where-Object { $_ -eq "code" }))
	
	
	
	    # With Auth Code, start getting token

	    # Construct URI
	    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

	    # Construct Body
	    $body = @{
		client_id    = $clientId
		scope        = $scope
		code         = $authCode[0]
		redirect_uri = $redirectUri
		grant_type   = "authorization_code"
			client_secret = $clientSecret;
	    }

	    # Get OAuth 2.0 Token
	    $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body

		# Access Token
		$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
	    #return ($tokenRequest.Content | ConvertFrom-Json).access_token

#}

# Code starts here _____________________________________________
#$token = GetToken
$now = Get-Date -Format "MM_dd_yyyy_HH_mm"

$csvOut = "C:\temp\meetingsoutput_" + $now + ".csv"
if (Test-Path $csvOut) {
	Remove-Item $csvOut
	}
	
$logs = "C:\temp\meetingslogs_" + $now + ".txt"

$string = "Token: " + $token
Add-content $logs -value $string

# Get input CSV and scan line by line
$csv = Get-Content "C:\temp\InputCorsi.csv"
$counter = 1
foreach ($line in $csv)
{
	$lineValues = $line.Split(';').Trim()
	$string =  "#" + $counter++ + " | " + $line
	Add-content $logs -value $string

	$code = $lineValues[0]
	$subject = $lineValues[1]
	$listAttendees = $lineValues[2..($lineValues.Length-1)]
	
	# Create meeting and output generation: c_corso,subject,meeting id,join link,meeting options url

	$staticUserId = 'c134eca6-67d6-4bec-a4f3-21a1fd0fd4b8'
	$staticUserDispName = 'Adele Vance'
	$staticUserUpn = 'AdeleV@M365EDU702432.OnMicrosoft.com'
	$meetingbody = '
    {
		"startDateTime":"2020-03-01T14:30:34.2444915-07:00",
		"endDateTime":"2020-06-30T16:00:34.2464912-07:00",
        "subject":"' + $subject + '",
        "participants": {
            "attendees": [
				{
					"identity":{
						"user":{
							"displayname":"' + $staticUserDispName + '",
							"id":"' + $staticUserId + '"
						}
					},
					"upn":"' + $staticUserUpn + '"
				}
				'
				
	$presentersPayload = '{"objectId":"'+$staticUserId+'","mri":"8:orgid:'+$staticUserId+'","upn":"'+$staticUserUpn+'"}'
    foreach ($attendee in $listAttendees) {

        $getUserQueryUrl = $graphUrl + "/v1.0/users/$attendee"
		$user = Invoke-WebRequest -Method Get -Uri $getUserQueryUrl -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} | ConvertFrom-Json

        $meetingbody += '     
			,{
				"identity":
				{
					"user":
					{
						"displayName": "'+$user.displayName+'",
						"id": "'+$user.id+'"
					}
				},
				"upn": "'+$user.userPrincipalName+'"
			}
		'

        if (-not ([string]::IsNullOrEmpty($presentersPayload)))
        {
            $presentersPayload += ","
		}
		
		$presentersPayload += '{"objectId":"'+$user.id+'","mri":"8:orgid:'+$user.id+'","upn":"'+$user.userPrincipalName+'"}'
	}
	
	$meetingbody+='
		]
		}
	}'

	$string = "Body:" + $meetingbody
	Add-content $logs -value $string
	
	# Create meeting Invok
	$meetingQueryUrl = $graphUrl + "/v1.0/me/onlineMeetings"
	$response = Invoke-RestMethod -Method Post -Uri $meetingQueryUrl -ContentType 'application/json' -Headers @{Authorization = "Bearer $token"} -Body $meetingbody

	$string = "Response" + $response
	Add-content $logs -value $string

	$meetingOrganizer = $response.participants.organizer.identity.user.id
	$threadId = $response.chatInfo.threadId.Replace(":meeting","_meeting")

    $meetingOptionUrl = "https://teams.microsoft.com/meetingOptions/?organizerId=" + $meetingOrganizer + "&tenantId=" + $tenantId + "&threadId=" + $threadId + "&messageId=0&language=en-US"
	$joinWebUrl= $response.joinWebUrl
	# c_corso,subject,meeting id,join link,meeting options url
	$meetingOutput = $code + "," + $subject + "," + $response.id + "," + $joinWebUrl + "," + $meetingOptionUrl
	
	# Write output
	Add-content $csvOut -value $meetingOutput
	
	# Test: change meeting options
	$apiTeamsAuth = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IkhsQzBSMTJza3hOWjFXUXdtak9GXzZ0X3RERSIsImtpZCI6IkhsQzBSMTJza3hOWjFXUXdtak9GXzZ0X3RERSJ9.eyJhdWQiOiJodHRwczovL2FwaS5zcGFjZXMuc2t5cGUuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvNDJkNGE0NmQtOWJjNS00NTRiLTgyMWMtYjE2MTBhYzVkZTliLyIsImlhdCI6MTU4Mjk2NzUzOCwibmJmIjoxNTgyOTY3NTM4LCJleHAiOjE1ODI5NzE0MzgsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiI0Mk5nWUFnUGNINVV4QmRVbjY3NCtQVHBqb0k1d3NrZVRNV21WcW5OdWZ1WkxlWW1hQUlBIiwiYW1yIjpbInB3ZCJdLCJhcHBpZCI6IjVlM2NlNmMwLTJiMWYtNDI4NS04ZDRiLTc1ZWU3ODc4NzM0NiIsImFwcGlkYWNyIjoiMCIsImF1dGhfdGltZSI6MTU4MjkwNTA5OSwiZmFtaWx5X25hbWUiOiJBZG1pbmlzdHJhdG9yIiwiZ2l2ZW5fbmFtZSI6IlN5c3RlbSIsImlwYWRkciI6IjkzLjQzLjIwMi4xMiIsIm5hbWUiOiJTeXN0ZW0gQWRtaW5pc3RyYXRvciIsIm9pZCI6IjQ0NTMxNDkyLTA3MTgtNDAyNC04YjZjLWMxNGZkM2E5ZTczYyIsInB1aWQiOiIxMDAzMjAwMDRENkFFN0Y2Iiwic2NwIjoidXNlcl9pbXBlcnNvbmF0aW9uIiwic3ViIjoiWjhIZEZKdTRMc0syUjVjaFNvZzVhVUJ6TGZwb3hBZlNuOTdOXy1ob1l3cyIsInRpZCI6IjQyZDRhNDZkLTliYzUtNDU0Yi04MjFjLWIxNjEwYWM1ZGU5YiIsInVuaXF1ZV9uYW1lIjoiYWRtaW5ATTM2NUVEVTcwMjQzMi5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBNMzY1RURVNzAyNDMyLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6InNDOHNkM0x4WkVxZFdXd0ZraVYzQUEiLCJ2ZXIiOiIxLjAifQ.ZjHMfIgtSbEUELNgff-_JF3vNvof51ggJvVF7etDH0VJEs3POtPpU5xYHuBiXtmXBurYAhsVPO_KPJALq08po8JjSj1JQIErZlVik1y7V7_qeE09VGML5DhJEFs5X8kd3GrxYfWU8QIZx_rvb2kLpCdPZM344kZjssJ4ZfybsOnwvVsgkY-3QbEWLCzZryNUOw3z-_WrOuzNelseYpi9OSPxID7-bQ8qphEMV_M7LrvAL3yJzhCPhMUGGhjyX8a0hQRIlusw32ccfyVTGd0y5-_qet9ndo5WMye-Ff-eUr2XLeH3KuMFJgcXV65c8qIwOfYUf3fhmfSthakPIa1v3A"
	$apiTeamsEndPoint = "https://teams.microsoft.com/api/mt/emea/beta/meetings/v1/options/42d4a46d-9bc5-454b-821c-b1610ac5de9b/44531492-0718-4024-8b6c-c14fd3a9e73c/$threadId/0/"
    $apiTeamsheaders = @{
        Authorization = ($apiTeamsAuth)
        Accept= "*/*"
	}
	
    $attendeeSelectionPayload = '{"options":[{"name":"AutoAdmittedUsers","currentValue":"EveryoneInCompany","type":"List"},{"name":"PresenterOption","currentValue":"SpecifiedPeople","type":"List"},{"name":"SpecifiedPresenters","type":"PresenterSelection","selectedPeople":[' + $presentersPayload + ']}]}'

	$string = "Presenters selection payload:" + $attendeeSelectionPayload
	Add-content $logs -value $string

    $attendeelistSetting = Invoke-WebRequest -Uri $apiTeamsEndPoint -Method Post -Headers $apiTeamsheaders -Body $attendeeSelectionPayload -ContentType "application/json;charset=UTF-8"
	
	$string = "Response change presenters" + $attendeelistSetting
	Add-content $logs -value $string
}
