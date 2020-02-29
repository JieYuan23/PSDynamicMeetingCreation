
$graphUrl = 'https://graph.microsoft.com'
$tenantId = "0a17712b-6df3-425d-808e-309df28a5eeb"

# Add required assemblies
#Function GetToken
#{
	# Azure AD OAuth User Token for Graph API
	# Get OAuth token for a AAD User (returned as $token)
	Add-Type -AssemblyName System.Web, PresentationFramework, PresentationCore
	# Application (client) ID, tenant ID and redirect URI
	$clientId = "e1b7697f-93df-4607-9f30-d854e0844f88"
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

$csvOut = "C:\temp\meetingsoutput.csv"
if (Test-Path $csvOut) {
	Remove-Item $csvOut
	}

	
$now = Get-Date -Format "MM_dd_yyyy_HH_mm"
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
							"displayname":"Adele Vance",
							"id":"c134eca6-67d6-4bec-a4f3-21a1fd0fd4b8"
						}
					},
					"upn":"AdeleV@M365EDU702432.OnMicrosoft.com"
				}
				'
				
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
	$meetingCode = $response.chatInfo.threadId.Replace(":meeting","_meeting")

    $meetingOptionUrl = "https://teams.microsoft.com/meetingOptions/?organizerId=" + $meetingOrganizer + "&tenantId=" + $tenantId + "&threadId=" + $meetingCode + "&messageId=0&language=en-US"
	$joinWebUrl= "https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri="+$response.joinWebUrl
	# c_corso,subject,meeting id,join link,meeting options url
	$meetingOutput = $code + "," + $subject + "," + $response.id + "," + $joinWebUrl + "," + $meetingOptionUrl
	

	# Write output
	Add-content $csvOut -value $meetingOutput
}
