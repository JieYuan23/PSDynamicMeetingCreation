# Azure AD OAuth User Token for Graph API
# Get OAuth token for a AAD User (returned as $token)

# Add required assemblies
Add-Type -AssemblyName System.Web, PresentationFramework, PresentationCore

Function GetToken{
	# Application (client) ID, tenant ID and redirect URI
	$clientId = "0ba51b74-369c-4494-8d48-0b4041a96f0c"
	$tenantId = "42d4a46d-9bc5-454b-821c-b1610ac5de9b"
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
	$graphUrl = 'https://graph.microsoft.com'

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
	
	
	if ($authCode) {

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
	    return $token
	}
	else {
	    Write-Error "Unable to obtain Auth Code!"
	}
}

Function CreateMeeting{
	param ($code, $subject, $listAttendees, $token)
	
	#build body
	$meetingparam = '
	{
		"startDateTime":"2020-03-01T14:30:34.2444915-07:00",
		"endDateTime":"2020-06-30T15:00:34.2464912-07:00",
		"subject":"PS Graph Meeting",
		"participants": {

			"attendees": [
				{
					"identity":{
						"user":{
							"displayname":"$name",
							"id":"$iduser"
						}
					},
					"upn":"$upn"
				}
			]
		}
	}'


	$queryUrl = $graphUrl + "/v1.0/me/onlineMeetings"
	
	$response = Invoke-RestMethod -Method Post -Uri $queryUrl -ContentType 'application/json' -Headers @{Authorization = "Bearer $token"} -Body $meetingparam

	$result = Invoke-RestMethod -Headers $authHeader -Uri $apiUrl -Body $body -Method Post -ContentType 'application/json'

	$queryUrlUsers = $graphUrl + "/v1.0/me"

	$user = Invoke-WebRequest -Method Get -Uri $queryUrlUsers -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop

}

Function GenerateAttendeeList{
	
}

$token = GetToken();

$csvOut = "meetingsoutput.csv"

# Get input CSV and scan line by line
$csv = Import-Csv ""
foreach ($line in $file)
{
	$line.Read | out-host
	$lineValues = ([string]$file[$i]).Split(',').Trim()

	$code = $lineValues[0]
	$subject = $lineValues[1]
	$listAttendees = $lineValues[2..($lineValues.Length-1)]
	
	# Create meeting and output generation: c_corso,subject,join link,meeting options url
	$meetingOutput = CreateMeeting($code, $subject, $listAttendees, $token)
	
	# Write output
	Add-content $csvOut -value $meetingOutput
}


