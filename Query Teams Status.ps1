Add-Type -AssemblyName System.Windows.Forms

# Function to query call status for a given team ID
function Get-TeamCallStatus {
    param(
        $teamId,
        $accessToken
    )

    $teamMembersUrl = "https://graph.microsoft.com/v1.0/teams/$teamId/members"
    $teamMembers = Invoke-RestMethod -Uri $teamMembersUrl -Headers @{Authorization = "Bearer $accessToken"}

    $result = @()

    foreach ($member in $teamMembers.value) {
        $userId = $member.id
        $displayName = $member.displayName

        $callStatusUrl = "https://graph.microsoft.com/v1.0/users/$userId/communications/callRecords"
        $callStatus = Invoke-RestMethod -Uri $callStatusUrl -Headers @{Authorization = "Bearer $accessToken"}

        foreach ($call in $callStatus.value) {
            $result += [PSCustomObject]@{
                DisplayName   = $displayName
                CallDirection = $call.direction
                CallState     = $call.callState
                StartTime     = $call.startDateTime
            }
        }
    }

    return $result
}

# Create a form
$form = New-Object Windows.Forms.Form
$form.Text = "Microsoft Teams Call Status Query"
$form.Size = New-Object Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"

$labelClientId = New-Object Windows.Forms.Label
$labelClientId.Text = "Client ID:"
$labelClientId.Location = New-Object Drawing.Point(10, 10)

$textBoxClientId = New-Object Windows.Forms.TextBox
$textBoxClientId.Location = New-Object Drawing.Point(120, 10)
$textBoxClientId.Size = New-Object Drawing.Size(200, 20)

$labelClientSecret = New-Object Windows.Forms.Label
$labelClientSecret.Text = "Client Secret:"
$labelClientSecret.Location = New-Object Drawing.Point(10, 40)

$textBoxClientSecret = New-Object Windows.Forms.TextBox
$textBoxClientSecret.Location = New-Object Drawing.Point(120, 40)
$textBoxClientSecret.Size = New-Object Drawing.Size(200, 20)

$labelTenantId = New-Object Windows.Forms.Label
$labelTenantId.Text = "Tenant ID:"
$labelTenantId.Location = New-Object Drawing.Point(10, 70)

$textBoxTenantId = New-Object Windows.Forms.TextBox
$textBoxTenantId.Location = New-Object Drawing.Point(120, 70)
$textBoxTenantId.Size = New-Object Drawing.Size(200, 20)

# Create a button to trigger the call status query
$buttonQuery = New-Object Windows.Forms.Button
$buttonQuery.Text = "Query Call Status"
$buttonQuery.Location = New-Object Drawing.Point(10, 100)
$buttonQuery.Add_Click({
    # Get access token using provided credentials
    $clientId = $textBoxClientId.Text
    $clientSecret = $textBoxClientSecret.Text
    $tenantId = $textBoxTenantId.Text
    $grantType = "client_credentials"
    $resource = "https://graph.microsoft.com"
    $accessTokenBody = @{
        client_id     = $clientId
        client_secret = $clientSecret
        grant_type    = $grantType
        scope         = "https://graph.microsoft.com/.default"
    }

    $authUrl = "https://login.microsoftonline.com/$tenantId/oauth2/token"
    $accessTokenResponse = Invoke-RestMethod -Uri $authUrl -Method POST -Body $accessTokenBody
    $accessToken = $accessTokenResponse.access_token

    # Query call status for the team and display the results
    $teamId = "ReformIT Team"

    $callStatus = Get-TeamCallStatus -teamId $teamId -accessToken $accessToken

    $output = "Call Status for Team: $teamId`r`n`r`n"
    foreach ($call in $callStatus) {
        $output += "Display Name: $($call.DisplayName)`r`n"
        $output += "Call Direction: $($call.CallDirection)`r`n"
        $output += "Call State: $($call.CallState)`r`n"
        $output += "Start Time: $($call.StartTime)`r`n`r`n"
    }

    $textBoxOutput.Text = $output
})

# Create a text box to display the call status query results
$textBoxOutput = New-Object Windows.Forms.TextBox
$textBoxOutput.Multiline = $true
$textBoxOutput.ScrollBars = "Vertical"
$textBoxOutput.Size = New-Object Drawing.Size(550, 230)
$textBoxOutput.Location = New-Object Drawing.Point(10, 140)

# Add controls to the form
$form.Controls.Add($labelClientId)
$form.Controls.Add($textBoxClientId)
$form.Controls.Add($labelClientSecret)
$form.Controls.Add($textBoxClientSecret)
$form.Controls.Add($labelTenantId)
$form.Controls.Add($textBoxTenantId)
$form.Controls.Add($buttonQuery)
$form.Controls.Add($textBoxOutput)

# Show the form
$form.ShowDialog()
