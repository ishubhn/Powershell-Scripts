#ServiceNow creds
$snowusername = ""
$snowpassword = ""

# Build auth header  
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $snowusername, $snowpassword)))

# Set proper headers
$postheaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$postheaders.Add("Authorization", ("Basic {0}" -f $base64AuthInfo))
$postheaders.Add("Accept", "Application/json")
$postheaders.Add("Content-Type", "Application/json")
$postmethod = "post"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Tls11, [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Ssl3
# Specify endpoint uri
$uri = ""

# Specify request body
$body = @{ 
    requested_for = "Wanjari Gaurav Manoharrao-ext"
    contact_type = "Phone"
    contact_number= "+919123456765"
    state = "New"
    assignment_group = "TS-AMER-L2-INTEGRATION HUB"
    its_category = "Application"
    location = "Asia & JPac/India/Pune/PNQ External Site"
    urgency = "2"
    service_offering = "Symphony - Sanofi NA Veeva CRM"
    short_description = "Test"
    caller_id = "Wanjari Gaurav Manoharrao-ext"
    its_symptom = ""
}

$bodyjson = $body | ConvertTo-Json

# Send HTTP request
$response = Invoke-RestMethod -Headers $postheaders -Method $postmethod -Uri $uri -Body $bodyjson -ErrorAction Stop

$response.result

# Print response
#$CreateServiceIncident.RawContent
#-ContentType "application/json