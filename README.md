# ServiceNowPS
This is a PowerShell module for ServiceNow based on the JSONv2 web service and other Front End API calls:</br>
[ServiceNow JSONv2 Web Service](https://docs.servicenow.com/bundle/washingtondc-api-reference/page/integrate/inbound-other-web-services/concept/c_JSONv2WebService.html)</br></br>
The JSONv2 web service does not require API access, it only requires the "ITIL" user role.</br>(Basically, if you can create/view tickets, this should work for you!)</br></br>

## Examples
### Create ServiceNow Session
```
#Username/Password authentication - Server is not needed if global Server variable is changed at top of script
New-ServiceNowSession -Server "******.service-now.com" -Username "admin" -Pass "pass"

#Smart Card authentication
New-ServiceNowSession -Server "******.service-now.com" -CertificateAuth
```
