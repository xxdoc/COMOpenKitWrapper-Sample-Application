# COMOpenKitWrapper-Sample-Application
Sample VB6 Application Using the COMOpenKitWrapper
# Description 
This is an example VB6 Application that uses the [COMOpenKitWrapper](https://github.com/damass/COMOpenKitWrapper).

It provide various use cases for Initializing the Dynatrace OpenKit, Reporting User Actions, Values & Events, & Tracing Web Requests. 

# Usage 
You will need to provide a few strings before this application can be used. 

In the mainModule (mainModule.bas) you will need to provide the applicationID and the endpointURL. Also optionally, you can change the appVersion, operatingSystem, manufacturer, and ModelID. 
```
'The 3 variables below are necessary for OpenKit initialization
applicationID = "" 'Supply Application ID found from Dynatrace Web UI
deviceID = GUID.GUID()
endpointURL = "" ' Supply EndPoint found from Dynatrace Web UI

'These variables are optional, will take default values if not supplied
appVersion = "1.5.0"
operatingSystem = "Windows 10"
manufacturer = "Dynatrace"
ModelID = "VB6-OpenKitDevice"

'Initialize OpenKit Instance
dummyVariable = COMOpenKit.InitializeOpenKit(endpointURL, applicationID, deviceID, appVersion, operatingSystem, manufacturer, ModelID)

```
Then in the frmLogin (frmLogin.frm) you will need to supply an IPAddress.
```
    Dim IPAddress As String
    IPAddress = "" 'Supply IP Address
    
    COMOpenKit.CreateOpenKitSession (IPAddress)
```
