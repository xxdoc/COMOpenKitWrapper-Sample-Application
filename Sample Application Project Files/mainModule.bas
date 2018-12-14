Attribute VB_Name = "mainModule"
'Define all Global Variables
'The first are used to initialize OpenKit Instance
Public applicationID As String
Public deviceID As String
Public endpointURL As String
Public appVersion As String
Public operatingSystem As String
Public manufacturer As String
Public ModelID As String

'This will be used throughout the application to call OpenKit Wrapper Methods
Public COMOpenKit As New COMOpenKitWrapper.COMOpenKitWrapper

'Object used to create unique ID for the deviceID for OpenKit
Dim GUID As New GUID

Sub Main()

'Setting values to all OpenKit Initialization Variables
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

'Start Graphical piece of Application
frmLogin.Show
    
    
End Sub
