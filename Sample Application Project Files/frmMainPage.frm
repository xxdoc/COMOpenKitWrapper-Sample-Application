VERSION 5.00
Begin VB.Form frmMainPage 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      Height          =   855
      Left            =   6360
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdRootActionTraceWebRequest 
      Caption         =   "User Action trace Web Request"
      Height          =   855
      Left            =   960
      TabIndex        =   5
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdRootActionError 
      Caption         =   "User Action With Error"
      Height          =   855
      Left            =   840
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdCrash 
      Caption         =   "Simulate Crash"
      Height          =   975
      Left            =   6360
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdRootReportValue 
      Caption         =   "Root Action Report String Value"
      Height          =   855
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdUserAction2 
      Caption         =   "User Action 2 (With Child Action)"
      Height          =   855
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton cmdUserAction1 
      Caption         =   "User Action 1"
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'User Clicked Log Out Button
'OpenKit session is ended and shutdown
'Application is closed.
Private Sub cmdLogOut_Click()
    COMOpenKit.EnterRootAction ("Log Out")
    Dim dummyVariable
    
    
    Unload Me
    dummyVariable = COMOpenKit.LeaveRootAction()
    
    dummyVariable = COMOpenKit.EndOpenKitSession()
    dummyVariable = COMOpenKit.ShutdownOpenKit()
    End
End Sub

'User Clicked User Action trace Web Request button
'Simulates a WebRequest being traced by OpenKit
Private Sub cmdRootActionTraceWebRequest_Click()
    COMOpenKit.EnterRootAction ("Root User Action with Web Request")
    Dim dummyVariable
    Dim headerName As String
    Dim headerValue As String
    
    COMOpenKit.RootActionTraceWebRequest ("https://dynatrace.com")
    dummyVariable = COMOpenKit.RootWebRequestTracerStart()
    
    headerName = COMOpenKit.RootWebRequestTracerGetHeaderName()
    MsgBox headerName, vbOKOnly
    
    headerValue = COMOpenKit.RootWebRequestTracerGetHeaderValue()
    MsgBox headerValue, vbOKOnly
    
    COMOpenKit.RootWebRequestTracerSetByteSent (802)
    COMOpenKit.RootWebRequestTracerSetBytesReceived (323)
    COMOpenKit.RootWebRequestTracerSetResponseCode (200)
    
    dummyVariable = COMOpenKit.RootWebRequestTracerStop()
    
    dummyVariable = COMOpenKit.LeaveRootAction()
End Sub

'If someone closes the Main App Form, the OpenKit Session is
' ended and the OpenKit instance is shutdown
Private Sub Form_Terminate()
    Dim dummyVariable
    dummyVariable = COMOpenKit.EndOpenKitSession()
    dummyVariable = COMOpenKit.ShutdownOpenKit()
    
    Unload Me
    End
End Sub

'User Clicked Simulate Crash, will report a "crash" event
' to Dynatrace & end/shutdown OpenKit & application
Private Sub cmdCrash_Click()
    Dim dummyVariable
    Dim errorName As String
    Dim reason As String
    Dim stacktrace As String
    
    errorName = "Simulated Crash"
    reason = "Button Clicked"
    stacktrace = "1:System Crashed" + vbNewLine + "2:Simulate Crash Button Clicked"
    
    dummyVariable = COMOpenKit.ReportCrash(errorName, reason, stacktrace)
    
    dummyVariable = COMOpenKit.EndOpenKitSession()
    dummyVariable = COMOpenKit.ShutdownOpenKit()
    
    Me.Hide
    Unload Me
End Sub

'User Clicked User Action with Error Button,
' Sends Root User Action Error to Dynatrace
Private Sub cmdRootActionError_Click()
    COMOpenKit.EnterRootAction ("Root User Action with Error")
    Dim dummyVariable
    dummyVariable = COMOpenKit.ReportRootActionError("Unknown Error", 42, "Not sure what's going on here")
    dummyVariable = COMOpenKit.LeaveRootAction()
End Sub

'User Clicked Root Action Report String Value button
' Will send String value to associate with Root User Action
Private Sub cmdRootReportValue_Click()
    COMOpenKit.EnterRootAction ("Root User Action with String Value")
    Dim dummyVariable
    dummyVariable = COMOpenKit.ReportRootActionStringValue("Programming Language", "VB6")
    dummyVariable = COMOpenKit.LeaveRootAction()
End Sub

'User Clicked Root User Action 1 Button,
'Simulates simple Root User Action
Private Sub cmdUserAction1_Click()
    Dim dummyVariable

    COMOpenKit.EnterRootAction ("Root User Action 1 Button")
    COMOpenKit.ReportRootActionEvent ("User Action 1 Triggered")
    MsgBox "User Action 1 Triggered", vbOKOnly
    dummyVariable = COMOpenKit.LeaveRootAction()
End Sub

'User Clicked Root User Action 2 Button
'Loads Child Action Form to simulate Child Action
'within Root Action
Private Sub cmdUserAction2_Click()
    Dim dummyVariable

    COMOpenKit.EnterRootAction ("Root User Action 2 Button")
    MsgBox "User Action 2 Triggered", vbOKOnly
    frmChildAction.Show
    dummyVariable = COMOpenKit.LeaveRootAction()
End Sub
