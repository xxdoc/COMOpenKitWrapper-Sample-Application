VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2205
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1302.787
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Hint : ""password"""
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblPassHint 
      Caption         =   "The Password is ""password"" without quotes"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'When Login Form Loads, it creates OpenKit Session
' with supplied IP address
Private Sub Form_Load()
    frmLogin.Show
    Dim IPAddress As String
    IPAddress = "" 'Supply IP Address
    
    COMOpenKit.CreateOpenKitSession (IPAddress)
    
End Sub

'If someone closes the Login Form, the OpenKit Session i
' ended and the OpenKit instance is shutdown.
Private Sub Form_Terminate()
    Dim dummyVariable
    dummyVariable = COMOpenKit.EndOpenKitSession()
    dummyVariable = COMOpenKit.ShutdownOpenKit()
    Unload Me
    End
End Sub


'User clicked the cancel button on the Login Form.
'Does not do anything functional only creates user action
Private Sub cmdCancel_Click()
    COMOpenKit.EnterRootAction ("Cancel")
    Dim dummyVariable
    
    MsgBox "You cannot Cancel", vbOKOnly
    
    dummyVariable = COMOpenKit.LeaveRootAction()
End Sub

'User clicked the OK button on Login Form. If password
' is correct, will tag session with UserName input text
' if no UserName is provided, "Default" will be used.
' Once user is logged in the Main App Form is loaded.
Private Sub cmdOK_Click()
    COMOpenKit.EnterRootAction ("Login Button")
    Dim dummyVariable
    Dim UserName As String
    UserName = "Default"
    
    'check for correct password
    If txtPassword = "password" Then
        UserName = txtUserName.Text
        COMOpenKit.IdentifyOpenKitUser (UserName)
        Me.Hide
        dummyVariable = COMOpenKit.LeaveRootAction()
        Unload Me
        frmMainPage.Show
    'if password is incorrect, shows MessageBox to try again
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        dummyVariable = COMOpenKit.LeaveRootAction()
    End If
    
End Sub

