VERSION 5.00
Begin VB.Form frmChildAction 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   855
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Child Action Form"
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1230
   End
End
Attribute VB_Name = "frmChildAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'User Clicked Okay button, closes Child Action Form
'and ends child action
Private Sub cmdOK_Click()
    Dim dummyVariable
    dummyVariable = COMOpenKit.LeaveChildAction()
    Me.Hide
    Unload Me
End Sub

'Starts Child Action when form is loaded.
'Also reports Int value to child action in Dynatrace
Private Sub Form_Load()
    COMOpenKit.EnterChildAction ("Child Action 1")
    Dim dummyVariable
    dummyVariable = COMOpenKit.ReportChildActionIntValue("VB6 SubVersion", 123)
End Sub

'User clicked red x to close form, same behavior as
'if user clicked okay button.
Private Sub Form_Terminate()
    Dim dummyVariable
    dummyVariable = COMOpenKit.LeaveChildAction()
    Me.Hide
    Unload Me
End Sub
