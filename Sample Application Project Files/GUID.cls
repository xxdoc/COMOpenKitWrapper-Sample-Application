VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type GUID
    PartOne As Long
    PartTwo As Integer
    PartThree As Integer
    PartFour(7) As Byte
End Type
      
Private Declare Function CoCreateGuid Lib "OLE32.DLL" _
(ptrGuid As GUID) As Long


Private Sub Class_Initialize()
End Sub

'RETURNS:  GUID if successful; blank string otherwise.
'Unlike the GUIDS in the registry, this function returns GUID
'without "-" characters.  See comments for how to modify if you
'want the dash.
'http://www.freevbcode.com/ShowCode.asp?ID=21


Public Function GUID() As String
    Dim lRetVal As Long
    Dim udtGuid As GUID
    
    Dim sPartOne As String
    Dim sPartTwo As String
    Dim sPartThree As String
    Dim sPartFour As String
    Dim iDataLen As Integer
    Dim iStrLen As Integer
    Dim iCtr As Integer
    Dim sAns As String
   
    On Error GoTo errorhandler
    sAns = ""
    
    lRetVal = CoCreateGuid(udtGuid)
    
    If lRetVal = 0 Then
    
       'First 8 chars
        sPartOne = Hex$(udtGuid.PartOne)
        iStrLen = Len(sPartOne)
        iDataLen = Len(udtGuid.PartOne)
        sPartOne = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartOne)
        
        'Next 4 Chars
        sPartTwo = Hex$(udtGuid.PartTwo)
        iStrLen = Len(sPartTwo)
        iDataLen = Len(udtGuid.PartTwo)
        sPartTwo = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartTwo)
           
        'Next 4 Chars
        sPartThree = Hex$(udtGuid.PartThree)
        iStrLen = Len(sPartThree)
        iDataLen = Len(udtGuid.PartThree)
        sPartThree = String((iDataLen * 2) - iStrLen, "0") _
        & Trim$(sPartThree)   'Next 2 bytes (4 hex digits)
           
        'Final 16 chars
        For iCtr = 0 To 7
            sPartFour = sPartFour & _
            Format$(Hex$(udtGuid.PartFour(iCtr)), "00")
        Next
 
     'To create GUID with "-", change line below to:
     'sAns = sPartOne & "-" & sPartTwo & "-" & sPartThree _
     '& "-" & sPartFour
       
       sAns = sPartOne & sPartTwo & sPartThree & sPartFour
            
        End If
        
        GUID = sAns
Exit Function


errorhandler:
'return a blank string if there's an error
Exit Function
End Function
