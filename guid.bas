Attribute VB_Name = "guid"
'MSDN Article ID: Q176790
Option Explicit

Type guid
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(7)    As Byte
End Type

Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As guid) As Long
Global Const S_OK = 0 ' return value from CoCreateGuid

Public Function GetGUID() As String

        Dim lResult         As Long
        Dim lguid           As guid
        Dim MyguidString    As String
        Dim MyGuidString1   As String
        Dim MyGuidString2   As String
        Dim MyGuidString3   As String
        Dim DataLen         As Integer
        Dim StringLen       As Integer
        Dim i%

        On Error GoTo error_olemsg

        lResult = CoCreateGuid(lguid)

        If lResult = S_OK Then
        
           MyGuidString1 = Hex$(lguid.Data1)
           StringLen = Len(MyGuidString1)
           DataLen = Len(lguid.Data1)
           MyGuidString1 = LeadingZeros(2 * DataLen, StringLen) _
              & MyGuidString1 'First 4 bytes (8 hex digits)

           MyGuidString2 = Hex$(lguid.Data2)
           StringLen = Len(MyGuidString2)
           DataLen = Len(lguid.Data2)
           MyGuidString2 = LeadingZeros(2 * DataLen, StringLen) _
              & Trim$(MyGuidString2) 'Next 2 bytes (4 hex digits)

           MyGuidString3 = Hex$(lguid.Data3)
           StringLen = Len(MyGuidString3)
           DataLen = Len(lguid.Data3)
           MyGuidString3 = LeadingZeros(2 * DataLen, StringLen) _
              & Trim$(MyGuidString3) 'Next 2 bytes (4 hex digits)

           GetGUID = _
              MyGuidString1 & MyGuidString2 & MyGuidString3

           Dim sByte As String
                
           For i% = 0 To 7
              sByte = Hex(lguid.Data4(i%))
              If Len(sByte) < 2 Then sByte = "0" & sByte
              MyguidString = MyguidString & sByte
           Next i%


           'MyGuidString contains last 8 bytes of Guid (16 hex digits)
           GetGUID = GetGUID & MyguidString

        Else
           GetGUID = "00000000" ' return zeros if function unsuccessful
        End If

        Exit Function

error_olemsg:
         lResult = MessageBox(0, "Error " & Str(Err) & ": " & error$(Err), Form_main.Caption, vbExclamation)
         GetGUID = "00000000"
         Exit Function

End Function

Private Function LeadingZeros(ExpectedLen As Integer, ActualLen As Integer) As String
    LeadingZeros = String$(ExpectedLen - ActualLen, "0")
End Function
