Option Compare Database
Option Explicit

Private Const CurrentModName = "mFunkcje"

Public Sub rstSetNothing(ByRef Rst As ADODB.Recordset)

    On Error GoTo rstSetNothing_Error

    If Rst Is Nothing Then Exit Sub

    If Rst.State = adStateOpen Then
        Rst.Close
    End If

    Set Rst = Nothing

    On Error GoTo 0
    Exit Sub

rstSetNothing_Error:
    Exit Sub
End Sub

Public Function OpenRst(ByRef Rst As ADODB.Recordset, ByVal Sql As String _
                   , Optional ShowPossibleErrorMsgBox As Boolean = True _
                   , Optional ByRef ErrNumber As Long = 0 _
                   , Optional ByRef ErrDescription As String = "" _
                   , Optional SqlTimeout As Long = 0 _
                   ) As Boolean
    
                  '

    Const sfName = "OpenRst"
On Error GoTo Err_OpenRst
    ErrNumber = 0

    Set Rst = New ADODB.Recordset

    If SqlTimeout > 0 Then
        Dim con As ADODB.Connection
        Set con = CurrentProject.Connection
        con.CommandTimeout = SqlTimeout

        Rst.Open Sql, con, adOpenKeyset, adLockOptimistic, adCmdUnknown
    Else
        Rst.Open Sql, CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdUnknown
    End If

    If Rst.RecordCount <= 0 Then
        OpenRst = False
    Else
        OpenRst = True
    End If

Exit_PROCEDURE:
    Exit Function

Err_OpenRst:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    If ShowPossibleErrorMsgBox Then
        Select Case ErrNumber
            '    Case NrBledu
            '        Debug.Print "(" & ErrNumber & ") - " & CurrentModName & "." & sfName
        Case Else
            VBA.MsgBox "(" & ErrNumber & IIf(Erl = 0, "", "," & Erl) & ") " & _
                       "- " & CurrentModName & "." & sfName & vbLf & _
                       ErrDescription, vbOKOnly + vbInformation, "Uwaga"
        End Select
        Resume Exit_PROCEDURE
        'Resume
    End If
End Function

Public Function WeryfikacjaTworzenieKataloguTymczasowego(Optional ByRef NazwaKataloguTymczasowego As String) As Boolean
    Dim werFolder As String
    Dim bExist As Boolean
    Dim cf, fso
    
    Const sfName = "WeryfikacjaTworzenieKataloguTymczasowego"
    Dim ErrNumber, ErrDescription
On Error GoTo Err_PROCEDURE

10        werFolder = Application.CurrentProject.Path & "\" & TEMPORARY_DIR
20        NazwaKataloguTymczasowego = werFolder

22        Set fso = CreateObject("Scripting.FileSystemObject")
          
30        If fso.FolderExists(werFolder) Then
40            bExist = True
50        Else
60            Set cf = fso.CreateFolder(werFolder)
70            If fso.FolderExists(werFolder) Then
80                bExist = True
90            End If
100       End If

Exit_PROCEDURE:
    WeryfikacjaTworzenieKataloguTymczasowego = bExist
    Exit Function

Err_PROCEDURE:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    Select Case ErrNumber
'    Case NrBledu
'        Debug.Print "(" & ErrNumber  & ") - " & CurrentModuleName & "." & sfName

    Case Else
        VBA.MsgBox "(" & ErrNumber & IIf(Erl = 0, "", "," & Erl) & ") " & _
               "- " & CurrentModuleName & "." & sfName & vbLf & _
               ErrDescription, vbOKOnly + vbInformation, "Uwaga"
    End Select
    Resume Exit_PROCEDURE
    Resume
End Function