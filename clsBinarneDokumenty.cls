Option Compare Database
Option Explicit

Private Const CurrentModName = "clsBinarneDokumenty"
Private sMsgPL As String

Public Function ZapiszPlikDokumentuWPostaciBinarnej(werIdDokumentu As Long) As Boolean
    Dim MsgBoxPrompt As String
    Dim Rst As Recordset
    
    Dim sTempDir As String
    Dim FilePath As String
    
    Call WeryfikacjaRekorduDokumentu(werIdDokumentu, sTempDir, FilePath, MsgBoxPrompt, Rst)

    If Len(MsgBoxPrompt) > 0 Then
        sMsgPL = "Operacja niedostępna!"
        MsgBoxPrompt = sMsgPL & vbLf & MsgBoxPrompt
        Exit Function
    End If
    
    '--- sprawdzam czy plik juz nie byl zapisany ---
    If Not IsNull(Rst!BinaryFormFile) Then
        sMsgPL = "Odnaleziono zapis binarny w bazie, czy chcesz go nadpisać ?"
        MsgBoxPrompt = sMsgPL
        sMsgPL = "UWAGA"
        If VBA.MsgBox(MsgBoxPrompt, vbQuestion + vbYesNo + vbDefaultButton1, sMsgPL) = vbNo Then
            Exit Function
        End If
    End If
    
    '--- zapis do bazy ---
    Dim sT As ADODB.Stream
    Set sT = New Stream
    
    sT.Type = adTypeBinary
    sT.Open
    sT.LoadFromFile FilePath
    
    'FileOriginalName    FileExtension
    Rst.Fields("FileOriginalName").Value = Mid(FilePath, InStrRev(FilePath, "\") + 1, InStrRev(FilePath, ".") - InStrRev(FilePath, "\") - 1)
    Rst.Fields("FileExtension").Value = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "."))
    Rst.Fields("BinaryFormFile").Value = sT.Read
    Rst.Update
    
    Call rstSetNothing(Rst)
    
    ZapiszPlikDokumentuWPostaciBinarnej = True
End Function

Public Function OdczytDokumentuWPostaciBinarnej(werIdDokumentu As Long) As String
    Dim bGenerujPlik As Boolean
    Dim MsgBoxPrompt As String
    Dim Rst As Recordset
    
    Dim sTempDir As String
    Dim FilePath As String
    
    Call WeryfikacjaRekorduDokumentu(werIdDokumentu, sTempDir, FilePath, MsgBoxPrompt, Rst)
    
    If Len(MsgBoxPrompt) > 0 Then
        sMsgPL = "Operacja niedostępna!"
        MsgBoxPrompt = sMsgPL & vbLf & MsgBoxPrompt
        Exit Function
    End If
    
    '--- odczyt
    
    Dim sTempFile As String
    Dim fso

    sTempFile = sTempDir & "\" & werIdDokumentu & "." & Rst.Fields("FileExtension")

    Set fso = CreateObject("scripting.FileSystemObject")
    
    '--- 1 --- sprawdzam czy plik juz nie istnieje ---
    bGenerujPlik = True
    'If fso.FileExists(sTempFile) Then  'jesli bysmy nie chcieli nadpisywac
    '    bGenerujPlik = False
    'End If

    If bGenerujPlik Then
        Dim LengthOfTheDataExtracted As Long
        
        LengthOfTheDataExtracted = BlobToFileOnProjectPath(Rst.Fields("BinaryFormFile"), sTempFile)
        If Len(sTempFile) > 0 And LengthOfTheDataExtracted <> 0 Then
        End If
    
    End If
    
    Dim sLinkDoPlikuTymczasowegoOdczytanegoDokumentu As String
    
    sLinkDoPlikuTymczasowegoOdczytanegoDokumentu = sTempFile
    
    OdczytDokumentuWPostaciBinarnej = sLinkDoPlikuTymczasowegoOdczytanegoDokumentu
End Function

Private Function WeryfikacjaRekorduDokumentu(werIdDokumentu As Long, werTempDir As String, werFilePath As String _
                                                           , Optional MsgBoxPrompt As String _
                                                           , Optional Rst As Recordset) As Boolean
    Dim sMsg As String
    Dim Sql As String
    Dim fso

    Set fso = CreateObject("scripting.FileSystemObject")

    If werIdDokumentu > 0 Then
        ' odwolanie sie do danych
        Sql = "SELECT * FROM tDokumenty WHERE IdDokumentu=" & werIdDokumentu
        If OpenRst(Rst, Sql) Then
            If Len(Trim(Nz(Rst!Link, ""))) = 0 Then
                sMsgPL = " - brak podanego linku,"
                sMsg = sMsg & vbLf & sMsgPL
            Else
                werFilePath = Trim(Nz(Rst!Link, ""))
                If fso.FileExists(werFilePath) Then
                    'weryfikacja czy wskazany link jest linkiem do dysku i pliku
                    'sMsgPL = " - wartosc w polu link nie dotyczy pliku,"
                    'sMsg = sMsg & vbLf & sMsgPL
                Else
                    sMsg = sMsg & vbLf & " - brak pliku,"
                End If
            End If
        Else
            sMsgPL = " - nie odnaleziono wskazanego rekordu dokumentu,"
            sMsg = sMsg & vbLf & sMsgPL
        End If
    Else
        sMsgPL = " - brak wskazanego rekordu dokumentu,"
        sMsg = sMsg & vbLf & sMsgPL
    End If
    
    If WeryfikacjaTworzenieKataloguTymczasowego(werTempDir) Then
    Else
        sMsgPL = " - bład odczytu/dostępu do katalogu tymczasowego,"
        sMsg = sMsg & vbLf & sMsgPL
    End If
    
    If Len(sMsg) > 0 Then
        If Len(MsgBoxPrompt) = 0 Then
            MsgBoxPrompt = sMsg
        Else
            MsgBoxPrompt = vbLf & sMsg
        End If
    End If
    
    WeryfikacjaRekorduDokumentu = True
End Function
