Option Compare Database
Option Explicit

Private Const CurrentModName = "clsModerAddPicture"
Private Const INIT_STRING_EVENT = "[Event Procedure]"

' Klasa do kopiowania plikow graficznych do projektu mdb, adp
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal drop_handle As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long

Private Const CF_HDROP As Long = 15

Private m_DestinationCatalogForAttachments As String
Private m_fileCount As Long
Private m_Files() As String
Private m_GraphicExtensions As String   ' Dozwolone rozszerzenia plików graficznych '
Private m_UseNewFileName As Boolean
Private m_NewFileName As String
Private m_PicturesOnly   As Boolean

Private m_ZapisLinku As Boolean
Private m_ZapisPostaciBinarnej As Boolean   '0 - pomin (NULL)
Private m_Odczyt(1 To 2) As Single          '1 - Odczyt z linku
                                            '2 - Odczyt z postaci binarnej
Private WithEvents ctrImage As Image
Private WithEvents ctrTextBoxFileList As TextBox
Private WithEvents ctrTextBoxDestinationCatalogForAttachments As TextBox
Private WithEvents ctrChkBoxNewFileName As CheckBox
Private WithEvents ctrTextBoxNewFileName As TextBox
Private WithEvents ctrChkBoxPicturesOnly As CheckBox

Private Sub Class_Initialize()
    '--- Konstruktor --- przypisanie domyslnych wartosci
    
'    m_DestinationCatalogForAttachments = CurrentProject.Path & "\Temp\"
    m_DestinationCatalogForAttachments = "C:\Ttemp\"
    m_GraphicExtensions = "bmp|jpg|jpeg|png|gif"
    
    m_UseNewFileName = False
    m_NewFileName = "PlikTmp"
    m_PicturesOnly = True

End Sub

Public Function INIT_PICTURE(objImage As Image, Optional objPath As TextBox)
    Set ctrImage = objImage
    ctrImage.OnDblClick = INIT_STRING_EVENT
    
    '--- inicjowanie funkcji dla obiektu obrazka ---
End Function

Public Function INIT_SETTINGS(objTextBoxDestinationCatalogForAttachments As TextBox _
                            , objTextBoxFileList As TextBox _
                            , objChkBoxNewFileName As CheckBox _
                            , objTextBoxNewFileName As TextBox _
                            , objChkBoxPicturesOnly As CheckBox)

    Set ctrTextBoxDestinationCatalogForAttachments = objTextBoxDestinationCatalogForAttachments
    Set ctrTextBoxFileList = objTextBoxFileList
    
    Set ctrChkBoxNewFileName = objChkBoxNewFileName
    Set ctrTextBoxNewFileName = objTextBoxNewFileName
    Set ctrChkBoxPicturesOnly = objChkBoxPicturesOnly
    
    '--- inicjowanie funkcji zwiaznych z ustawieniami i podglada czy dziala ---
    ctrTextBoxDestinationCatalogForAttachments.Value = m_DestinationCatalogForAttachments
    
    ctrChkBoxNewFileName.Value = m_UseNewFileName
    ctrTextBoxNewFileName.Value = m_NewFileName
    ctrChkBoxPicturesOnly.Value = m_PicturesOnly

End Function

Public Function SET_PARAMITERS(iPobieranie As Single, iWyswietlanie As Single) As Boolean
    Dim bError As Boolean
    
    If iPobieranie > 0 Then
        m_ZapisPostaciBinarnej = False
        m_ZapisLinku = True ' 1
        If iPobieranie = 2 Then _
            m_ZapisPostaciBinarnej = True
    End If

    If iWyswietlanie > 0 Then
        m_Odczyt(1) = 0
        m_Odczyt(2) = 0                         '0 - pomin (NULL)
        Select Case iWyswietlanie               '1 - Odczyt z linku
        Case 1                                  '2 - Odczyt z postaci binarnej
            m_Odczyt(1) = 1
        Case 2
            m_Odczyt(1) = 2
        Case 3
            m_Odczyt(1) = 1
            m_Odczyt(2) = 2
        Case 4
            m_Odczyt(1) = 2
            m_Odczyt(2) = 1
        End Select
    End If
    
    SET_PARAMITERS = Not (bError)
End Function

Public Sub CtrV()
    Me.GetAllFiles
    Me.UplodaPictureToControl
End Sub

Public Sub GetAllFiles()
' Pobiera ze schowka dane skopiowanych plików (przez CTRL-C). '
' Zwraca array string m_Files() ze ścizkami/nazwami plików, a przez referencję ilośc plików w zmiennej m_fileCount '
    Dim hDrop As Long, i As Long
    Dim sFileName As String * 1024

    m_fileCount = 0

    If Not CBool(IsClipboardFormatAvailable(CF_HDROP)) Then Exit Sub
    If Not CBool(OpenClipboard(0&)) Then Exit Sub

    hDrop = GetClipboardData(CF_HDROP)
    If Not CBool(hDrop) Then GoTo done

    m_fileCount = DragQueryFile(hDrop, -1, vbNullString, 0)

    ReDim m_Files(m_fileCount - 1)
    For i = 0 To m_fileCount - 1
        DragQueryFile hDrop, i, sFileName, Len(sFileName)
        m_Files(i) = Left$(sFileName, InStr(sFileName, vbNullChar) - 1)
    Next


    '--- Jeśli filtrujemy tylk obrazki to uruchamiam nową funkcję filtrującą FilterFiles()
    If m_PicturesOnly And m_fileCount > 0 Then Call FilterFiles

done:
    If m_fileCount = 0 Then
        VBA.MsgBox "Nie znaleziono plików "
    End If

    CloseClipboard

End Sub


Public Sub UplodaPictureToControl(Optional ctrOptionalImage As Image)
    Dim ctrTargetImage As Image

    '1) opcjonalnie podane inne pole do wyswieltenia plikow
    If Not ctrOptionalImage Is Nothing Then _
        Set ctrTargetImage = ctrOptionalImage
    
    '2) zdefiniowane pole image o ile nie podano opcjonalnego
    If Not ctrImage Is Nothing And ctrTargetImage Is Nothing Then _
        Set ctrTargetImage = ctrImage
       
    If ctrTargetImage Is Nothing Then Exit Sub
    If m_fileCount = 0 Then Exit Sub
'--- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
    Dim bPoprawnePodpiecie As Boolean
    
    bPoprawnePodpiecie = AttachFileToImageControl(ctrTargetImage)
    
    If Not (bPoprawnePodpiecie) Then
        VBA.MsgBox "Błąd załadowanie pliku do podglądu!"
    End If
End Sub

Private Function AttachFileToImageControl(ctrTargetImage As Image, Optional ctrTxtBox As TextBox) As Boolean
On Error GoTo ExitFunction
    ctrTargetImage.Picture = m_Files(0)
    'Me.TxtSourceFile = m_Files(0)
    AttachFileToImageControl = True
ExitFunction:
End Function

Public Sub ShowAllFiles(Optional ctrOptionalTextBox As TextBox)
    Dim ctrTargetTextBox As TextBox
    
    '1) opcjonalnie podane inne pole tekstowe do wyswieltenia plikow
    If Not ctrOptionalTextBox Is Nothing Then _
        Set ctrTargetTextBox = ctrOptionalTextBox
    
    '2) zdefiniowane pole textowe o ile nie podano opcjonalnego
    If Not ctrTextBoxFileList Is Nothing And ctrTargetTextBox Is Nothing Then _
        Set ctrTargetTextBox = ctrTextBoxFileList
    
    
    If ctrTargetTextBox Is Nothing Then
        VBA.MsgBox "Brak wskazanego obiektu do wyświetlenia danych"
        Exit Sub
    End If
'--- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
    ctrTargetTextBox.Value = vbNullString
    
    Dim i As Long

    If (m_fileCount = 0) Then
        VBA.MsgBox "Brak danych do wyświetlenia"
    Else
        For i = 0 To m_fileCount - 1
        '    MsgBox "Znaleziono: " & a(i)
            ctrTargetTextBox.Value = ctrTextBoxFileList.Value & m_Files(i) & vbCrLf
        Next
    End If
End Sub

Public Function SaveFilesFromClipboard_OLD(Optional DestinationFolder As String, Optional sNewFileName As String) As Boolean
' Wkleja wszystkie pliki ze schowka do katalogu docelowego DestinationFolder '
    Dim sTargetPath$
    
    If Len(Trim(DestinationFolder)) = 0 Then
        sTargetPath = ctrTextBoxDestinationCatalogForAttachments
    Else
        sTargetPath = DestinationFolder
    End If
    
    '--- utworzenie tymaczasoweg katalogu
    If Dir(sTargetPath, vbDirectory) = "" Then MkDir (sTargetPath)
    
    SaveFilesFromClipboard_OLD = False
    If Not CBool(IsClipboardFormatAvailable(CF_HDROP)) Then Exit Function
    CreateObject("Shell.Application").Namespace(CVar(sTargetPath)).self.InvokeVerb "Paste"
    SaveFilesFromClipboard_OLD = True
End Function


Public Function SaveFilesFromClipboard(Optional DestinationFolder As String, Optional GetPicturesOnly As Boolean = False, Optional sNewFileName As String) As Boolean
' Wkleja pliki ze schowka do katalogu docelowego DestinationFolder '
' Wywołania:
' ? SaveFilesFromClipboard("c:\graph")
' ? SaveFilesFromClipboard("c:\graph", TRUE)
' ? SaveFilesFromClipboard("c:\graph", TRUE, "Obraz")
    Dim sTargetPath As String
    Dim i As Integer
    Dim FileExtension As String, sFile As String    ' orginalne pliki nazwa i rozszerzenie
    Dim sNewFileNameFinal As String                      ' Nazwa nowego pliku
    Dim m_FilesFinal() As String     '              ' ostateczna tablica plików do wklejenia
    Dim GraphCount As Integer                       ' liczba obrazków wykrytych w schowku
    Dim fso As Object
    
    If Len(Trim(DestinationFolder)) = 0 Then
        sTargetPath = ctrTextBoxDestinationCatalogForAttachments
    Else
        sTargetPath = DestinationFolder
    End If
    
    '--- utworzenie tymaczasoweg katalogu
    If Dir(sTargetPath, vbDirectory) = "" Then MkDir (sTargetPath)
    
    m_FilesFinal = m_Files
    
    ' Przygotowaliśmy tablicę z ścieżką/nazwami plików WSZYSTKIE lub GRAFICZNE, teraz kolej na wklejenie do nowej lokalizacji
    ' albo z oryginalną nazwą, albo z wyznaczoną w ostatnim parametrze funkcji, czyli sNewFileName
    GraphCount = 0      ' Jeśli zapisujemy pod nową nazwą pliku, to to będzie licznik '
    For i = LBound(m_FilesFinal) To UBound(m_FilesFinal)
        
        If Len(Nz(sNewFileName, "")) > 0 Then
            ' musimy znać rozszerzenie oryginalnego pliku, jeśli mamy go przemianować na nową sztuczną nazwę '
            FileExtension = Right$(m_FilesFinal(i), Len(m_FilesFinal(i)) - InStrRev(m_FilesFinal(i), "."))
            sNewFileNameFinal = sNewFileName & "_" & Format(GraphCount, "0000") & "." & FileExtension
            GraphCount = GraphCount + 1
        Else    ' jeśli kopiujemy plik pod oryginalną nazwą na nową ścieżkę '
            sNewFileNameFinal = Mid(m_FilesFinal(i), InStrRev(m_FilesFinal(i), "\") + 1)    ' tylko nazwa pliku z rozszerzeniem bez ścieżki '
        End If
        
        ' Doklejamy ścieżkę aby powstał plik docelowy '
        sNewFileNameFinal = sTargetPath & sNewFileNameFinal

        ' Kopiujemy plik na nową lokalizację z nową lub oryginalną nazwą '
        Set fso = CreateObject("Scripting.FileSystemObject")
        Call fso.CopyFile(m_FilesFinal(i), sNewFileNameFinal, True)
        Set fso = Nothing
    
    Next i
    
    SaveFilesFromClipboard = True

End Function

Public Sub FilterFiles()
Dim m_FilesGraphics() As String         ' Tylko lista plików z grafikami
Dim GraphCount As Integer               ' liczba obrazków wykrytych w schowku
Dim i As Integer
Dim sFile As String, FileExtension As String    ' orginalne pliki nazwa i rozszerzenie

    ' Jeśli nie ustawiono formatów, które mają się znaleźć w schowku to wychodzimy
    If Nz(m_GraphicExtensions, "") = "" Then Exit Sub

    '--- Przepisuję listę plików z tablicy m_Files() do m_FilesGraphics() bez plików, których rozszerzenie sięnie zgadzają '
        GraphCount = 0
        For i = 0 To m_fileCount - 1
            sFile = m_Files(i)  ' kopiuję nazwę / ścieżkę do sprawdzenia czy grafika
            FileExtension = Right$(sFile, Len(sFile) - InStrRev(sFile, "."))
            
            ' jeśli rozszerzenie bieżącego pliku znajdjesię na liście grafik to dodajemy ten plik do nowej tablicy m_FilesGraphics
            If InStr(1, m_GraphicExtensions, FileExtension, vbTextCompare) > 0 Then
                ReDim Preserve m_FilesGraphics(GraphCount)      ' powiększam nową tablicę
                m_FilesGraphics(GraphCount) = m_Files(i)        ' przepisuję ścieżkę  grafiki z do nowej tablicy
                GraphCount = GraphCount + 1                     ' zwiększam licznik
            End If
        Next
        m_fileCount = GraphCount                ' ostateczna liczba plików jeśli graficzne (nie jako jako index tablicy, tylko rzeczywista )
        
        ' Na końcu zwracam przefiltrowaną tablicę do tej z której wywołałem '
        m_Files = m_FilesGraphics

End Sub

' HURRRRRRRA !!!!! Nauczyłem się robić property :D , BRAWO :))

Property Let CopyPicturesOnly(DefaultPicturesOnly As Boolean)
  m_PicturesOnly = DefaultPicturesOnly
End Property

Property Get CopyPicturesOnly() As Boolean
  CopyPicturesOnly = m_PicturesOnly
End Property

Property Let UseNewFileName(DefaultUseNewFileName As Boolean)
  m_UseNewFileName = DefaultUseNewFileName
End Property

Property Get UseNewFileName() As Boolean
  UseNewFileName = m_UseNewFileName
End Property

Property Let NewFileName(DefaultUseNewFileName As String)
  m_NewFileName = DefaultUseNewFileName
End Property

Property Get NewFileName() As String
  NewFileName = m_NewFileName
End Property

Property Let GraphicExtensions(DefaultUseGraphicExtensions As String)
  m_GraphicExtensions = DefaultUseGraphicExtensions
End Property

Property Get GraphicExtensions() As String
  GraphicExtensions = m_GraphicExtensions
End Property
