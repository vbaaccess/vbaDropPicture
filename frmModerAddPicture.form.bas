Option Compare Database
Option Explicit

Private Const CurrentModName = "frmModerAddPicture"

Private objPic As clsModerAddPicture
Private iTrybZapisu As Single
Private iTrybOdczytu As Single

Private Sub Form_Open(Cancel As Integer)
    '--- 1 ---
    Set objPic = New clsModerAddPicture
    
    '--- PARAMETRY USTAWIONE NA TESTY --- --- ---
    iTrybZapisu = 1 ' zapis linku do bazy
    iTrybOdczytu = 1 ' tylko odczyt z linku

End Sub

Private Sub Form_Load()
    '--- 2 ---

    '--- PARAMETRY wyswietlam na formularzu ---
    Call WyewietleniePodglatuTrybuOdczytuPliku(False)
    Me.RamSposobZapisu = iTrybZapisu
    '--- PARAMETRY przekaz do klasy ---
    Call objPic.SET_PARAMITERS(iTrybZapisu, iTrybOdczytu)
    
    '--- ---------------------------- --- --- ---
    
    Call objPic.INIT_PICTURE(Me.ObiekDoWyswietleniaObrazu)
    Call objPic.INIT_SETTINGS(Me.TxtKatalogDocelowy, Me.txtFileList, Me.chkNowaNazwaPliku, Me.TxtNowaNazwaPliku, Me.chkTylkoGraficzne)
    

    

End Sub

Private Sub Cmd_CtrV_Click()
    Call objPic.CtrV
   'Call objPic.GetAllFiles
   'Call objPic.UplodaPictureToControl
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 86 Then
        Call objPic.CtrV
    End If

End Sub

Private Sub CmdWklejNazwyPlikow_Click()
    Call objPic.GetAllFiles
    Call objPic.ShowAllFiles
End Sub

Private Sub CmdWklejPlikiDoKatalogu_Click()
   
   ' Call objPic.SaveFilesFromClipboard
    If Me.chkNowaNazwaPliku Then
        Call objPic.SaveFilesFromClipboard(Me.TxtKatalogDocelowy, Me.chkTylkoGraficzne, Me.TxtNowaNazwaPliku)
    Else
        Call objPic.SaveFilesFromClipboard(Me.TxtKatalogDocelowy, Me.chkTylkoGraficzne)
    End If
   
End Sub

Private Sub LstProcedurWyswietlania_DblClick(Cancel As Integer)
    Call WyewietleniePodglatuTrybuOdczytuPliku
    Call objPic.SET_PARAMITERS(0, iTrybOdczytu)
End Sub

Private Sub WyewietleniePodglatuTrybuOdczytuPliku(Optional bPrzelaczNaKolejnyTryb As Boolean = True)
    Dim sHead$
    Dim sDesc(0 To 2) As String
    Dim iNowyTryb As Single
    
    sHead = "ID;LP;Procedura odczytu"
    
    If bPrzelaczNaKolejnyTryb Then
        Select Case iTrybOdczytu
        Case 0, 4
            iTrybOdczytu = 1
        Case 1
            iTrybOdczytu = 2
        Case 2
            iTrybOdczytu = 3
        Case 3
            iTrybOdczytu = 4
        End Select
    End If
    
    Select Case iTrybOdczytu
    Case 1
        sDesc(1) = "1;1;Odczyt pliku z linku"
        sDesc(2) = "2;;Odczyt pliku z postacji binarnej"
    Case 2
        sDesc(1) = "2;1;Odczyt pliku z postacji binarnej"
        sDesc(2) = "1;;Odczyt pliku z linku"
    Case 3
        sDesc(1) = "1;1;Odczyt pliku z linku"
        sDesc(2) = "2;2;Odczyt pliku z postacji binarnej"
    Case 4
        sDesc(1) = "2;1;Odczyt pliku z postacji binarnej"
        sDesc(2) = "1;2;Odczyt pliku z linku"
    End Select
    
    sDesc(0) = sDesc(1) & ";" & sDesc(2)
    
    Me.LstProcedurWyswietlania.RowSource = sHead & ";" & sDesc(0)

End Sub

Private Sub RamSposobZapisu_AfterUpdate()
    iTrybZapisu = Nz(Me.RamSposobZapisu, 1)
    Call objPic.SET_PARAMITERS(iTrybZapisu, 0)
End Sub

Private Sub TxtNowaNazwaPliku_AfterUpdate()
    objPic.NewFileName = Me.TxtNowaNazwaPliku
End Sub

Private Sub chkNowaNazwaPliku_Click()
    objPic.UseNewFileName = Me.chkNowaNazwaPliku
End Sub

Private Sub chkTylkoGraficzne_Click()
    objPic.CopyPicturesOnly = Me.chkTylkoGraficzne
End Sub
