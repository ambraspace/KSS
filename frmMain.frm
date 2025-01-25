VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "KSS Zaposlenja"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Štampa"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ctlImageList 
      Left            =   5400
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "imgOK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0796
            Key             =   "imgBad"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar ctlStatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   4230
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Upišite niz znakova koji tražite"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.ComboBox cboSearch 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Polje u kojem tražite navedeni znakovni niz"
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Traži"
      Height          =   315
      Left            =   4800
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin MSComctlLib.ListView ctlListView 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   538
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ctlImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Obriši"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Izmijeni"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Dodaj"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin MSComctlLib.TabStrip ctlTabStrip 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5106
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   2819
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Davaoci usluga"
            Key             =   "tabDavaoci"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Korisnici usluga"
            Key             =   "tabKorisnici"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Poslovi"
            Key             =   "tabPoslovi"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  Dim c As MSComctlLib.ListItem
  Select Case Me.ctlTabStrip.SelectedItem.Index
    Case 1
      frmDavalac.AddNew
      If frmDavalac.OKPressed Then
        Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
        Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
        Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
        Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
        Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
        Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
        If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
        Else
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = ""
        End If
        For Each c In Me.ctlListView.ListItems
          c.Selected = False
        Next
        Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).Selected = True
        Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).EnsureVisible
      End If
    Case 2
      frmKorisnik.AddNew
      If frmKorisnik.OKPressed Then
        Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
        Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
        If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
        Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
        Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
        If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
        If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
        For Each c In Me.ctlListView.ListItems
          c.Selected = False
        Next
        Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).Selected = True
        Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).EnsureVisible
      End If
    Case 3
      If de.rsDavaoci.RecordCount = 0 Then
        MsgBox "Unesite bar jednog davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
      End If
      If de.rsKorisnici.RecordCount = 0 Then
        MsgBox "Unesite bar jednog korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
      End If
      frmPosao.AddNew
      If frmPosao.OKPressed Then
        de.rsDavaoci.MoveFirst
        de.rsDavaoci.Find "RB=" & de.rsPoslovi!DavalacRB
        de.rsKorisnici.MoveFirst
        de.rsKorisnici.Find "RB=" & de.rsPoslovi!KorisnikRB
        de.rsIntervali.MoveFirst
        de.rsIntervali.Find "ID=" & de.rsPoslovi!IntervalID
        Me.ctlListView.ListItems.Add , "ID" & de.rsPoslovi!ID, de.rsPoslovi!ID
        Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(1) = UCase(de.rsDavaoci!Prezime) & ", " & de.rsDavaoci!Ime
        If de.rsKorisnici!FL Then
          Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Ime
        Else
          Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Mjesto
        End If
        Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(3) = de.rsPoslovi!Datum
        Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(4) = Format(de.rsPoslovi!Iznos, "#,##0.00 KM")
        Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(5) = de.rsIntervali!Interval
        Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(6) = de.rsPoslovi!Kolicina
        For Each c In Me.ctlListView.ListItems
          c.Selected = False
        Next
        Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).Selected = True
        Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).EnsureVisible
        ApplyIcons
      End If
  End Select
  ctlListView_Click
End Sub

Private Sub cmdDel_Click()
  Dim a As Integer, i As Integer, bFinished As Boolean, dFinDate As Date
  Dim c As MSComctlLib.ListItem
  Select Case Me.ctlTabStrip.SelectedItem.Index
    Case 1
      a = MsgBox("Da li ste sigurni da želite obrisati" & vbCrLf & "odabrane davaoce usluga?", vbQuestion + vbYesNo, Me.Caption)
      If a = vbYes Then
        For i = Me.ctlListView.ListItems.Count To 1 Step -1
          Set c = Me.ctlListView.ListItems(i)
          If c.Selected Then
            de.cn1.Execute "DELETE FROM Davaoci WHERE RB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM Poslovi WHERE DavalacRB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM PosloviBKP WHERE DavalacRB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM RaspolozivostVeze WHERE RB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM Telefoni WHERE Korisnik<>TRUE AND RB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM ZanimanjaVeze WHERE RB=" & Mid(c.Key, 3)
            Me.ctlListView.ListItems.Remove c.Key
          End If
        Next
        de.rsDavaoci.Requery
        de.rsPoslovi.Requery
        de.rsPosloviBKP.Requery
        de.rsRaspolozivostVeze.Requery
        de.rsTelefoni.Requery
        de.rsZanimanjaVeze.Requery
      End If
    Case 2
      a = MsgBox("Da li ste sigurni da želite obrisati" & vbCrLf & "odabrane korisnike usluga?", vbQuestion + vbYesNo, Me.Caption)
      If a = vbYes Then
        For i = Me.ctlListView.ListItems.Count To 1 Step -1
          Set c = Me.ctlListView.ListItems(i)
          If c.Selected Then
            de.cn1.Execute "DELETE FROM Korisnici WHERE RB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM KorZanimanjaVeze WHERE RB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM Poslovi WHERE KorisnikRB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM PosloviBKP WHERE KorisnikRB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM Telefoni WHERE Korisnik=TRUE AND RB=" & Mid(c.Key, 3)
            de.cn1.Execute "DELETE FROM ZiroRacuni WHERE RB=" & Mid(c.Key, 3)
            Me.ctlListView.ListItems.Remove c.Key
          End If
        Next
        de.rsKorisnici.Requery
        de.rsKorZanimanjaVeze.Requery
        de.rsPoslovi.Requery
        de.rsPosloviBKP.Requery
        de.rsTelefoni.Requery
        de.rsZiroRacuni.Requery
      End If
    Case 3
      a = MsgBox("Da li ste sigurni da želite obrisati" & vbCrLf & "odabrane poslove?", vbQuestion + vbYesNo, Me.Caption)
      If a = vbYes Then
        For i = Me.ctlListView.ListItems.Count To 1 Step -1
          Set c = Me.ctlListView.ListItems(i)
          If c.Selected Then
            de.rsPoslovi.MoveFirst
            de.rsPoslovi.Find "ID=" & Mid(c.Key, 3)
            bFinished = False
            Select Case de.rsPoslovi!IntervalID
              Case 2
                dFinDate = de.rsPoslovi!Datum + de.rsPoslovi!Kolicina
                If Date > dFinDate Then bFinished = True
              Case 3
                dFinDate = de.rsPoslovi!Datum + (de.rsPoslovi!Kolicina * 7)
                If Date > dFinDate Then bFinished = True
              Case 4
                dFinDate = DateSerial(Year(de.rsPoslovi!Datum), Month(de.rsPoslovi!Datum) + de.rsPoslovi!Kolicina, Day(de.rsPoslovi!Datum))
                If Date > dFinDate Then bFinished = True
              Case 5
                dFinDate = DateSerial(Year(de.rsPoslovi!Datum) + de.rsPoslovi!Kolicina, Month(de.rsPoslovi!Datum), Day(de.rsPoslovi!Datum))
                If Date > dFinDate Then bFinished = True
            End Select
            de.rsDavaoci.MoveFirst
            de.rsDavaoci.Find "RB=" & de.rsPoslovi!DavalacRB
            If bFinished Then
              de.rsDavaoci!DatPosljednjegPosla = dFinDate
            Else
              de.rsDavaoci!DatPosljednjegPosla = Date
            End If
            de.rsPosloviBKP.AddNew
            de.rsPosloviBKP!OldID = de.rsPoslovi!ID
            de.rsPosloviBKP!DavalacRB = de.rsPoslovi!DavalacRB
            de.rsPosloviBKP!KorisnikRB = de.rsPoslovi!KorisnikRB
            de.rsPosloviBKP!Datum = de.rsPoslovi!Datum
            de.rsPosloviBKP!Iznos = de.rsPoslovi!Iznos
            de.rsPosloviBKP!IntervalID = de.rsPoslovi!IntervalID
            de.rsPosloviBKP!Kolicina = de.rsPoslovi!Kolicina
            de.rsPosloviBKP.Update
            de.cn1.Execute "DELETE FROM Poslovi WHERE ID=" & Mid(c.Key, 3)
            Me.ctlListView.ListItems.Remove c.Key
          End If
        Next
        de.rsPoslovi.Requery
      End If
  End Select
  ctlListView_Click
End Sub

Private Sub cmdEdit_Click()
  Select Case Me.ctlTabStrip.SelectedItem.Index
    Case 1
      frmDavalac.Edit
      If frmDavalac.OKPressed Then
        Me.ctlListView.SelectedItem.Key = "ID" & frmDavalac.txtRB
        Me.ctlListView.SelectedItem.Text = frmDavalac.txtRB
        Me.ctlListView.SelectedItem.SubItems(1) = UCase(frmDavalac.txtPrezime)
        Me.ctlListView.SelectedItem.SubItems(2) = frmDavalac.txtIme
        Me.ctlListView.SelectedItem.SubItems(3) = frmDavalac.txtRoditelj
        Me.ctlListView.SelectedItem.SubItems(4) = frmDavalac.txtAdresa
        Me.ctlListView.SelectedItem.SubItems(5) = frmDavalac.txtMjesto
        If Not IsNull(frmDavalac.ctlDatumPP.Value) Then
          Me.ctlListView.SelectedItem.SubItems(6) = frmDavalac.ctlDatumPP.Value
        Else
          Me.ctlListView.SelectedItem.SubItems(6) = ""
        End If
      End If
    Case 2
      frmKorisnik.Edit
      If frmKorisnik.OKPressed Then
        Me.ctlListView.SelectedItem.Key = "ID" & frmKorisnik.txtRB
        Me.ctlListView.SelectedItem.Text = frmKorisnik.txtRB
        Me.ctlListView.SelectedItem.SubItems(1) = UCase(frmKorisnik.txtPrezime)
        If frmKorisnik.cboTip.ListIndex = 1 Then
          Me.ctlListView.SelectedItem.SubItems(2) = frmKorisnik.txtIme
        Else
          Me.ctlListView.SelectedItem.SubItems(2) = ""
        End If
        Me.ctlListView.SelectedItem.SubItems(3) = frmKorisnik.txtAdresa
        Me.ctlListView.SelectedItem.SubItems(4) = frmKorisnik.txtMjesto
        If frmKorisnik.cboTip.ListIndex = 0 Then
          Me.ctlListView.SelectedItem.SubItems(5) = frmKorisnik.txtKontaktOsoba
          Me.ctlListView.SelectedItem.SubItems(6) = frmKorisnik.txtKontaktTel
        Else
          Me.ctlListView.SelectedItem.SubItems(5) = ""
          Me.ctlListView.SelectedItem.SubItems(6) = ""
        End If
      End If
    Case 3
      frmPosao.Edit
      If frmPosao.OKPressed Then
        Me.ctlListView.SelectedItem.Key = "ID" & frmPosao.txtRB
        Me.ctlListView.SelectedItem.Text = frmPosao.txtRB
        Me.ctlListView.SelectedItem.SubItems(1) = frmPosao.cboDavalac.List(frmPosao.cboDavalac.ListIndex)
        Me.ctlListView.SelectedItem.SubItems(2) = frmPosao.cboKorisnik.List(frmPosao.cboKorisnik.ListIndex)
        Me.ctlListView.SelectedItem.SubItems(3) = frmPosao.ctlDatum.Value
        Me.ctlListView.SelectedItem.SubItems(4) = Format(Val(frmPosao.txtIznos), "#,##0.00 KM")
        Me.ctlListView.SelectedItem.SubItems(5) = frmPosao.cboInterval.List(frmPosao.cboInterval.ListIndex)
        If frmPosao.cboInterval.ListIndex = 0 Then
          Me.ctlListView.SelectedItem.SubItems(6) = 0
        Else
          Me.ctlListView.SelectedItem.SubItems(6) = Val(frmPosao.txtKolicina)
        End If
        ApplyIcons
      End If
  End Select
  ctlListView_Click
End Sub

Private Sub cmdPrint_Click()
  frmPrint.Display
  If frmPrint.OKPressed Then
    PodesiStampac
    Stampaj
  End If
End Sub

Private Sub PodesiStampac()
  Dim pTMP As Printer
  For Each pTMP In Printers
    If pTMP.DeviceName = frmPrint.cboPrinter.Text Then
      Set Printer = pTMP
      Exit For
    End If
  Next
  Printer.Font.Name = "Arial Narrow"
  Printer.Font.Charset = 238
  Printer.Font.Size = 9
  Printer.Orientation = vbPRORPortrait
  Printer.PaperSize = vbPRPSA4
  Printer.ScaleMode = vbMillimeters
End Sub

Private Sub Stampaj()
  Dim x As Single, y As Single, sStop As Single
  Dim cIznos As Currency, cIznosSum As Currency
  Dim dBegin As Date, dEnd As Date, dDate As Date
  Dim i As Integer
  
  x = 15
  y = 10
  dBegin = DateSerial(Year(Date), Month(Date) - 1, 1)
  dEnd = DateSerial(Year(Date), Month(Date), 1) - 1
  Ispisi x + 90, y, "LISTA ISPLAÆENIH POSLOVA ZA PERIOD " & Format(DateSerial(Year(Date), Month(Date) - 1, 1), "d. M.") & " DO " & Format(DateSerial(Year(Date), Month(Date), 1) - 1, "d. M. yyyy."), True, , True, 2
  Printer.CurrentY = Printer.CurrentY + 5
  sStop = Printer.CurrentY
  Ispisi x + 10, sStop, "R. Br.", True, , , 3
  Ispisi x + 17, sStop, "Davalac usluge", True
  Ispisi x + 66, sStop, "Korisnik usluge", True
  Ispisi x + 116, sStop, "KM/Int.", True, , , 3
  Ispisi x + 135, sStop, "Interval", True, , , 2
  Ispisi x + 165, sStop, "Iznos", True, , , 3
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (x, Printer.CurrentY)-(x + 165, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  If de.rsPosloviBKP.RecordCount > 0 Then
    de.rsPosloviBKP.MoveFirst
    Do Until de.rsPosloviBKP.EOF
      de.rsDavaoci.MoveFirst
      de.rsDavaoci.Find "RB=" & de.rsPosloviBKP!DavalacRB
      If Not de.rsDavaoci.EOF Then
        de.rsKorisnici.MoveFirst
        de.rsKorisnici.Find "RB=" & de.rsPosloviBKP!KorisnikRB
        If Not de.rsKorisnici.EOF Then
          de.rsIntervali.MoveFirst
          de.rsIntervali.Find "ID=" & de.rsPosloviBKP!IntervalID
          Select Case de.rsIntervali!ID
            Case 1
              If (DateSerial(Year(de.rsPosloviBKP!Datum), Month(de.rsPosloviBKP!Datum) + 1, Day(de.rsPosloviBKP!Datum)) >= dBegin) And (DateSerial(Year(de.rsPosloviBKP!Datum), Month(de.rsPosloviBKP!Datum) + 1, Day(de.rsPosloviBKP!Datum)) <= dEnd) Then cIznos = de.rsPosloviBKP!Iznos
            Case 2
              For i = 1 To de.rsPosloviBKP!Kolicina
                If (de.rsPosloviBKP!Datum + i >= dBegin) And (de.rsPosloviBKP!Datum + i <= dEnd) Then cIznos = cIznos + de.rsPosloviBKP!Iznos
              Next
            Case 3
              For i = 1 To de.rsPosloviBKP!Kolicina
                If (de.rsPosloviBKP!Datum + (7 + i) >= dBegin) And (de.rsPosloviBKP!Datum + (7 * i) <= dEnd) Then cIznos = cIznos + de.rsPosloviBKP!Iznos
              Next
            Case 4
              For i = 1 To de.rsPosloviBKP!Kolicina
                If (DateSerial(Year(de.rsPosloviBKP!Datum), Month(de.rsPosloviBKP!Datum) + i, Day(de.rsPosloviBKP!Datum)) >= dBegin) And (DateSerial(Year(de.rsPosloviBKP!Datum), Month(de.rsPosloviBKP!Datum) + i, Day(de.rsPosloviBKP!Datum)) <= dEnd) Then
                  cIznos = de.rsPosloviBKP!Iznos
                  Exit For
                End If
              Next
            Case 5
              For i = 1 To de.rsPosloviBKP!Kolicina
                If (DateSerial(Year(de.rsPosloviBKP!Datum) + i, Month(de.rsPosloviBKP!Datum), Day(de.rsPosloviBKP!Datum)) >= dBegin) And (DateSerial(Year(de.rsPosloviBKP!Datum) + i, Month(de.rsPosloviBKP!Datum), Day(de.rsPosloviBKP!Datum)) <= dEnd) Then
                  cIznos = de.rsPosloviBKP!Iznos
                  Exit For
                End If
              Next
          End Select
          If cIznos > 0 Then
            If Printer.CurrentY + Printer.TextHeight("0") > Printer.ScaleHeight Then
              Printer.NewPage
              Ispisi x + 90, y, "LISTA ISPLAÆENIH POSLOVA ZA PERIOD " & Format(DateSerial(Year(Date), Month(Date) - 1, 1), "d. M.") & " DO " & Format(DateSerial(Year(Date), Month(Date), 1) - 1, "d. M. yyyy."), True, , True, 2
              Printer.CurrentY = Printer.CurrentY + 5
              sStop = Printer.CurrentY
              Ispisi x + 10, sStop, "R. Br.", True, , , 3
              Ispisi x + 17, sStop, "Davalac usluge", True
              Ispisi x + 66, sStop, "Korisnik usluge", True
              Ispisi x + 116, sStop, "KM/Int.", True, , , 3
              Ispisi x + 135, sStop, "Interval", True, , , 2
              Ispisi x + 165, sStop, "Iznos", True, , , 3
              Printer.CurrentY = Printer.CurrentY + 1
              Printer.Line (x, Printer.CurrentY)-(x + 165, Printer.CurrentY)
              Printer.CurrentY = Printer.CurrentY + 1
            End If
            sStop = Printer.CurrentY
            Ispisi x + 10, sStop, de.rsPosloviBKP!OldID, , , , 3
            Ispisi x + 17, sStop, UCase(de.rsDavaoci!Prezime) & ", " & de.rsDavaoci!Ime
            If de.rsKorisnici!FL Then
              Ispisi x + 66, sStop, UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Ime
            Else
              Ispisi x + 66, sStop, UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Mjesto
            End If
            Ispisi x + 116, sStop, Format(de.rsPosloviBKP!Iznos, "#,##0.00"), , , , 3
            Ispisi x + 135, sStop, de.rsIntervali!Interval, , , , 2
            Ispisi x + 165, sStop, Format(cIznos, "#,##0.00"), , , , 3
            cIznosSum = cIznosSum + cIznos
            cIznos = 0
          End If
        End If
      End If
      de.rsPosloviBKP.MoveNext
    Loop
  End If
  If de.rsPoslovi.RecordCount > 0 Then
    de.rsPoslovi.MoveFirst
    Do Until de.rsPoslovi.EOF
      de.rsDavaoci.MoveFirst
      de.rsDavaoci.Find "RB=" & de.rsPoslovi!DavalacRB
      If Not de.rsDavaoci.EOF Then
        de.rsKorisnici.MoveFirst
        de.rsKorisnici.Find "RB=" & de.rsPoslovi!KorisnikRB
        If Not de.rsKorisnici.EOF Then
          de.rsIntervali.MoveFirst
          de.rsIntervali.Find "ID=" & de.rsPoslovi!IntervalID
          Select Case de.rsIntervali!ID
            Case 1
              If (DateSerial(Year(de.rsPoslovi!Datum), Month(de.rsPoslovi!Datum) + 1, Day(de.rsPoslovi!Datum)) >= dBegin) And (DateSerial(Year(de.rsPoslovi!Datum), Month(de.rsPoslovi!Datum) + 1, Day(de.rsPoslovi!Datum)) <= dEnd) Then cIznos = de.rsPoslovi!Iznos
            Case 2
              For i = 1 To de.rsPoslovi!Kolicina
                If (de.rsPoslovi!Datum + i >= dBegin) And (de.rsPoslovi!Datum + i <= dEnd) Then cIznos = cIznos + de.rsPoslovi!Iznos
              Next
            Case 3
              For i = 1 To de.rsPoslovi!Kolicina
                If (de.rsPoslovi!Datum + (7 + i) >= dBegin) And (de.rsPoslovi!Datum + (7 * i) <= dEnd) Then cIznos = cIznos + de.rsPoslovi!Iznos
              Next
            Case 4
              For i = 1 To de.rsPoslovi!Kolicina
                If (DateSerial(Year(de.rsPoslovi!Datum), Month(de.rsPoslovi!Datum) + i, Day(de.rsPoslovi!Datum)) >= dBegin) And (DateSerial(Year(de.rsPoslovi!Datum), Month(de.rsPoslovi!Datum) + i, Day(de.rsPoslovi!Datum)) <= dEnd) Then
                  cIznos = de.rsPoslovi!Iznos
                  Exit For
                End If
              Next
            Case 5
              For i = 1 To de.rsPoslovi!Kolicina
                If (DateSerial(Year(de.rsPoslovi!Datum) + i, Month(de.rsPoslovi!Datum), Day(de.rsPoslovi!Datum)) >= dBegin) And (DateSerial(Year(de.rsPoslovi!Datum) + i, Month(de.rsPoslovi!Datum), Day(de.rsPoslovi!Datum)) <= dEnd) Then
                  cIznos = de.rsPoslovi!Iznos
                  Exit For
                End If
              Next
          End Select
          If cIznos > 0 Then
            If Printer.CurrentY + Printer.TextHeight("0") > Printer.ScaleHeight Then
              Printer.NewPage
              Ispisi x + 90, y, "LISTA ISPLAÆENIH POSLOVA ZA PERIOD " & Format(DateSerial(Year(Date), Month(Date) - 1, 1), "d. M.") & " DO " & Format(DateSerial(Year(Date), Month(Date), 1) - 1, "d. M. yyyy."), True, , True, 2
              Printer.CurrentY = Printer.CurrentY + 5
              sStop = Printer.CurrentY
              Ispisi x + 10, sStop, "R. Br.", True, , , 3
              Ispisi x + 17, sStop, "Davalac usluge", True
              Ispisi x + 66, sStop, "Korisnik usluge", True
              Ispisi x + 116, sStop, "KM/Int.", True, , , 3
              Ispisi x + 135, sStop, "Interval", True, , , 2
              Ispisi x + 165, sStop, "Iznos", True, , , 3
              Printer.CurrentY = Printer.CurrentY + 1
              Printer.Line (x, Printer.CurrentY)-(x + 165, Printer.CurrentY)
              Printer.CurrentY = Printer.CurrentY + 1
            End If
            sStop = Printer.CurrentY
            Ispisi x + 10, sStop, de.rsPoslovi!ID, , , , 3
            Ispisi x + 17, sStop, UCase(de.rsDavaoci!Prezime) & ", " & de.rsDavaoci!Ime
            If de.rsKorisnici!FL Then
              Ispisi x + 66, sStop, UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Ime
            Else
              Ispisi x + 66, sStop, UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Mjesto
            End If
            Ispisi x + 116, sStop, Format(de.rsPoslovi!Iznos, "#,##0.00"), , , , 3
            Ispisi x + 135, sStop, de.rsIntervali!Interval, , , , 2
            Ispisi x + 165, sStop, Format(cIznos, "#,##0.00"), , , , 3
            cIznosSum = cIznosSum + cIznos
            cIznos = 0
          End If
        End If
      End If
      de.rsPoslovi.MoveNext
    Loop
  End If
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (x, Printer.CurrentY)-(x + 165, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  Ispisi x + 165, Printer.CurrentY, Format(cIznosSum, "#,##0.00"), True, , , 3
  cIznosSum = 0
  Printer.EndDoc
End Sub

Private Sub Ispisi(Optional x As Single = -1, Optional y As Single = -1, Optional sText As String = "", Optional bBold As Boolean = False, Optional bItalic As Boolean = False, Optional bUnder As Boolean = False, Optional iAlign As Integer = 1, Optional iVAlign As Integer = 1)
    
    If x = -1 Then x = Printer.CurrentX
    If y = -1 Then y = Printer.CurrentY
    
    Printer.FontBold = bBold
    Printer.FontItalic = bItalic
    Printer.FontUnderline = bUnder
    Select Case iAlign
        Case 1
            Printer.CurrentX = x
        Case 2
            Printer.CurrentX = x - Printer.TextWidth(sText) / 2
        Case 3
            Printer.CurrentX = x - Printer.TextWidth(sText)
    End Select
    Select Case iVAlign
        Case 1
            Printer.CurrentY = y
        Case 2
            Printer.CurrentY = y - Printer.TextHeight(sText) / 2
        Case 3
            Printer.CurrentY = y - Printer.TextHeight(sText)
    End Select
    Printer.Print sText
End Sub


Private Sub cmdSearch_Click()
  Dim i As Integer
  If Me.cmdSearch.Caption = "Traži" Then
    If Trim(Me.txtSearch) = "" Then Exit Sub
    Me.txtSearch = Trim(Me.txtSearch)
    Select Case Me.ctlTabStrip.SelectedItem.Index
      Case 1
        TraziDavaoce
      Case 2
        TraziKorisnike
      Case 3
        TraziPoslove
    End Select
    Me.cmdSearch.Caption = "Poništi"
    Me.ctlListView.Tag = ""
    ApplyIcons
  Else
    Me.cmdSearch.Caption = "Traži"
    i = Me.cboSearch.ListIndex
    FillList Me.ctlTabStrip.SelectedItem.Index
    Me.cboSearch.ListIndex = i
  End If
End Sub

Private Sub TraziDavaoce()
  If de.rsDavaoci.RecordCount > 0 Then
    de.rsDavaoci.MoveFirst
    Select Case Me.cboSearch.ListIndex
      Case 0 ' datum prijave
        If Not IsDate(Me.txtSearch) Then
          MsgBox "Unesite pravilan datum!", vbExclamation + vbOKOnly, Me.Caption
          Exit Sub
        End If
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "Datum=#" & Format(CDate(Me.txtSearch), "yyyy-MM-dd") & "#"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
          de.rsDavaoci.Find "Datum=#" & Format(CDate(Me.txtSearch), "yyyy-MM-dd") & "#", 1
        Loop
      Case 1 ' ime
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "Ime LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
        de.rsDavaoci.Find "Ime LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 2 ' prezime
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "Prezime LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
        de.rsDavaoci.Find "Prezime LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 3 ' ime roditelja
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "Roditelj LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
        de.rsDavaoci.Find "Roditelj LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 4 ' zanimanja
        Me.ctlListView.ListItems.Clear
        Do Until de.rsDavaoci.EOF
          de.rsZanimanjaVeze.MoveFirst
          de.rsZanimanjaVeze.Find "RB=" & de.rsDavaoci!RB
          Do While Not de.rsZanimanjaVeze.EOF
            de.rsZanimanja.MoveFirst
            de.rsZanimanja.Find "ID=" & de.rsZanimanjaVeze!ZanimanjeID
            If InStr(LCase(de.rsZanimanja!Naziv), LCase(Me.txtSearch)) > 0 Then
              Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
              Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
              Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
              Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
              Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
              Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
              If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
              Exit Do
            End If
            de.rsZanimanjaVeze.Find "RB=" & de.rsDavaoci!RB, 1
          Loop
          de.rsDavaoci.MoveNext
        Loop
      Case 5 ' raspoloživost
        Me.ctlListView.ListItems.Clear
        If de.rsRaspolozivost.RecordCount > 0 Then
          Do Until de.rsDavaoci.EOF
            de.rsRaspolozivostVeze.MoveFirst
            de.rsRaspolozivostVeze.Find "RB=" & de.rsDavaoci!RB
            Do While Not de.rsRaspolozivostVeze.EOF
              de.rsRaspolozivost.MoveFirst
              de.rsRaspolozivost.Find "ID=" & de.rsRaspolozivostVeze!RaspolozivostID
              If InStr(LCase(de.rsRaspolozivost!Naziv), LCase(Me.txtSearch)) > 0 Then
                Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
                If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
                Exit Do
              End If
              de.rsRaspolozivostVeze.Find "RB=" & de.rsDavaoci!RB, 1
            Loop
            de.rsDavaoci.MoveNext
          Loop
        End If
      Case 6 ' adresa
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "Adresa LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
        de.rsDavaoci.Find "Adresa LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 7 ' mjesto
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "Mjesto LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
        de.rsDavaoci.Find "Mjesto LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 8 ' telefon
        Me.ctlListView.ListItems.Clear
        If de.rsTelefoni.RecordCount > 0 Then
          Do Until de.rsDavaoci.EOF
            de.rsTelefoni.MoveFirst
            de.rsTelefoni.Find "RB=" & de.rsDavaoci!RB
            Do While Not de.rsTelefoni.EOF
              If (Not de.rsTelefoni!Korisnik) And InStr(LCase(de.rsTelefoni!Telefon), LCase(Me.txtSearch)) > 0 Then
                Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
                Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
                If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
                Exit Do
              End If
              de.rsTelefoni.Find "RB=" & de.rsDavaoci!RB, 1
            Loop
            de.rsDavaoci.MoveNext
          Loop
        End If
      Case 9 ' JMB
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "JMB LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
        de.rsDavaoci.Find "JMB LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 10 ' LK
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "BRLK LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
        de.rsDavaoci.Find "BRLK LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 11 ' Komentar
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "Komentar LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
        de.rsDavaoci.Find "Komentar LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 12 ' Datum posljednjeg posla
        If Not IsDate(Me.txtSearch) Then
          MsgBox "Unesite pravilan datum!", vbExclamation + vbOKOnly, Me.Caption
          Exit Sub
        End If
        Me.ctlListView.ListItems.Clear
        de.rsDavaoci.Find "DatPosljednjegPosla=#" & Format(CDate(Me.txtSearch), "yyyy-MM-dd") & "#"
        Do While Not de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
          de.rsDavaoci.Find "DatPosljednjegPosla=#" & Format(CDate(Me.txtSearch), "yyyy-MM-dd") & "#", 1
        Loop
    End Select
  End If
End Sub

Private Sub TraziKorisnike()
  If de.rsKorisnici.RecordCount > 0 Then
    de.rsKorisnici.MoveFirst
    Select Case Me.cboSearch.ListIndex
      Case 0 ' naziv (prezime)
        Me.ctlListView.ListItems.Clear
        de.rsKorisnici.Find "NazivPrezime LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsKorisnici.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsKorisnici.Find "NazivPrezime LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 1 ' (ime)
        Me.ctlListView.ListItems.Clear
        de.rsKorisnici.Find "Ime LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsKorisnici.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsKorisnici.Find "Ime LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 2 ' adresa
        Me.ctlListView.ListItems.Clear
        de.rsKorisnici.Find "Adresa LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsKorisnici.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsKorisnici.Find "Adresa LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 3 ' mjesto
        Me.ctlListView.ListItems.Clear
        de.rsKorisnici.Find "Mjesto LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsKorisnici.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsKorisnici.Find "Mjesto LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 4 ' telefon
        Me.ctlListView.ListItems.Clear
        If de.rsTelefoni.RecordCount > 0 Then
          Do Until de.rsKorisnici.EOF
            de.rsTelefoni.MoveFirst
            de.rsTelefoni.Find "RB=" & de.rsKorisnici!RB
            Do While Not de.rsTelefoni.EOF
              If de.rsTelefoni!Korisnik And InStr(LCase(de.rsTelefoni!Telefon), LCase(Me.txtSearch)) > 0 Then
                Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
                Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
                If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
                Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
                Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
                If Not de.rsKorisnici!FL Then
                  If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
                  If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
                End If
                Exit Do
              End If
              de.rsTelefoni.Find "RB=" & de.rsKorisnici!RB, 1
            Loop
            de.rsKorisnici.MoveNext
          Loop
        End If
      Case 5 ' tražena zanimanja
        Me.ctlListView.ListItems.Clear
        Do Until de.rsKorisnici.EOF
          de.rsKorZanimanjaVeze.MoveFirst
          de.rsKorZanimanjaVeze.Find "RB=" & de.rsKorisnici!RB
          Do While Not de.rsKorZanimanjaVeze.EOF
            de.rsZanimanja.MoveFirst
            de.rsZanimanja.Find "ID=" & de.rsKorZanimanjaVeze!ZanimanjeID
            If InStr(LCase(de.rsZanimanja!Naziv), LCase(Me.txtSearch)) > 0 Then
              Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
              Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
              If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
              Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
              Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
              If Not de.rsKorisnici!FL Then
                If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
                If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
              End If
              Exit Do
            End If
            de.rsKorZanimanjaVeze.Find "RB=" & de.rsKorisnici!RB, 1
          Loop
          de.rsKorisnici.MoveNext
        Loop
      Case 6 ' ziro-racun
        Me.ctlListView.ListItems.Clear
        If de.rsZiroRacuni.RecordCount > 0 Then
          Do Until de.rsKorisnici.EOF
            de.rsZiroRacuni.MoveFirst
            de.rsZiroRacuni.Find "RB=" & de.rsKorisnici!RB
            Do While Not de.rsZiroRacuni.EOF
              If InStr(LCase(de.rsZiroRacuni!ZR), LCase(Me.txtSearch)) > 0 Then
                Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
                Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
                If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
                Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
                Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
                If Not de.rsKorisnici!FL Then
                  If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
                  If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
                End If
                Exit Do
              End If
              de.rsZiroRacuni.Find "RB=" & de.rsKorisnici!RB, 1
            Loop
            de.rsKorisnici.MoveNext
          Loop
        End If
      Case 7 ' kontakt osoba
        Me.ctlListView.ListItems.Clear
        de.rsKorisnici.Find "Kontakt LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsKorisnici.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsKorisnici.Find "Kontakt LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 8 ' kontakt tel
        Me.ctlListView.ListItems.Clear
        de.rsKorisnici.Find "KontaktTel LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsKorisnici.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsKorisnici.Find "KontaktTel LIKE '*" & Me.txtSearch & "*'", 1
        Loop
      Case 9 ' komentar
        Me.ctlListView.ListItems.Clear
        de.rsKorisnici.Find "Komentar LIKE '*" & Me.txtSearch & "*'"
        Do While Not de.rsKorisnici.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsKorisnici.Find "Komentar LIKE '*" & Me.txtSearch & "*'", 1
        Loop
    End Select
  End If
End Sub

Private Sub TraziPoslove()
  If de.rsPoslovi.RecordCount > 0 Then
    de.rsPoslovi.MoveFirst
    Select Case Me.cboSearch.ListIndex
      Case 0 ' davalac
        Me.ctlListView.ListItems.Clear
        Do Until de.rsPoslovi.EOF
          de.rsDavaoci.MoveFirst
          de.rsDavaoci.Find "RB=" & de.rsPoslovi!DavalacRB
          If (InStr(LCase(de.rsDavaoci!Prezime), LCase(Me.txtSearch)) > 0) Or (InStr(LCase(de.rsDavaoci!Ime), LCase(Me.txtSearch)) > 0) Then
            Me.ctlListView.ListItems.Add , "ID" & de.rsPoslovi!ID, de.rsPoslovi!ID
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(1) = UCase(de.rsDavaoci!Prezime) & ", " & de.rsDavaoci!Ime
            de.rsKorisnici.MoveFirst
            de.rsKorisnici.Find "RB=" & de.rsPoslovi!KorisnikRB
            If de.rsKorisnici!FL Then
              Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Ime
            Else
              Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Mjesto
            End If
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(3) = de.rsPoslovi!Datum
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(4) = Format(de.rsPoslovi!Iznos, "#,##0.00 KM")
            de.rsIntervali.MoveFirst
            de.rsIntervali.Find "ID=" & de.rsPoslovi!IntervalID
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(5) = de.rsIntervali!Interval
            If Not IsNull(de.rsPoslovi!Kolicina) Then Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(6) = de.rsPoslovi!Kolicina
          End If
          de.rsPoslovi.MoveNext
        Loop
      Case 1 ' korisnik
        Me.ctlListView.ListItems.Clear
        Do Until de.rsPoslovi.EOF
          de.rsKorisnici.MoveFirst
          de.rsKorisnici.Find "RB=" & de.rsPoslovi!KorisnikRB
          If (InStr(LCase(de.rsKorisnici!NazivPrezime), LCase(Me.txtSearch)) > 0) Or (InStr(LCase(de.rsKorisnici!Ime), LCase(Me.txtSearch)) > 0) Then
            Me.ctlListView.ListItems.Add , "ID" & de.rsPoslovi!ID, de.rsPoslovi!ID
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(1) = UCase(de.rsDavaoci!Prezime) & ", " & de.rsDavaoci!Ime
            de.rsKorisnici.MoveFirst
            de.rsKorisnici.Find "RB=" & de.rsPoslovi!KorisnikRB
            If de.rsKorisnici!FL Then
              Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Ime
            Else
              Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Mjesto
            End If
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(3) = de.rsPoslovi!Datum
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(4) = Format(de.rsPoslovi!Iznos, "#,##0.00 KM")
            de.rsIntervali.MoveFirst
            de.rsIntervali.Find "ID=" & de.rsPoslovi!IntervalID
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(5) = de.rsIntervali!Interval
            If Not IsNull(de.rsPoslovi!Kolicina) Then Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(6) = de.rsPoslovi!Kolicina
          End If
          de.rsPoslovi.MoveNext
        Loop
      Case 2 ' datum
        If Not IsDate(Me.txtSearch) Then
          MsgBox "Unesite pravilan datum!", vbExclamation + vbOKOnly, Me.Caption
          Exit Sub
        End If
        Me.ctlListView.ListItems.Clear
        de.rsPoslovi.Find "Datum=#" & Me.txtSearch & "#"
        Do While Not de.rsPoslovi.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsPoslovi.Find "Datum=#" & Me.txtSearch & "#", 1
        Loop
      Case 3 ' iznos
        If Not IsNumeric(Me.txtSearch) Then
          MsgBox "Unesite pravilan iznos!", vbExclamation + vbOKOnly, Me.Caption
          Exit Sub
        End If
        Me.ctlListView.ListItems.Clear
        de.rsPoslovi.Find "Iznos=" & Me.txtSearch
        Do While Not de.rsPoslovi.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsPoslovi.Find "Iznos=" & Me.txtSearch, 1
        Loop
      Case 4 ' interval
        Me.ctlListView.ListItems.Clear
        Do Until de.rsPoslovi.EOF
          de.rsIntervali.MoveFirst
          de.rsIntervali.Find "RB=" & de.rsPoslovi!IntervalID
          If (InStr(LCase(de.rsIntervali!Interval), LCase(Me.txtSearch)) > 0) Then
            Me.ctlListView.ListItems.Add , "ID" & de.rsPoslovi!ID, de.rsPoslovi!ID
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(1) = UCase(de.rsDavaoci!Prezime) & ", " & de.rsDavaoci!Ime
            de.rsKorisnici.MoveFirst
            de.rsKorisnici.Find "RB=" & de.rsPoslovi!KorisnikRB
            If de.rsKorisnici!FL Then
              Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Ime
            Else
              Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Mjesto
            End If
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(3) = de.rsPoslovi!Datum
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(4) = Format(de.rsPoslovi!Iznos, "#,##0.00 KM")
            de.rsIntervali.MoveFirst
            de.rsIntervali.Find "ID=" & de.rsPoslovi!IntervalID
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(5) = de.rsIntervali!Interval
            If Not IsNull(de.rsPoslovi!Kolicina) Then Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(6) = de.rsPoslovi!Kolicina
          End If
          de.rsPoslovi.MoveNext
        Loop
      Case 5 ' kolièina
        If Not IsNumeric(Me.txtSearch) Then
          MsgBox "Unesite pravilnu kolièinu!", vbExclamation + vbOKOnly, Me.Caption
          Exit Sub
        End If
        Me.ctlListView.ListItems.Clear
        de.rsPoslovi.Find "Kolicina=" & Me.txtSearch
        Do While Not de.rsPoslovi.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not de.rsKorisnici!FL Then
            If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
            If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          End If
          de.rsPoslovi.Find "Kolicina=" & Me.txtSearch, 1
        Loop
    End Select
  End If
End Sub

Private Sub ctlListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Me.ctlListView.Sorted = True
  If Me.ctlListView.SortKey = ColumnHeader.Index - 1 Then
    If Me.ctlListView.SortOrder = lvwAscending Then
      Me.ctlListView.SortOrder = lvwDescending
    Else
      Me.ctlListView.SortOrder = lvwAscending
    End If
  Else
    Me.ctlListView.SortOrder = lvwAscending
    Me.ctlListView.SortKey = ColumnHeader.Index - 1
  End If
End Sub

Private Sub ctlListView_Click()
  Dim Item As MSComctlLib.ListItem
  Dim sel As Boolean
  Dim sStatus As String
  sel = False
  
  For Each Item In Me.ctlListView.ListItems
    If Item.Selected Then
      sel = True
      Exit For
    End If
  Next
  If sel Then
    Me.cmdEdit.Enabled = True
    Me.cmdDel.Enabled = True
  Else
    Me.cmdEdit.Enabled = False
    Me.cmdDel.Enabled = False
  End If
End Sub

Private Sub ctlTabStrip_Click()
  Dim bPoslovi As Boolean
  FillList Me.ctlTabStrip.SelectedItem.Index
  bPoslovi = bPoslovi Or (de.rsPoslovi.RecordCount > 0)
  bPoslovi = bPoslovi Or (de.rsPosloviBKP.RecordCount > 0)
  If Me.ctlTabStrip.SelectedItem.Index = 3 Then
    Me.cmdPrint.Visible = True
  Else
    Me.cmdPrint.Visible = False
  End If
  Me.cmdPrint.Enabled = bPoslovi
End Sub

Private Sub Form_Load()
  Dim w As Integer
  Dim h As Integer
  Dim s As Integer
  w = CInt(GetSetting(Me.Caption, "Main", "Window Width", "7000"))
  h = CInt(GetSetting(Me.Caption, "Main", "Window Height", "5000"))
  s = CInt(GetSetting(Me.Caption, "Main", "Window State", "2"))
  If w < 7000 Then w = 7000
  If h < 5000 Then h = 5000
  If s = 1 Then s = 2
  Me.Width = w
  Me.Height = h
  Me.WindowState = s
  
  de.cn1.Open
  
  If de.rsDavaoci.State = adStateClosed Then de.Davaoci
  If de.rsIntervali.State = adStateClosed Then de.Intervali
  If de.rsKorisnici.State = adStateClosed Then de.Korisnici
  If de.rsKorZanimanjaVeze.State = adStateClosed Then de.KorZanimanjaVeze
  If de.rsPoslovi.State = adStateClosed Then de.Poslovi
  If de.rsPosloviBKP.State = adStateClosed Then de.PosloviBKP
  If de.rsRaspolozivost.State = adStateClosed Then de.Raspolozivost
  If de.rsRaspolozivostVeze.State = adStateClosed Then de.RaspolozivostVeze
  If de.rsTelefoni.State = adStateClosed Then de.Telefoni
  If de.rsZanimanja.State = adStateClosed Then de.Zanimanja
  If de.rsZanimanjaVeze.State = adStateClosed Then de.ZanimanjaVeze
  If de.rsZiroRacuni.State = adStateClosed Then de.ZiroRacuni
  
  
  FillList Me.ctlTabStrip.SelectedItem.Index
  
End Sub

Private Sub Form_Resize()
  If Me.WindowState = 1 Then Exit Sub
  If Me.Width < 7000 Then Me.Width = 7000
  If Me.Height < 5000 Then Me.Height = 5000
  Me.cmdAdd.Top = Me.ScaleHeight - Me.cmdAdd.Height - Me.cmdAdd.Left - Me.ctlStatusBar.Height
  Me.cmdEdit.Top = Me.cmdAdd.Top
  Me.cmdDel.Top = Me.cmdAdd.Top
  Me.cmdPrint.Top = Me.cmdAdd.Top
  Me.txtSearch.Top = Me.cmdAdd.Top - Me.txtSearch.Height - Me.txtSearch.Left
  Me.cboSearch.Top = Me.txtSearch.Top
  Me.cmdSearch.Top = Me.txtSearch.Top
  Me.ctlTabStrip.Width = Me.ScaleWidth - 2 * Me.ctlTabStrip.Left
  Me.ctlTabStrip.Height = Me.txtSearch.Top - 2 * Me.ctlTabStrip.Top
  Me.ctlListView.Width = Me.ctlTabStrip.Width - 2 * (Me.ctlListView.Left - Me.ctlTabStrip.Left)
  Me.ctlListView.Height = Me.ctlTabStrip.Height - 2 * (Me.ctlListView.Top - Me.ctlTabStrip.Top) + Me.ctlTabStrip.TabFixedHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting Me.Caption, "Main", "Window State", CStr(Me.WindowState)
  Me.WindowState = 0
  SaveSetting Me.Caption, "Main", "Window Width", CStr(Me.Width)
  SaveSetting Me.Caption, "Main", "Window Height", CStr(Me.Height)
End Sub

Private Sub FillList(iType As Integer)
  Dim i As Integer
  
  If Val(Me.ctlListView.Tag) = iType Then Exit Sub
  Me.ctlListView.ListItems.Clear
  Me.ctlListView.ColumnHeaders.Clear
  Me.cboSearch.Clear
  
  Me.cmdSearch.Caption = "Traži"
  
  Select Case iType
    Case 1
      Me.ctlListView.ColumnHeaders.Add , "colRB", "Red. br.", 840
      Me.ctlListView.ColumnHeaders.Add , "colPrezime", "Prezime", 1800
      Me.ctlListView.ColumnHeaders.Add , "colIme", "Ime", 1400
      Me.ctlListView.ColumnHeaders.Add , "colRoditelj", "Ime roditelja", 1400
      Me.ctlListView.ColumnHeaders.Add , "colAdresa", "Adresa", 2800
      Me.ctlListView.ColumnHeaders.Add , "colMjesto", "Mjesto", 1500
      Me.ctlListView.ColumnHeaders.Add , "colDatum", "Posljednji posao", 1335, 2
      
      Me.cboSearch.AddItem "Datum prijave"
      Me.cboSearch.AddItem "Ime"
      Me.cboSearch.AddItem "Prezime"
      Me.cboSearch.AddItem "Ime roditelja"
      Me.cboSearch.AddItem "Zanimanja"
      Me.cboSearch.AddItem "Raspoloživost"
      Me.cboSearch.AddItem "Adresa"
      Me.cboSearch.AddItem "Mjesto"
      Me.cboSearch.AddItem "Telefon"
      Me.cboSearch.AddItem "JMBG"
      Me.cboSearch.AddItem "Broj LK"
      Me.cboSearch.AddItem "Komentar"
      Me.cboSearch.AddItem "Datum posljednjeg posla"
      
      If de.rsDavaoci.RecordCount > 0 Then
        de.rsDavaoci.MoveFirst
        Do Until de.rsDavaoci.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsDavaoci!RB, de.rsDavaoci!RB
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(1) = UCase(de.rsDavaoci!Prezime)
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(2) = de.rsDavaoci!Ime
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(3) = de.rsDavaoci!Roditelj
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(4) = de.rsDavaoci!Adresa
          Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(5) = de.rsDavaoci!Mjesto
          If Not IsNull(de.rsDavaoci!DatPosljednjegPosla) Then
            Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = de.rsDavaoci!DatPosljednjegPosla
          Else
            Me.ctlListView.ListItems("ID" & de.rsDavaoci!RB).SubItems(6) = ""
          End If
          de.rsDavaoci.MoveNext
        Loop
      End If
      Me.ctlListView.Tag = "1"
      Me.ctlListView.ToolTipText = "Lista davalaca usluga"
    Case 2
      Me.ctlListView.ColumnHeaders.Add , "colRB", "Red. br.", 840
      Me.ctlListView.ColumnHeaders.Add , "colNazivPrezime", "Naziv (Prezime)", 1800
      Me.ctlListView.ColumnHeaders.Add , "colIme", "(Ime)", 1800
      Me.ctlListView.ColumnHeaders.Add , "colAdresa", "Adresa", 2800
      Me.ctlListView.ColumnHeaders.Add , "colMjesto", "Mjesto", 1500
      Me.ctlListView.ColumnHeaders.Add , "colKontakt", "Kontakt osoba", 2000
      Me.ctlListView.ColumnHeaders.Add , "colKontaktTel", "Kontakt tel.", 1500
      
      Me.cboSearch.AddItem "Naziv (Prezime)"
      Me.cboSearch.AddItem "(Ime)"
      Me.cboSearch.AddItem "Adresa"
      Me.cboSearch.AddItem "Mjesto"
      Me.cboSearch.AddItem "Telefon"
      Me.cboSearch.AddItem "Tražena zanimanja"
      Me.cboSearch.AddItem "Žiro-raèun"
      Me.cboSearch.AddItem "Kontakt osoba"
      Me.cboSearch.AddItem "Kontakt telefon"
      Me.cboSearch.AddItem "Komentar"
      
      If de.rsKorisnici.RecordCount > 0 Then
        de.rsKorisnici.MoveFirst
        Do Until de.rsKorisnici.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsKorisnici!RB, de.rsKorisnici!RB
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(1) = UCase(de.rsKorisnici!NazivPrezime)
          If de.rsKorisnici!FL Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(2) = de.rsKorisnici!Ime
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(3) = de.rsKorisnici!Adresa
          Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(4) = de.rsKorisnici!Mjesto
          If Not IsNull(de.rsKorisnici!Kontakt) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(5) = de.rsKorisnici!Kontakt
          If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.ctlListView.ListItems("ID" & de.rsKorisnici!RB).SubItems(6) = de.rsKorisnici!KontaktTel
          de.rsKorisnici.MoveNext
        Loop
      End If
      
      Me.ctlListView.Tag = "2"
      Me.ctlListView.ToolTipText = "Lista korisnika usluga"
    Case 3
      Me.ctlListView.ColumnHeaders.Add , "colRB", "Red. br.", 840
      Me.ctlListView.ColumnHeaders.Add , "colDavalac", "Davalac usluge", 2000
      Me.ctlListView.ColumnHeaders.Add , "colKorisnik", "Korisnik usluge", 2000
      Me.ctlListView.ColumnHeaders.Add , "colDatum", "Datum", 1200, 2
      Me.ctlListView.ColumnHeaders.Add , "colIznos", "Iznos", 1000, 1
      Me.ctlListView.ColumnHeaders.Add , "colInterval", "Interval", 1000
      Me.ctlListView.ColumnHeaders.Add , "colKolicina", "Kolièina", 800, 1
      
      Me.cboSearch.AddItem "Davalac usluge"
      Me.cboSearch.AddItem "Korisnik usluge"
      Me.cboSearch.AddItem "Datum"
      Me.cboSearch.AddItem "Iznos"
      Me.cboSearch.AddItem "Interval"
      Me.cboSearch.AddItem "Kolièina"
      
      If de.rsPoslovi.RecordCount > 0 Then
        de.rsPoslovi.MoveFirst
        Do Until de.rsPoslovi.EOF
          Me.ctlListView.ListItems.Add , "ID" & de.rsPoslovi!ID, de.rsPoslovi!ID
          de.rsDavaoci.MoveFirst
          de.rsDavaoci.Find "RB=" & de.rsPoslovi!DavalacRB
          de.rsKorisnici.MoveFirst
          de.rsKorisnici.Find "RB=" & de.rsPoslovi!KorisnikRB
          Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(1) = UCase(de.rsDavaoci!Prezime) & ", " & de.rsDavaoci!Ime
          If de.rsKorisnici!FL Then
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Ime
          Else
            Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(2) = UCase(de.rsKorisnici!NazivPrezime) & ", " & de.rsKorisnici!Mjesto
          End If
          Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(3) = de.rsPoslovi!Datum
          Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(4) = Format(de.rsPoslovi!Iznos, "#,##0.00 KM")
          de.rsIntervali.MoveFirst
          de.rsIntervali.Find "ID=" & de.rsPoslovi!IntervalID
          Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(5) = de.rsIntervali!Interval
          Me.ctlListView.ListItems("ID" & de.rsPoslovi!ID).SubItems(6) = de.rsPoslovi!Kolicina
          de.rsPoslovi.MoveNext
        Loop
      End If
      
      Me.ctlListView.Tag = "3"
      Me.ctlListView.ToolTipText = "Lista ugovorenih poslova"
  End Select
  Me.ctlListView.Sorted = False
  Me.cboSearch.ListIndex = 0
  ctlListView_Click
  ApplyIcons
End Sub

Private Sub SelAll()
  Me.ActiveControl.SelStart = 0
  Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtSearch_GotFocus()
  SelAll
End Sub

Private Sub ApplyIcons()
  Dim bFinished As Boolean, i As Integer
  If Me.ctlListView.ListItems.Count = 0 Then Exit Sub
  If de.rsPoslovi.RecordCount = 0 Then Exit Sub
  Select Case Me.ctlTabStrip.SelectedItem.Index
    Case 1
      For i = 1 To Me.ctlListView.ListItems.Count
        de.rsPoslovi.MoveFirst
        de.rsPoslovi.Find "DavalacRB=" & Mid(Me.ctlListView.ListItems(i).Key, 3)
        If Not de.rsPoslovi.EOF Then
          bFinished = False
          Select Case de.rsPoslovi!IntervalID
            Case 2
              If Date > (de.rsPoslovi!Datum + de.rsPoslovi!Kolicina) Then bFinished = True
            Case 3
              If Date > (de.rsPoslovi!Datum + (de.rsPoslovi!Kolicina * 7)) Then bFinished = True
            Case 4
              If Date > DateSerial(Year(de.rsPoslovi!Datum), Month(de.rsPoslovi!Datum) + de.rsPoslovi!Kolicina, Day(de.rsPoslovi!Datum)) Then bFinished = True
            Case 5
              If Date > DateSerial(Year(de.rsPoslovi!Datum) + de.rsPoslovi!Kolicina, Month(de.rsPoslovi!Datum), Day(de.rsPoslovi!Datum)) Then bFinished = True
          End Select
          If bFinished Then
            Me.ctlListView.ListItems(i).SmallIcon = 2
          Else
            Me.ctlListView.ListItems(i).SmallIcon = 1
          End If
        End If
      Next
    Case 2
      For i = 1 To Me.ctlListView.ListItems.Count
        de.rsPoslovi.MoveFirst
        de.rsPoslovi.Find "KorisnikRB=" & Mid(Me.ctlListView.ListItems(i).Key, 3)
        If Not de.rsPoslovi.EOF Then
          bFinished = False
          Select Case de.rsPoslovi!IntervalID
            Case 2
              If Date > (de.rsPoslovi!Datum + de.rsPoslovi!Kolicina) Then bFinished = True
            Case 3
              If Date > (de.rsPoslovi!Datum + (de.rsPoslovi!Kolicina * 7)) Then bFinished = True
            Case 4
              If Date > DateSerial(Year(de.rsPoslovi!Datum), Month(de.rsPoslovi!Datum) + de.rsPoslovi!Kolicina, Day(de.rsPoslovi!Datum)) Then bFinished = True
            Case 5
              If Date > DateSerial(Year(de.rsPoslovi!Datum) + de.rsPoslovi!Kolicina, Month(de.rsPoslovi!Datum), Day(de.rsPoslovi!Datum)) Then bFinished = True
          End Select
          If bFinished Then
            Me.ctlListView.ListItems(i).SmallIcon = 2
          Else
            Me.ctlListView.ListItems(i).SmallIcon = 1
          End If
        End If
      Next
    Case 3
      For i = 1 To Me.ctlListView.ListItems.Count
        de.rsPoslovi.MoveFirst
        de.rsPoslovi.Find "ID=" & Mid(Me.ctlListView.ListItems(i).Key, 3)
        bFinished = False
        Select Case de.rsPoslovi!IntervalID
          Case 2
            If Date > (de.rsPoslovi!Datum + de.rsPoslovi!Kolicina) Then bFinished = True
          Case 3
            If Date > (de.rsPoslovi!Datum + (de.rsPoslovi!Kolicina * 7)) Then bFinished = True
          Case 4
            If Date > DateSerial(Year(de.rsPoslovi!Datum), Month(de.rsPoslovi!Datum) + de.rsPoslovi!Kolicina, Day(de.rsPoslovi!Datum)) Then bFinished = True
          Case 5
            If Date > DateSerial(Year(de.rsPoslovi!Datum) + de.rsPoslovi!Kolicina, Month(de.rsPoslovi!Datum), Day(de.rsPoslovi!Datum)) Then bFinished = True
        End Select
        If bFinished Then
          Me.ctlListView.ListItems(i).SmallIcon = 2
        Else
          Me.ctlListView.ListItems(i).SmallIcon = 1
        End If
      Next
  End Select
End Sub
