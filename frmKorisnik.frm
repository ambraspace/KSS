VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmKorisnik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Korisnik usluga"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Podaci o korisniku usluga"
      Height          =   7695
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5535
      Begin VB.ListBox lstDelZanimanja 
         Height          =   450
         Left            =   3840
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ListBox lstDelTelefoni 
         Height          =   450
         Left            =   2280
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtKomentar 
         Height          =   660
         Left            =   120
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         ToolTipText     =   "Komentar o korisniku usluga"
         Top             =   6840
         Width           =   5295
      End
      Begin VB.TextBox txtIme 
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Ime korisnika usluga"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtPrezime 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtAdresa 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Adresa korisnika usluga (ulica i broj)"
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox txtMjesto 
         Height          =   285
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   5
         ToolTipText     =   "Mjesto prebivališta korisnika usluga"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox cboTip 
         Height          =   315
         ItemData        =   "frmKorisnik.frx":0000
         Left            =   120
         List            =   "frmKorisnik.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Tip korisnika usluga (pravno ili fizièko lice)"
         Top             =   360
         Width           =   1695
      End
      Begin VB.ListBox lstDelZiroRacuni 
         Height          =   450
         Left            =   3840
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtKontaktOsoba 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Kontakt osoba - predstavnik korisnika usluga"
         Top             =   6240
         Width           =   3615
      End
      Begin VB.TextBox txtKontaktTel 
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   10
         ToolTipText     =   "Tel. broj kontakt osobe"
         Top             =   6240
         Width           =   1575
      End
      Begin VB.TextBox txtRB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Ime davaoca usluga"
         Top             =   1080
         Width           =   855
      End
      Begin MSComctlLib.ListView lstZanimanja 
         Height          =   660
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Zanimanja koja korisnik usluga potražuje (kliknite desnim dugmetom miša)"
         Top             =   5280
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1164
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Zanimanje"
            Object.Width           =   6165
         EndProperty
      End
      Begin MSComctlLib.ListView lstTelefoni 
         Height          =   900
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Telefonski brojevi korisnika usluga (kliknite desnim dugmetom miša)"
         Top             =   2880
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Telefon"
            Object.Width           =   3440
         EndProperty
      End
      Begin MSComctlLib.ListView lstZiroRacuni 
         Height          =   900
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Žiro-raèuni korisnika usluga (kliknite desnim dugmetom miša)"
         Top             =   4080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Broj"
            Object.Width           =   4710
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tip"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Banka"
            Object.Width           =   2672
         EndProperty
      End
      Begin VB.Label Label12 
         Caption         =   "Komentar:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   6600
         Width           =   3135
      End
      Begin VB.Label lblIme 
         Caption         =   "Ime:"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblNazivPrezime 
         Caption         =   "Prezime:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Adresa:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   "Mjesto:"
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Telefoni:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Tražena zanimanja:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   5040
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Žiro-raèuni:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3840
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Kontakt osoba:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   6000
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "Kontakt telefon:"
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Redni broj:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAddUpdate 
      Caption         =   "Dodaj"
      Height          =   495
      Left            =   5760
      TabIndex        =   12
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Poništi"
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Dodaj"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Izmijeni"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Obriši"
      End
   End
End
Attribute VB_Name = "frmKorisnik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurrControl As Control
Private m_bPressed As Boolean

Public Property Get OKPressed() As Boolean
  OKPressed = m_bPressed
End Property

Public Property Let OKPressed(bValue As Boolean)
  m_bPressed = bValue
End Property

Private Sub cboTip_Click()
  If Me.cboTip.ListIndex = 0 Then
    Me.lblNazivPrezime.Caption = "Naziv:"
    Me.lblIme.Visible = False
    Me.txtIme.Visible = False
    Me.txtPrezime.ToolTipText = "Naziv korisnika usluga - pravnog lica"
    Me.txtKontaktOsoba.Enabled = True
    Me.txtKontaktTel.Enabled = True
    Me.txtKontaktOsoba.BackColor = vbWindowBackground
    Me.txtKontaktTel.BackColor = vbWindowBackground
  Else
    Me.lblNazivPrezime.Caption = "Prezime:"
    Me.lblIme.Visible = True
    Me.txtIme.Visible = True
    Me.txtPrezime.ToolTipText = "Prezime korisnika usluga - fizièkog lica"
    Me.txtKontaktOsoba.Enabled = False
    Me.txtKontaktTel.Enabled = False
    Me.txtKontaktOsoba.BackColor = vbButtonFace
    Me.txtKontaktTel.BackColor = vbButtonFace
  End If
End Sub

Public Sub AddNew()
  Me.OKPressed = False
  Me.txtRB = 1
  If de.rsKorisnici.RecordCount > 0 Then
    de.rsKorisnici.MoveLast
    Me.txtRB = de.rsKorisnici!RB + 1
  End If
  Me.txtIme = ""
  Me.txtPrezime = ""
  Me.txtAdresa = ""
  Me.txtMjesto = ""
  Me.cboTip.ListIndex = 1
  Me.lstTelefoni.ListItems.Clear
  Me.lstDelTelefoni.Clear
  Me.lstZiroRacuni.ListItems.Clear
  Me.lstDelZiroRacuni.Clear
  Me.lstZanimanja.ListItems.Clear
  Me.lstDelZanimanja.Clear
  Me.txtKontaktOsoba = ""
  Me.txtKontaktTel = ""
  Me.txtKomentar = ""
  Me.Caption = "Korisnici usluga - dodavanje"
  Me.cmdAddUpdate.Caption = "Dodaj"
  Me.Show 1, frmMain
End Sub

Public Sub Edit()
  
  Me.OKPressed = False
  
  de.rsKorisnici.MoveFirst
  de.rsKorisnici.Find "RB=" & Mid(frmMain.ctlListView.SelectedItem.Key, 3)
  Me.txtRB = de.rsKorisnici!RB
  Me.txtPrezime = de.rsKorisnici!NazivPrezime
  If de.rsKorisnici!FL Then
    Me.cboTip.ListIndex = 1
    Me.txtIme = de.rsKorisnici!Ime
  Else
    Me.cboTip.ListIndex = 0
    Me.txtIme = ""
  End If
  Me.txtAdresa = de.rsKorisnici!Adresa
  Me.txtMjesto = de.rsKorisnici!Mjesto
  Me.lstTelefoni.ListItems.Clear
  If de.rsTelefoni.RecordCount > 0 Then
    de.rsTelefoni.MoveFirst
    de.rsTelefoni.Find "RB=" & de.rsKorisnici!RB
  End If
  Do While Not de.rsTelefoni.EOF
    If de.rsTelefoni!Korisnik Then Me.lstTelefoni.ListItems.Add , "ID" & de.rsTelefoni!ID, de.rsTelefoni!Telefon
    de.rsTelefoni.Find "RB=" & de.rsKorisnici!RB, 1
  Loop
  Me.lstZiroRacuni.ListItems.Clear
  If de.rsZiroRacuni.RecordCount > 0 Then
    de.rsZiroRacuni.MoveFirst
    de.rsZiroRacuni.Find "RB=" & de.rsKorisnici!RB
  End If
  Do While Not de.rsZiroRacuni.EOF
    Me.lstZiroRacuni.ListItems.Add , "ID" & de.rsZiroRacuni!ID, de.rsZiroRacuni!ZR
    Me.lstZiroRacuni.ListItems("ID" & de.rsZiroRacuni!ID).SubItems(1) = de.rsZiroRacuni!Tip
    Me.lstZiroRacuni.ListItems("ID" & de.rsZiroRacuni!ID).SubItems(2) = de.rsZiroRacuni!Banka
    de.rsZiroRacuni.Find "RB=" & de.rsKorisnici!RB, 1
  Loop
  Me.lstZanimanja.ListItems.Clear
  de.rsKorZanimanjaVeze.MoveFirst
  de.rsKorZanimanjaVeze.Find "RB=" & de.rsKorisnici!RB
  Do While Not de.rsKorZanimanjaVeze.EOF
    de.rsZanimanja.MoveFirst
    de.rsZanimanja.Find "ID=" & de.rsKorZanimanjaVeze!ZanimanjeID
    Me.lstZanimanja.ListItems.Add , "ID" & de.rsKorZanimanjaVeze!ID, de.rsZanimanja!Naziv
    de.rsKorZanimanjaVeze.Find "RB=" & de.rsKorisnici!RB, 1
  Loop
  Me.txtKontaktOsoba = ""
  Me.txtKontaktTel = ""
  Me.txtKomentar = ""
  If Not IsNull(de.rsKorisnici!Kontakt) Then Me.txtKontaktOsoba = de.rsKorisnici!Kontakt
  If Not IsNull(de.rsKorisnici!KontaktTel) Then Me.txtKontaktTel = de.rsKorisnici!KontaktTel
  If Not IsNull(de.rsKorisnici!Komentar) Then Me.txtKomentar = de.rsKorisnici!Komentar
  Me.Caption = "Korisnici usluga - izmjena"
  Me.cmdAddUpdate.Caption = "Izmijeni"
  
  Me.lstDelTelefoni.Clear
  Me.lstDelZanimanja.Clear
  Me.lstDelZiroRacuni.Clear
  
  Me.Show 1, frmMain
End Sub


Private Sub cmdAddUpdate_Click()
  Dim i As Integer
  
  If Not CheckData Then Exit Sub

  If Me.cmdAddUpdate.Caption = "Dodaj" Then
    If de.rsKorisnici.RecordCount > 0 Then
      de.rsKorisnici.MoveFirst
      de.rsKorisnici.Find "RB=" & Me.txtRB
      If Not de.rsKorisnici.EOF Then
        MsgBox "Korisnik usluga sa rednim brojem " & Me.txtRB & " veæ postoji!" & vbCrLf & "Izaberite neki drugi redni broj.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
      End If
    End If
    de.rsKorisnici.AddNew
    de.rsKorisnici!RB = Me.txtRB
    de.rsKorisnici!NazivPrezime = Me.txtPrezime
    de.rsKorisnici!Adresa = Me.txtAdresa
    de.rsKorisnici!Mjesto = Me.txtMjesto
    If Me.cboTip.ListIndex = 1 Then
      de.rsKorisnici!FL = True
      de.rsKorisnici!Ime = Me.txtIme
      de.rsKorisnici!Kontakt = ""
      de.rsKorisnici!KontaktTel = ""
    Else
      Me.txtIme = ""
      de.rsKorisnici!Kontakt = Me.txtKontaktOsoba
      de.rsKorisnici!KontaktTel = Me.txtKontaktTel
    End If
    de.rsKorisnici!Komentar = Me.txtKomentar
    de.rsKorisnici.Update
    For i = 1 To Me.lstTelefoni.ListItems.Count
      de.rsTelefoni.AddNew
      de.rsTelefoni!Korisnik = True
      de.rsTelefoni!RB = de.rsKorisnici!RB
      de.rsTelefoni!Telefon = Me.lstTelefoni.ListItems(i).Text
      de.rsTelefoni.Update
    Next
    If Me.lstZiroRacuni.ListItems.Count > 0 Then
      For i = 1 To Me.lstZiroRacuni.ListItems.Count
        de.rsZiroRacuni.AddNew
        de.rsZiroRacuni!RB = de.rsKorisnici!RB
        de.rsZiroRacuni!ZR = Me.lstZiroRacuni.ListItems(i).Text
        de.rsZiroRacuni!Tip = Me.lstZiroRacuni.ListItems(i).SubItems(1)
        de.rsZiroRacuni!Banka = Me.lstZiroRacuni.ListItems(i).SubItems(2)
        de.rsZiroRacuni.Update
      Next
    End If
    For i = 1 To Me.lstZanimanja.ListItems.Count
      If de.rsZanimanja.RecordCount > 0 Then
        de.rsZanimanja.MoveFirst
        de.rsZanimanja.Find "Naziv='" & Me.lstZanimanja.ListItems(i).Text & "'"
      End If
      If de.rsZanimanja.EOF Then
        de.rsZanimanja.AddNew
        de.rsZanimanja!Naziv = Me.lstZanimanja.ListItems(i).Text
        de.rsZanimanja.Update
      End If
      de.rsKorZanimanjaVeze.AddNew
      de.rsKorZanimanjaVeze!RB = de.rsKorisnici!RB
      de.rsKorZanimanjaVeze!ZanimanjeID = de.rsZanimanja!ID
      de.rsKorZanimanjaVeze.Update
    Next
  Else 'IZMIJENI
    de.rsKorisnici.MoveFirst
    de.rsKorisnici.Find "RB=" & Me.txtRB
    If Not de.rsKorisnici.EOF Then
      If Me.txtRB <> Mid(frmMain.ctlListView.SelectedItem.Key, 3) Then
        MsgBox "Korisnik usluga sa rednim brojem " & Me.txtRB & " veæ postoji!" & vbCrLf & "Izaberite neki drugi redni broj.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
      End If
    End If
    de.rsKorisnici.MoveFirst
    de.rsKorisnici.Find "RB=" & Mid(frmMain.ctlListView.SelectedItem.Key, 3)
    If de.rsKorisnici!RB <> Me.txtRB Then
      de.cn1.Execute "UPDATE Poslovi SET KorisnikRB=" & Me.txtRB & " WHERE KorisnikRB=" & de.rsKorisnici!RB
      de.cn1.Execute "UPDATE PosloviBKP SET KorisnikRB=" & Me.txtRB & " WHERE KorisnikRB=" & de.rsKorisnici!RB
      de.cn1.Execute "UPDATE KorZanimanjaVeze SET RB=" & Me.txtRB & " WHERE RB=" & de.rsKorisnici!RB
      de.cn1.Execute "UPDATE Telefoni SET RB=" & Me.txtRB & " WHERE Korisnik=TRUE AND RB=" & de.rsKorisnici!RB
      de.cn1.Execute "UPDATE ZiroRacuni SET RB=" & Me.txtRB & " WHERE RB=" & de.rsKorisnici!RB
      de.rsPoslovi.Requery
      de.rsPosloviBKP.Requery
      de.rsKorZanimanjaVeze.Requery
      de.rsTelefoni.Requery
      de.rsZiroRacuni.Requery
      de.rsKorisnici!RB = Me.txtRB
    End If
    de.rsKorisnici!FL = False
    If Me.cboTip.ListIndex = 1 Then de.rsKorisnici!FL = True
    de.rsKorisnici!NazivPrezime = Me.txtPrezime
    If Me.cboTip.ListIndex = 1 Then
      de.rsKorisnici!Ime = Me.txtIme
      de.rsKorisnici!Kontakt = ""
      de.rsKorisnici!KontaktTel = ""
    Else
      de.rsKorisnici!Ime = ""
      de.rsKorisnici!Kontakt = Me.txtKontaktOsoba
      de.rsKorisnici!KontaktTel = Me.txtKontaktTel
    End If
    de.rsKorisnici!Adresa = Me.txtAdresa
    de.rsKorisnici!Mjesto = Me.txtMjesto
    de.rsKorisnici!Komentar = Me.txtKomentar
    de.rsKorisnici.Update
    For i = 1 To Me.lstTelefoni.ListItems.Count
      If Me.lstTelefoni.ListItems(i).Key = "" Then
        de.rsTelefoni.AddNew
        de.rsTelefoni!Korisnik = True
      Else
        de.rsTelefoni.MoveFirst
        de.rsTelefoni.Find "ID=" & Mid(Me.lstTelefoni.ListItems(i).Key, 3)
      End If
      de.rsTelefoni!RB = de.rsKorisnici!RB
      de.rsTelefoni!Telefon = Me.lstTelefoni.ListItems(i).Text
      de.rsTelefoni.Update
    Next
    If Me.lstDelTelefoni.ListCount > 0 Then
      For i = 1 To Me.lstDelTelefoni.ListCount
        de.cn1.Execute "DELETE FROM Telefoni WHERE ID=" & Me.lstDelTelefoni.List(i - 1)
      Next
      de.rsTelefoni.Requery
    End If
    If Me.lstZiroRacuni.ListItems.Count > 0 Then
      For i = 1 To Me.lstZiroRacuni.ListItems.Count
        If Me.lstZiroRacuni.ListItems(i).Key = "" Then
          de.rsZiroRacuni.AddNew
        Else
          de.rsZiroRacuni.MoveFirst
          de.rsZiroRacuni.Find "ID=" & Mid(Me.lstZiroRacuni.ListItems(i).Key, 3)
        End If
        de.rsZiroRacuni!RB = de.rsKorisnici!RB
        de.rsZiroRacuni!ZR = Me.lstZiroRacuni.ListItems(i).Text
        de.rsZiroRacuni!Tip = Me.lstZiroRacuni.ListItems(i).SubItems(1)
        de.rsZiroRacuni!Banka = Me.lstZiroRacuni.ListItems(i).SubItems(2)
        de.rsZiroRacuni.Update
      Next
    End If
    If Me.lstDelZiroRacuni.ListCount > 0 Then
      For i = 1 To Me.lstDelZiroRacuni.ListCount
        de.cn1.Execute "DELETE FROM ZiroRacuni WHERE ID=" & Me.lstDelZiroRacuni.List(i - 1)
      Next
      de.rsZiroRacuni.Requery
    End If
    For i = 1 To Me.lstZanimanja.ListItems.Count
      If de.rsZanimanja.RecordCount > 0 Then
        de.rsZanimanja.MoveFirst
        de.rsZanimanja.Find "Naziv='" & Me.lstZanimanja.ListItems(i).Text & "'"
      End If
      If de.rsZanimanja.EOF Then
        de.rsZanimanja.AddNew
        de.rsZanimanja!Naziv = Me.lstZanimanja.ListItems(i).Text
        de.rsZanimanja.Update
      End If
      If Me.lstZanimanja.ListItems(i).Key = "" Then
        de.rsKorZanimanjaVeze.AddNew
      Else
        de.rsKorZanimanjaVeze.MoveFirst
        de.rsKorZanimanjaVeze.Find "ID=" & Mid(Me.lstZanimanja.ListItems(i).Key, 3)
      End If
      de.rsKorZanimanjaVeze!RB = de.rsKorisnici!RB
      de.rsKorZanimanjaVeze!ZanimanjeID = de.rsZanimanja!ID
      de.rsKorZanimanjaVeze.Update
    Next
  End If
  Me.txtRB.SetFocus
  Me.OKPressed = True
  Me.Hide
End Sub

Private Sub Form_Load()
  Me.cboTip.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
  Me.txtRB.SetFocus
  Me.Hide
End Sub


Private Sub lstTelefoni_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim c As MSComctlLib.ListItem, sel As Boolean
  sel = False
  For Each c In Me.lstTelefoni.ListItems
    If c.Selected Then
      sel = True
      Exit For
    End If
  Next
  If Button = vbRightButton Then
    If sel Then
      Me.mnuDel.Enabled = True
      Me.mnuEdit.Enabled = True
    Else
      Me.mnuDel.Enabled = False
      Me.mnuEdit.Enabled = False
    End If
    Set CurrControl = Me.lstTelefoni
    Me.PopupMenu mnuPopUp, 2, x + Me.lstTelefoni.Left + Me.Frame1.Left, y + Me.lstTelefoni.Top + Me.Frame1.Top
  End If
End Sub

Private Sub lstZanimanja_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim c As MSComctlLib.ListItem, sel As Boolean
  sel = False
  For Each c In Me.lstZanimanja.ListItems
    If c.Selected Then
      sel = True
      Exit For
    End If
  Next
  If Button = vbRightButton Then
    If sel Then
      Me.mnuDel.Enabled = True
      Me.mnuEdit.Enabled = True
    Else
      Me.mnuDel.Enabled = False
      Me.mnuEdit.Enabled = False
    End If
    Set CurrControl = Me.lstZanimanja
    Me.PopupMenu mnuPopUp, 2, x + Me.lstZanimanja.Left + Me.Frame1.Left, y + Me.lstZanimanja.Top + Me.Frame1.Top
  End If
End Sub

Private Sub lstZiroRacuni_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim c As MSComctlLib.ListItem, sel As Boolean
  sel = False
  For Each c In Me.lstZiroRacuni.ListItems
    If c.Selected Then
      sel = True
      Exit For
    End If
  Next
  If Button = vbRightButton Then
    If sel Then
      Me.mnuDel.Enabled = True
      Me.mnuEdit.Enabled = True
    Else
      Me.mnuDel.Enabled = False
      Me.mnuEdit.Enabled = False
    End If
    Set CurrControl = Me.lstZiroRacuni
    Me.PopupMenu mnuPopUp, 2, x + Me.lstZiroRacuni.Left + Me.Frame1.Left, y + Me.lstZiroRacuni.Top + Me.Frame1.Top
  End If
End Sub



Private Sub mnuAdd_Click()
  Select Case CurrControl.Name
    Case "lstTelefoni"
      frmList.Caption = "Telefoni"
      frmList.lblNaziv.Caption = "Telefon:"
      frmList.cboVrijednost.Visible = False
      frmList.txtVrijednost.Visible = True
      frmList.txtVrijednost.ToolTipText = "Telefonski broj korisnika usluga"
      frmList.cmdAddEdit.Caption = "Dodaj"
      frmList.txtVrijednost.MaxLength = 15
      frmList.txtVrijednost = ""
      frmList.OKPressed = False
      frmList.Show 1, Me
      If frmList.OKPressed Then
        Me.lstTelefoni.ListItems.Add , , frmList.txtVrijednost
      End If
    Case "lstZiroRacuni"
      frmZR.txtBroj = ""
      frmZR.cboTip.ListIndex = 0
      frmZR.txtBanka = ""
      frmZR.cmdAddEdit.Caption = "Dodaj"
      frmZR.OKPressed = False
      frmZR.Show 1, Me
      If frmZR.OKPressed Then
        Me.lstZiroRacuni.ListItems.Add , , frmZR.txtBroj
        Me.lstZiroRacuni.ListItems(Me.lstZiroRacuni.ListItems.Count).SubItems(1) = frmZR.cboTip
        Me.lstZiroRacuni.ListItems(Me.lstZiroRacuni.ListItems.Count).SubItems(2) = frmZR.txtBanka
      End If
    Case "lstZanimanja"
      frmList.Caption = "Zanimanja"
      frmList.lblNaziv.Caption = "Zanimanje:"
      frmList.cboVrijednost.Visible = True
      frmList.cboVrijednost.ToolTipText = "Zanimanje koje korisnik usluga potražuje"
      frmList.txtVrijednost.Visible = False
      frmList.cmdAddEdit.Caption = "Dodaj"
      frmList.cboVrijednost.Clear
      If de.rsZanimanja.RecordCount > 0 Then
        de.rsZanimanja.MoveFirst
        Do Until de.rsZanimanja.EOF
          frmList.cboVrijednost.AddItem de.rsZanimanja!Naziv
          de.rsZanimanja.MoveNext
        Loop
      End If
      frmList.OKPressed = False
      frmList.Show 1, Me
      If frmList.OKPressed Then
        Me.lstZanimanja.ListItems.Add , , frmList.cboVrijednost.Text
      End If
  End Select
End Sub

Private Sub mnuDel_Click()
  Dim a As Integer
  Select Case CurrControl.Name
    Case "lstTelefoni"
      a = MsgBox("Da li želite obrisati izabrani tel. broj?", vbQuestion + vbYesNo, Me.Caption)
      If a = vbYes Then
        If Left(Me.lstTelefoni.SelectedItem.Key, 2) = "ID" Then Me.lstDelTelefoni.AddItem Mid(Me.lstTelefoni.SelectedItem.Key, 3)
        Me.lstTelefoni.ListItems.Remove Me.lstTelefoni.SelectedItem.Index
      End If
    Case "lstZiroRacuni"
      a = MsgBox("Da li želite obrisati izabrani žiro-raèun?", vbQuestion + vbYesNo, Me.Caption)
      If a = vbYes Then
        If Left(Me.lstZiroRacuni.SelectedItem.Key, 2) = "ID" Then Me.lstDelZiroRacuni.AddItem Mid(Me.lstZiroRacuni.SelectedItem.Key, 3)
        Me.lstZiroRacuni.ListItems.Remove Me.lstZiroRacuni.SelectedItem.Index
      End If
    Case "lstZanimanja"
      a = MsgBox("Da li želite obrisati izabrano zanimanje?", vbQuestion + vbYesNo, Me.Caption)
      If a = vbYes Then
        If Left(Me.lstZanimanja.SelectedItem.Key, 2) = "ID" Then Me.lstDelZanimanja.AddItem Mid(Me.lstZanimanja.SelectedItem.Key, 3)
        Me.lstZanimanja.ListItems.Remove Me.lstZanimanja.SelectedItem.Index
      End If
  End Select
End Sub

Private Sub mnuEdit_Click()
  Select Case CurrControl.Name
    Case "lstTelefoni"
      frmList.Caption = "Telefoni"
      frmList.lblNaziv.Caption = "Telefon:"
      frmList.cboVrijednost.Visible = False
      frmList.txtVrijednost.Visible = True
      frmList.txtVrijednost.ToolTipText = "Telefonski broj korisnika usluga"
      frmList.cmdAddEdit.Caption = "Izmijeni"
      frmList.txtVrijednost.MaxLength = 15
      frmList.txtVrijednost = Me.lstTelefoni.SelectedItem.Text
      frmList.OKPressed = False
      frmList.Show 1, Me
      If frmList.OKPressed Then
        Me.lstTelefoni.SelectedItem.Text = frmList.txtVrijednost
      End If
    Case "lstZiroRacuni"
      frmZR.txtBroj = Me.lstZiroRacuni.SelectedItem.Text
      If Me.lstZiroRacuni.SelectedItem.SubItems(1) = "KM" Then
        frmZR.cboTip.ListIndex = 0
      Else
        frmZR.cboTip.ListIndex = 1
      End If
      frmZR.txtBanka = Me.lstZiroRacuni.SelectedItem.SubItems(2)
      frmZR.cmdAddEdit.Caption = "Izmijeni"
      frmZR.OKPressed = False
      frmZR.Show 1, Me
      If frmZR.OKPressed Then
        Me.lstZiroRacuni.SelectedItem.Text = frmZR.txtBroj
        Me.lstZiroRacuni.SelectedItem.SubItems(1) = frmZR.cboTip
        Me.lstZiroRacuni.SelectedItem.SubItems(2) = frmZR.txtBanka
      End If
    Case "lstZanimanja"
      frmList.Caption = "Zanimanja"
      frmList.lblNaziv.Caption = "Zanimanje:"
      frmList.cboVrijednost.Visible = True
      frmList.cboVrijednost.ToolTipText = "Zanimanje koje korisnik usluga potražuje"
      frmList.txtVrijednost.Visible = False
      frmList.cmdAddEdit.Caption = "Izmijeni"
      frmList.cboVrijednost.Clear
      If de.rsZanimanja.RecordCount > 0 Then
        de.rsZanimanja.MoveFirst
        Do Until de.rsZanimanja.EOF
          frmList.cboVrijednost.AddItem de.rsZanimanja!Naziv
          de.rsZanimanja.MoveNext
        Loop
      End If
      frmList.cboVrijednost = Me.lstZanimanja.SelectedItem.Text
      frmList.OKPressed = False
      frmList.Show 1, Me
      If frmList.OKPressed Then
        Me.lstZanimanja.SelectedItem.Text = frmList.cboVrijednost.Text
      End If
  End Select
End Sub

Private Function CheckData() As Boolean
  CheckData = False
  If CStr(Val(Me.txtRB)) <> Me.txtRB Then
    MsgBox "Unesite ispravan redni broj korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtPrezime) = "" Then
    If Me.cboTip.ListIndex = 0 Then
      MsgBox "Unesite naziv korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
    Else
      MsgBox "Unesite prezime korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
    End If
    Exit Function
  End If
  If Me.cboTip.ListIndex = 1 And Trim(Me.txtIme) = "" Then
    MsgBox "Unesite ime korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtAdresa) = "" Then
    MsgBox "Unesite adresu korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtMjesto) = "" Then
    MsgBox "Unesite mjesto prebivališta korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Me.lstTelefoni.ListItems.Count < 1 Then
    MsgBox "Unesite bar jedan tel. broj korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Me.lstZanimanja.ListItems.Count < 1 Then
    MsgBox "Unesite bar jedno traženo zanimanje za korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  CheckData = True
End Function

Private Sub txtAdresa_GotFocus()
  SelAll
End Sub

Private Sub txtIme_GotFocus()
  SelAll
End Sub

Private Sub txtKomentar_GotFocus()
  SelAll
End Sub

Private Sub txtKontaktOsoba_GotFocus()
  SelAll
End Sub

Private Sub txtKontaktTel_GotFocus()
  SelAll
End Sub

Private Sub txtMjesto_GotFocus()
  SelAll
End Sub

Private Sub txtPrezime_GotFocus()
  SelAll
End Sub

Private Sub SelAll()
  Me.ActiveControl.SelStart = 0
  Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtRB_GotFocus()
  SelAll
End Sub
