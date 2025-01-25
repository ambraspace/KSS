VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPosao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Poslovi"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
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
   ScaleHeight     =   4095
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Podaci o ugovorenom poslu"
      Height          =   3855
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtKorisnikRB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Ime davaoca usluga"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtDavalacRB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Ime davaoca usluga"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtRB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Ime davaoca usluga"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtKolicina 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Text            =   "0"
         ToolTipText     =   "Dužina ugovorenog radnog odnosa izražena brojem navedenih intervala"
         Top             =   3360
         Width           =   975
      End
      Begin VB.ComboBox cboInterval 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Interval po kojem je davalac usluga plaæen"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtIznos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Text            =   "0"
         ToolTipText     =   "Novèani iznos koji plaæa korisnik usluga po intervalu"
         Top             =   2640
         Width           =   975
      End
      Begin MSComCtl2.DTPicker ctlDatum 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Datum poèetka radnog odnosa (današnji, ako je nepoznat)"
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24510465
         CurrentDate     =   37977
      End
      Begin VB.ComboBox cboKorisnik 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Korisnik usluga s kojim ugovarate posao"
         Top             =   1920
         Width           =   3975
      End
      Begin VB.ComboBox cboDavalac 
         Height          =   315
         ItemData        =   "frmPosao.frx":0000
         Left            =   840
         List            =   "frmPosao.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Davalac usluga kojem ugovarate posao"
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "Redni broj:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Kolièina:"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Interval:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Iznos (KM):"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Datum:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Korisnik usluga:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Davalac usluga:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdAddUpdate 
      Caption         =   "Dodaj"
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Poništi"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "frmPosao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bPressed As Boolean

Public Property Get OKPressed() As Boolean
  OKPressed = m_bPressed
End Property

Public Property Let OKPressed(bValue As Boolean)
  m_bPressed = bValue
End Property


Private Sub cboDavalac_Click()
  If Me.cboDavalac.ItemData(Me.cboDavalac.ListIndex) = Val(Me.txtDavalacRB) Then Exit Sub
  Me.txtDavalacRB.Text = Me.cboDavalac.ItemData(Me.cboDavalac.ListIndex)
End Sub

Private Sub cboInterval_Click()
  If Me.cboInterval.ListIndex = 0 Then
    Me.txtKolicina.Enabled = False
    Me.txtKolicina.BackColor = vbButtonFace
  Else
    Me.txtKolicina.Enabled = True
    Me.txtKolicina.BackColor = vbWindowBackground
  End If
End Sub

Private Function ProvjeriZanimanja() As Boolean
  ProvjeriZanimanja = False
  de.rsZanimanjaVeze.MoveFirst
  de.rsZanimanjaVeze.Find "RB=" & Me.cboDavalac.ItemData(Me.cboDavalac.ListIndex)
  Do While Not de.rsZanimanjaVeze.EOF
    de.rsKorZanimanjaVeze.MoveFirst
    de.rsKorZanimanjaVeze.Find "RB=" & Me.cboKorisnik.ItemData(Me.cboKorisnik.ListIndex)
    Do While Not de.rsKorZanimanjaVeze.EOF
      If de.rsZanimanjaVeze!ZanimanjeID = de.rsKorZanimanjaVeze!ZanimanjeID Then
        ProvjeriZanimanja = True
        Exit Function
      End If
      de.rsKorZanimanjaVeze.Find "RB=" & Me.cboKorisnik.ItemData(Me.cboKorisnik.ListIndex), 1
    Loop
    de.rsZanimanjaVeze.Find "RB=" & Me.cboDavalac.ItemData(Me.cboDavalac.ListIndex), 1
  Loop
End Function

Private Sub cboKorisnik_Click()
  If Me.cboKorisnik.ItemData(Me.cboKorisnik.ListIndex) = Val(Me.txtKorisnikRB) Then Exit Sub
  Me.txtKorisnikRB = Me.cboKorisnik.ItemData(Me.cboKorisnik.ListIndex)
End Sub

Private Sub cmdAddUpdate_Click()
  Dim a As Integer
  If Not CheckData() Then Exit Sub
  
  If Me.cmdAddUpdate.Caption = "Dodaj" Then
    If de.rsPoslovi.RecordCount > 0 Then
      de.rsPoslovi.MoveFirst
      de.rsPoslovi.Find "ID=" & Me.txtRB
      If Not de.rsPoslovi.EOF Then
        MsgBox "Ugovoreni posao sa rednim brojem " & Me.txtRB & " veæ postoji!" & vbCrLf & "Izaberite neki drugi redni broj.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
      End If
    End If
  Else
    de.rsPoslovi.MoveFirst
    de.rsPoslovi.Find "ID=" & Me.txtRB
    If Not de.rsPoslovi.EOF Then
      If Me.txtRB <> Mid(frmMain.ctlListView.SelectedItem.Key, 3) Then
        MsgBox "Ugovoreni posao sa rednim brojem " & Me.txtRB & " veæ postoji!" & vbCrLf & "Izaberite neki drugi redni broj.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
      End If
    End If
    de.rsPoslovi.MoveFirst
    de.rsPoslovi.Find "ID=" & Mid(frmMain.ctlListView.SelectedItem.Key, 3)
  End If
  
  If Not ProvjeriZanimanja() Then
    a = MsgBox("Korisnik ima potražnju za zanimanjem za koje davalac nije kvalifikovan!" & vbCrLf & _
      "Želite li ipak nastaviti sa ugovaranjem posla?", vbExclamation + vbYesNo, Me.Caption)
    If a = vbNo Then Exit Sub
  End If
  If Me.ctlDatum.Value <= Date Then
    a = MsgBox("Datum " & Me.ctlDatum.Value & " je prošao ili je današnji!" & vbCrLf & "Želite li ipak nastaviti sa ugovaranjem posla?", vbExclamation + vbYesNo, Me.Caption)
    If a = vbNo Then Exit Sub
  End If
  If Me.cmdAddUpdate.Caption = "Dodaj" Then
    de.rsPoslovi.AddNew
  End If
  de.rsPoslovi!ID = Me.txtRB
  de.rsPoslovi!DavalacRB = Me.cboDavalac.ItemData(Me.cboDavalac.ListIndex)
  de.rsPoslovi!KorisnikRB = Me.cboKorisnik.ItemData(Me.cboKorisnik.ListIndex)
  de.rsPoslovi!Datum = Format(Me.ctlDatum.Value, "yyyy-MM-dd")
  de.rsPoslovi!Iznos = Val(Me.txtIznos)
  de.rsPoslovi!IntervalID = Me.cboInterval.ItemData(Me.cboInterval.ListIndex)
  If Me.cboInterval.ListIndex = 0 Then
    de.rsPoslovi!Kolicina = 0
  Else
    de.rsPoslovi!Kolicina = Val(Me.txtKolicina)
  End If
  de.rsPoslovi.Update
  Me.OKPressed = True
  Me.txtRB.SetFocus
  Me.Hide
End Sub

Private Sub cmdCancel_Click()
  Me.txtRB.SetFocus
  Me.Hide
End Sub

Private Sub Form_Load()
  de.rsIntervali.MoveFirst
  Do Until de.rsIntervali.EOF
    Me.cboInterval.AddItem de.rsIntervali!Interval
    Me.cboInterval.ItemData(Me.cboInterval.ListCount - 1) = de.rsIntervali!ID
    de.rsIntervali.MoveNext
  Loop
End Sub

Public Sub AddNew()
  Me.OKPressed = False
  FillCombos
  Me.txtRB = 1
  If de.rsPoslovi.RecordCount > 0 Then
    de.rsPoslovi.MoveLast
    Me.txtRB = de.rsPoslovi!ID + 1
  End If
  Me.ctlDatum.Value = Date + 1
  Me.txtIznos = 0
  Me.cboInterval.ListIndex = 0
  Me.cboDavalac.ListIndex = 0
  Me.cboKorisnik.ListIndex = 0
  Me.txtKolicina = 0
  Me.Caption = "Poslovi - dodavanje"
  Me.cmdAddUpdate.Caption = "Dodaj"
  Me.Show 1, frmMain
End Sub

Public Sub Edit()
  Dim i As Integer
  Me.OKPressed = False
  FillCombos
  de.rsPoslovi.MoveFirst
  de.rsPoslovi.Find "ID=" & Mid(frmMain.ctlListView.SelectedItem.Key, 3)
  Me.txtRB = de.rsPoslovi!ID
  For i = 1 To Me.cboDavalac.ListCount
    If de.rsPoslovi!DavalacRB = Me.cboDavalac.ItemData(i - 1) Then
      Me.cboDavalac.ListIndex = i - 1
      Exit For
    End If
  Next
  For i = 1 To Me.cboKorisnik.ListCount
    If de.rsPoslovi!KorisnikRB = Me.cboKorisnik.ItemData(i - 1) Then
      Me.cboKorisnik.ListIndex = i - 1
      Exit For
    End If
  Next
  For i = 1 To Me.cboInterval.ListCount
    If de.rsPoslovi!IntervalID = Me.cboInterval.ItemData(i - 1) Then
      Me.cboInterval.ListIndex = i - 1
      Exit For
    End If
  Next
  Me.txtIznos = de.rsPoslovi!Iznos
  Me.txtKolicina = de.rsPoslovi!Kolicina
  Me.ctlDatum.Value = de.rsPoslovi!Datum
  Me.Caption = "Poslovi - izmjena"
  Me.cmdAddUpdate.Caption = "Izmijeni"
  Me.Show 1, frmMain
End Sub

Private Sub FillCombos()
  Dim rsTMP As ADODB.Recordset
  Me.cboDavalac.Clear
  Me.cboKorisnik.Clear
  Set rsTMP = de.cn1.Execute("SELECT * FROM Davaoci ORDER BY Prezime, Ime")
  rsTMP.MoveFirst
  Do Until rsTMP.EOF
    Me.cboDavalac.AddItem UCase(rsTMP!Prezime) & ", " & rsTMP!Ime
    Me.cboDavalac.ItemData(Me.cboDavalac.ListCount - 1) = rsTMP!RB
    rsTMP.MoveNext
  Loop
  Set rsTMP = Nothing
  Set rsTMP = de.cn1.Execute("SELECT * FROM Korisnici ORDER BY NazivPrezime, Ime")
  rsTMP.MoveFirst
  Do Until rsTMP.EOF
    If rsTMP!FL Then
      Me.cboKorisnik.AddItem UCase(rsTMP!NazivPrezime) & ", " & rsTMP!Ime
    Else
      Me.cboKorisnik.AddItem UCase(rsTMP!NazivPrezime) & ", " & rsTMP!Mjesto
    End If
    Me.cboKorisnik.ItemData(Me.cboKorisnik.ListCount - 1) = rsTMP!RB
    rsTMP.MoveNext
  Loop
  Set rsTMP = Nothing
End Sub

Private Function CheckData() As Boolean
  CheckData = False
  If Me.cboDavalac.Enabled = False Then
    MsgBox "Izaberite davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Me.cboKorisnik.Enabled = False Then
    MsgBox "Izaberite korisnika usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If CStr(Val(Me.txtRB)) <> Me.txtRB Then
    MsgBox "Unesite ispravan redni broj ugovorenog posla!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Me.cboInterval.ListIndex <> 0 And Val(Me.txtKolicina) = 0 Then
    MsgBox "Zadajte broj intervala!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  CheckData = True
End Function

Private Sub txtDavalacRB_Change()
  Dim i As Long
  If Me.cboDavalac.ItemData(Me.cboDavalac.ListIndex) = Val(Me.txtDavalacRB) Then
    If Me.cboDavalac.Enabled Then Exit Sub
  End If
  Me.cboDavalac.Enabled = False
  For i = 0 To Me.cboDavalac.ListCount - 1
    If Me.cboDavalac.ItemData(i) = Val(Me.txtDavalacRB) Then
      Me.cboDavalac.Enabled = True
      Me.cboDavalac.ListIndex = i
      Exit For
    End If
  Next
End Sub

Private Sub txtDavalacRB_GotFocus()
  SelAll
End Sub

Private Sub txtIznos_GotFocus()
  SelAll
End Sub

Private Sub SelAll()
  Me.ActiveControl.SelStart = 0
  Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtKolicina_GotFocus()
  SelAll
End Sub

Private Sub txtKorisnikRB_Change()
  Dim i As Long
  If Me.cboKorisnik.ItemData(Me.cboKorisnik.ListIndex) = Val(Me.txtKorisnikRB) Then
    If Me.cboKorisnik.Enabled Then Exit Sub
  End If
  Me.cboKorisnik.Enabled = False
  For i = 0 To Me.cboKorisnik.ListCount - 1
    If Me.cboKorisnik.ItemData(i) = Val(Me.txtKorisnikRB) Then
      Me.cboKorisnik.Enabled = True
      Me.cboKorisnik.ListIndex = i
      Exit For
    End If
  Next
End Sub

Private Sub txtKorisnikRB_GotFocus()
  SelAll
End Sub

Private Sub txtRB_GotFocus()
  SelAll
End Sub
