VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDavalac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Davalac usluga"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Poništi"
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddUpdate 
      Caption         =   "Dodaj"
      Height          =   495
      Left            =   5760
      TabIndex        =   16
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Podaci o davaocu usluga"
      Height          =   6735
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   5535
      Begin MSComCtl2.DTPicker ctlDatumPP 
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   24510465
         CurrentDate     =   38035
      End
      Begin VB.TextBox txtRB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Ime davaoca usluga"
         Top             =   480
         Width           =   855
      End
      Begin VB.ListBox lstDelRaspolozivost 
         Height          =   450
         Left            =   3960
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ListBox lstDelZanimanja 
         Height          =   450
         Left            =   3960
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ListBox lstDelTelefoni 
         Height          =   450
         Left            =   2280
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComctlLib.ListView lstRaspolozivost 
         Height          =   780
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Raspoloživost davaoca usluga (kliknite desnim dugmetom miša)"
         Top             =   5760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1376
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
            Text            =   "Raspoloživost"
            Object.Width           =   3387
         EndProperty
      End
      Begin MSComctlLib.ListView lstZanimanja 
         Height          =   660
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Zanimanja za koja je davalac usluga kvalifikovan ili može da radi (kliknite desnim dugmetom miša)"
         Top             =   4800
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VB.TextBox txtKomentar 
         Height          =   780
         Left            =   2280
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         ToolTipText     =   "Komentar o davaocu usluga"
         Top             =   5760
         Width           =   3135
      End
      Begin VB.TextBox txtIme 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Ime davaoca usluga"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtPrezime 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Prezime davaoca usluga"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtRoditelj 
         Height          =   285
         Left            =   3960
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Ime jednog roditelja davaoca usluga"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Frame frmPol 
         Caption         =   "Pol"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Pol davaoca usluga"
         Top             =   1440
         Width           =   2055
         Begin VB.OptionButton optZ 
            Caption         =   "Ženski"
            Height          =   255
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optM 
            Caption         =   "Muški"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.TextBox txtAdresa 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Adresa davaoca usluga (ulica i broj)"
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txtMjesto 
         Height          =   285
         Left            =   3960
         MaxLength       =   30
         TabIndex        =   8
         ToolTipText     =   "Mjesto prebivališta davaoca usluga"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtJMB 
         Height          =   285
         Left            =   120
         MaxLength       =   13
         TabIndex        =   11
         ToolTipText     =   "Jedinstveni matièni broj davaoca usluga"
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox txtBRLK 
         Height          =   285
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   12
         ToolTipText     =   "Broj liène karte davaoca usluga"
         Top             =   4200
         Width           =   1575
      End
      Begin MSComctlLib.ListView lstTelefoni 
         Height          =   900
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Telefoni davaoca usluga (kliknite desnim dugmetom miša)"
         Top             =   3000
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
      Begin VB.Label Label10 
         Caption         =   "Dat. posljednjeg posla:"
         Height          =   255
         Left            =   2280
         TabIndex        =   34
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Redni broj:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Raspoloživost:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "Komentar:"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label lblIme 
         Caption         =   "Ime:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Prezime:"
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Ime roditelja:"
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Adresa:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   "Mjesto:"
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Telefoni:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "JMB:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Broj liène karte:"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Zanimanja / poslovi:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Width           =   3615
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
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
Attribute VB_Name = "frmDavalac"
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

Public Sub AddNew()
  Me.OKPressed = False
  Me.txtRB = 1
  If de.rsDavaoci.RecordCount > 0 Then
    de.rsDavaoci.MoveLast
    Me.txtRB = de.rsDavaoci!RB + 1
  End If
  Me.txtIme = ""
  Me.txtPrezime = ""
  Me.txtRoditelj = ""
  Me.optM = True
  Me.txtAdresa = ""
  Me.txtMjesto = ""
  Me.lstTelefoni.ListItems.Clear
  Me.lstDelTelefoni.Clear
  Me.ctlDatumPP.Value = ""
  Me.txtJMB = ""
  Me.txtBRLK = ""
  Me.lstZanimanja.ListItems.Clear
  Me.lstDelZanimanja.Clear
  Me.lstRaspolozivost.ListItems.Clear
  Me.lstDelRaspolozivost.Clear
  Me.txtKomentar = ""
  Me.Caption = "Davaoci usluga - dodavanje"
  Me.cmdAddUpdate.Caption = "Dodaj"
  Me.Show 1, frmMain
End Sub

Public Sub Edit()
  Me.OKPressed = False
  de.rsDavaoci.MoveFirst
  de.rsDavaoci.Find "RB=" & Mid(frmMain.ctlListView.SelectedItem.Key, 3)
  Me.txtRB = de.rsDavaoci!RB
  Me.txtIme = de.rsDavaoci!Ime
  Me.txtPrezime = de.rsDavaoci!Prezime
  Me.txtRoditelj = de.rsDavaoci!Roditelj
  Me.optM = True
  If de.rsDavaoci!Pol = "Ž" Then Me.optZ = True
  Me.txtAdresa = de.rsDavaoci!Adresa
  Me.txtMjesto = de.rsDavaoci!Mjesto
  Me.lstTelefoni.ListItems.Clear
  If de.rsTelefoni.RecordCount > 0 Then
    de.rsTelefoni.MoveFirst
    de.rsTelefoni.Find "RB=" & de.rsDavaoci!RB
  End If
  Do While Not de.rsTelefoni.EOF
    If Not de.rsTelefoni!Korisnik Then Me.lstTelefoni.ListItems.Add , "ID" & de.rsTelefoni!ID, de.rsTelefoni!Telefon
    de.rsTelefoni.Find "RB=" & de.rsDavaoci!RB, 1
  Loop
  If IsNull(de.rsDavaoci!DatPosljednjegPosla) Then
    Me.ctlDatumPP.Value = ""
  Else
    Me.ctlDatumPP.Value = de.rsDavaoci!DatPosljednjegPosla
  End If
  Me.txtJMB = de.rsDavaoci!JMB
  Me.txtBRLK = de.rsDavaoci!BRLK
  Me.lstZanimanja.ListItems.Clear
  de.rsZanimanjaVeze.MoveFirst
  de.rsZanimanjaVeze.Find "RB=" & de.rsDavaoci!RB
  Do While Not de.rsZanimanjaVeze.EOF
    de.rsZanimanja.MoveFirst
    de.rsZanimanja.Find "ID=" & de.rsZanimanjaVeze!ZanimanjeID
    Me.lstZanimanja.ListItems.Add , "ID" & de.rsZanimanjaVeze!ID, de.rsZanimanja!Naziv
    de.rsZanimanjaVeze.Find "RB=" & de.rsDavaoci!RB, 1
  Loop
  Me.lstRaspolozivost.ListItems.Clear
  If de.rsRaspolozivostVeze.RecordCount > 0 Then
    de.rsRaspolozivostVeze.MoveFirst
    de.rsRaspolozivostVeze.Find "RB=" & de.rsDavaoci!RB
  End If
  Do While Not de.rsRaspolozivostVeze.EOF
    de.rsRaspolozivost.MoveFirst
    de.rsRaspolozivost.Find "ID=" & de.rsRaspolozivostVeze!RaspolozivostID
    Me.lstRaspolozivost.ListItems.Add , "ID" & de.rsRaspolozivostVeze!ID, de.rsRaspolozivost!Naziv
    de.rsRaspolozivostVeze.Find "RB=" & de.rsDavaoci!RB, 1
  Loop
  Me.txtKomentar = de.rsDavaoci!Komentar
  Me.Caption = "Davaoci usluga - izmjena"
  Me.cmdAddUpdate.Caption = "Izmijeni"
  
  Me.lstDelRaspolozivost.Clear
  Me.lstDelTelefoni.Clear
  Me.lstDelZanimanja.Clear
  
  Me.Show 1, frmMain
End Sub

Private Sub cmdAddUpdate_Click()
  Dim i As Integer
  If Not CheckData() Then Exit Sub
    
  If Me.cmdAddUpdate.Caption = "Dodaj" Then
    If de.rsDavaoci.RecordCount > 0 Then
      de.rsDavaoci.MoveFirst
      de.rsDavaoci.Find "RB=" & Me.txtRB
      If Not de.rsDavaoci.EOF Then
        MsgBox "Davalac usluga sa rednim brojem " & Me.txtRB & " veæ postoji!" & vbCrLf & "Izaberite neki drugi redni broj.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
      End If
    End If
    With de.rsDavaoci
      .AddNew
      !RB = Me.txtRB
      !Ime = Me.txtIme
      !Prezime = Me.txtPrezime
      !Roditelj = Me.txtRoditelj
      !Pol = "M"
      If Me.optZ Then !Pol = "Ž"
      !Adresa = Me.txtAdresa
      !Mjesto = Me.txtMjesto
      !DatPosljednjegPosla = Me.ctlDatumPP.Value
      !JMB = Me.txtJMB
      !BRLK = Me.txtBRLK
      !Komentar = Me.txtKomentar
      .Update
    End With
    If Me.lstTelefoni.ListItems.Count > 0 Then
      For i = 1 To Me.lstTelefoni.ListItems.Count
        de.rsTelefoni.AddNew
        de.rsTelefoni!RB = de.rsDavaoci!RB
        de.rsTelefoni!Telefon = Me.lstTelefoni.ListItems(i).Text
        de.rsTelefoni.Update
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
      de.rsZanimanjaVeze.AddNew
      de.rsZanimanjaVeze!RB = de.rsDavaoci!RB
      de.rsZanimanjaVeze!ZanimanjeID = de.rsZanimanja!ID
      de.rsZanimanjaVeze.Update
    Next
    If Me.lstRaspolozivost.ListItems.Count > 0 Then
      For i = 1 To Me.lstRaspolozivost.ListItems.Count
        If de.rsRaspolozivost.RecordCount > 0 Then
          de.rsRaspolozivost.MoveFirst
          de.rsRaspolozivost.Find "Naziv='" & Me.lstRaspolozivost.ListItems(i).Text & "'"
        End If
        If de.rsRaspolozivost.EOF Then
          de.rsRaspolozivost.AddNew
          de.rsRaspolozivost!Naziv = Me.lstRaspolozivost.ListItems(i).Text
          de.rsRaspolozivost.Update
        End If
        de.rsRaspolozivostVeze.AddNew
        de.rsRaspolozivostVeze!RB = de.rsDavaoci!RB
        de.rsRaspolozivostVeze!RaspolozivostID = de.rsRaspolozivost!ID
        de.rsRaspolozivostVeze.Update
      Next
    End If
    Me.txtIme.SetFocus
    Me.OKPressed = True
    Me.Hide
  Else ' IZMIJENI
    de.rsDavaoci.MoveFirst
    de.rsDavaoci.Find "RB=" & Me.txtRB
    If Not de.rsDavaoci.EOF Then
      If Me.txtRB <> Mid(frmMain.ctlListView.SelectedItem.Key, 3) Then
        MsgBox "Davalac usluga sa rednim brojem " & Me.txtRB & " veæ postoji!" & vbCrLf & "Izaberite neki drugi redni broj.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
      End If
    End If
    de.rsDavaoci.MoveFirst
    de.rsDavaoci.Find "RB=" & Mid(frmMain.ctlListView.SelectedItem.Key, 3)
    If de.rsDavaoci!RB <> Me.txtRB Then
      de.cn1.Execute "UPDATE Poslovi SET DavalacRB=" & Me.txtRB & " WHERE DavalacRB=" & de.rsDavaoci!RB
      de.cn1.Execute "UPDATE PosloviBKP SET DavalacRB=" & Me.txtRB & " WHERE DavalacRB=" & de.rsDavaoci!RB
      de.cn1.Execute "UPDATE RaspolozivostVeze SET RB=" & Me.txtRB & " WHERE RB=" & de.rsDavaoci!RB
      de.cn1.Execute "UPDATE Telefoni SET RB=" & Me.txtRB & " WHERE Korisnik=FALSE AND RB=" & de.rsDavaoci!RB
      de.cn1.Execute "UPDATE ZanimanjaVeze SET RB=" & Me.txtRB & " WHERE RB=" & de.rsDavaoci!RB
      de.rsPoslovi.Requery
      de.rsPosloviBKP.Requery
      de.rsRaspolozivostVeze.Requery
      de.rsTelefoni.Requery
      de.rsZanimanjaVeze.Requery
      de.rsDavaoci!RB = Me.txtRB
    End If
    With de.rsDavaoci
      !Ime = Me.txtIme
      !Prezime = Me.txtPrezime
      !Roditelj = Me.txtRoditelj
      !Pol = "M"
      If Me.optZ Then !Pol = "Ž"
      !Adresa = Me.txtAdresa
      !Mjesto = Me.txtMjesto
      !DatPosljednjegPosla = Me.ctlDatumPP.Value
      !JMB = Me.txtJMB
      !BRLK = Me.txtBRLK
      !Komentar = Me.txtKomentar
      .Update
    End With
    If Me.lstTelefoni.ListItems.Count > 0 Then
      For i = 1 To Me.lstTelefoni.ListItems.Count
        If Me.lstTelefoni.ListItems(i).Key = "" Then
          de.rsTelefoni.AddNew
        Else
          de.rsTelefoni.MoveFirst
          de.rsTelefoni.Find "ID=" & Mid(Me.lstTelefoni.ListItems(i).Key, 3)
        End If
        de.rsTelefoni!RB = de.rsDavaoci!RB
        de.rsTelefoni!Telefon = Me.lstTelefoni.ListItems(i).Text
        de.rsTelefoni.Update
      Next
    End If
    If Me.lstDelTelefoni.ListCount > 0 Then
      For i = 1 To Me.lstDelTelefoni.ListCount
        de.cn1.Execute "DELETE FROM Telefoni WHERE ID=" & Me.lstDelTelefoni.List(i - 1)
      Next
      de.rsTelefoni.Requery
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
        de.rsZanimanjaVeze.AddNew
      Else
        de.rsZanimanjaVeze.MoveFirst
        de.rsZanimanjaVeze.Find "ID=" & Mid(Me.lstZanimanja.ListItems(i).Key, 3)
      End If
      de.rsZanimanjaVeze!RB = de.rsDavaoci!RB
      de.rsZanimanjaVeze!ZanimanjeID = de.rsZanimanja!ID
      de.rsZanimanjaVeze.Update
    Next
    If Me.lstDelZanimanja.ListCount > 0 Then
      For i = 1 To Me.lstDelZanimanja.ListCount
        de.cn1.Execute "DELETE FROM ZanimanjaVeze WHERE RB=" & Me.lstDelZanimanja.List(i - 1)
      Next
      de.rsZanimanjaVeze.Requery
    End If
    If Me.lstRaspolozivost.ListItems.Count > 0 Then
      For i = 1 To Me.lstRaspolozivost.ListItems.Count
        If de.rsRaspolozivost.RecordCount > 0 Then
          de.rsRaspolozivost.MoveFirst
          de.rsRaspolozivost.Find "Naziv='" & Me.lstRaspolozivost.ListItems(i).Text & "'"
        End If
        If de.rsRaspolozivost.EOF Then
          de.rsRaspolozivost.AddNew
          de.rsRaspolozivost!Naziv = Me.lstRaspolozivost.ListItems(i).Text
          de.rsRaspolozivost.Update
        End If
        If Me.lstRaspolozivost.ListItems(i).Key = "" Then
          de.rsRaspolozivostVeze.AddNew
        Else
          de.rsRaspolozivostVeze.MoveFirst
          de.rsRaspolozivostVeze.Find "ID=" & Mid(Me.lstRaspolozivost.ListItems(i).Key, 3)
        End If
        de.rsRaspolozivostVeze!RB = de.rsDavaoci!RB
        de.rsRaspolozivostVeze!RaspolozivostID = de.rsRaspolozivost!ID
        de.rsRaspolozivostVeze.Update
      Next
    End If
    If Me.lstDelRaspolozivost.ListCount > 0 Then
      For i = 1 To Me.lstDelRaspolozivost.ListCount
        de.cn1.Execute "DELETE FROM RaspolozivostVeze WHERE RB=" & Me.lstDelRaspolozivost.List(i - 1)
      Next
      de.rsRaspolozivostVeze.Requery
    End If
    Me.txtRB.SetFocus
    Me.OKPressed = True
    Me.Hide
  End If

End Sub

Private Function CheckData() As Boolean
  CheckData = False
  If CStr(Val(Me.txtRB)) <> Me.txtRB Then
    MsgBox "Unesite ispravan redni broj davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtIme) = "" Then
    MsgBox "Unesite ime davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtPrezime) = "" Then
    MsgBox "Unesite prezime davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtRoditelj) = "" Then
    MsgBox "Unesite ime jednog roditelja!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtAdresa) = "" Then
    MsgBox "Unesite adresu davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtMjesto) = "" Then
    MsgBox "Unesite mjesto prebivališta davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtJMB) = "" Then
    MsgBox "Unesite JMB davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Trim(Me.txtBRLK) = "" Then
    MsgBox "Unesite broj liène karte davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  If Me.lstZanimanja.ListItems.Count = 0 Then
    MsgBox "Unesite bar jedno osnovno zanimanje davaoca usluga!", vbExclamation + vbOKOnly, Me.Caption
    Exit Function
  End If
  CheckData = True
End Function

Private Sub cmdCancel_Click()
  Me.txtRB.SetFocus
  Me.Hide
End Sub


Private Sub lstRaspolozivost_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim c As MSComctlLib.ListItem
  Dim sel As Boolean
  sel = False
  For Each c In Me.lstRaspolozivost.ListItems
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
    Set CurrControl = Me.lstRaspolozivost
    Me.PopupMenu mnuPopup, 2, x + Me.lstRaspolozivost.Left + Me.Frame1.Left, y + Me.lstRaspolozivost.Top + Me.Frame1.Top
  End If
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
    Me.PopupMenu mnuPopup, 2, x + Me.lstTelefoni.Left + Me.Frame1.Left, y + Me.lstTelefoni.Top + Me.Frame1.Top
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
    Me.PopupMenu mnuPopup, 2, x + Me.lstZanimanja.Left + Me.Frame1.Left, y + Me.lstZanimanja.Top + Me.Frame1.Top
  End If
End Sub


Private Sub mnuAdd_Click()
  Select Case CurrControl.Name
    Case "lstTelefoni"
      frmList.Caption = "Telefoni"
      frmList.lblNaziv.Caption = "Telefon:"
      frmList.cboVrijednost.Visible = False
      frmList.txtVrijednost.Visible = True
      frmList.txtVrijednost.ToolTipText = "Telefonski broj davaoca usluga"
      frmList.cmdAddEdit.Caption = "Dodaj"
      frmList.txtVrijednost.MaxLength = 15
      frmList.txtVrijednost = ""
      frmList.OKPressed = False
      frmList.Show 1, Me
      If frmList.OKPressed Then
        Me.lstTelefoni.ListItems.Add , , frmList.txtVrijednost
      End If
    Case "lstZanimanja"
      frmList.Caption = "Zanimanja"
      frmList.lblNaziv.Caption = "Zanimanje:"
      frmList.cboVrijednost.Visible = True
      frmList.cboVrijednost.ToolTipText = "Zanimanje davaoca usluga"
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
    Case "lstRaspolozivost"
      frmList.Caption = "Raspoloživosti"
      frmList.lblNaziv.Caption = "Raspoloživost:"
      frmList.cboVrijednost.Visible = True
      frmList.cboVrijednost.ToolTipText = "Raspoloživost davaoca usluga"
      frmList.txtVrijednost.Visible = False
      frmList.cmdAddEdit.Caption = "Dodaj"
      frmList.cboVrijednost.Clear
      If de.rsRaspolozivost.RecordCount > 0 Then
        de.rsRaspolozivost.MoveFirst
        Do Until de.rsRaspolozivost.EOF
          frmList.cboVrijednost.AddItem de.rsRaspolozivost!Naziv
          de.rsRaspolozivost.MoveNext
        Loop
      End If
      frmList.OKPressed = False
      frmList.Show 1, Me
      If frmList.OKPressed Then
        Me.lstRaspolozivost.ListItems.Add , , frmList.cboVrijednost.Text
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
    Case "lstZanimanja"
      a = MsgBox("Da li želite obrisati izabrano zanimanje?", vbQuestion + vbYesNo, Me.Caption)
      If a = vbYes Then
        If Left(Me.lstZanimanja.SelectedItem.Key, 2) = "ID" Then Me.lstDelZanimanja.AddItem Mid(Me.lstZanimanja.SelectedItem.Key, 3)
        Me.lstZanimanja.ListItems.Remove Me.lstZanimanja.SelectedItem.Index
      End If
    Case "lstRaspolozivost"
      a = MsgBox("Da li želite obrisati izabranu raspolozivost?", vbQuestion + vbYesNo, Me.Caption)
      If a = vbYes Then
        If Left(Me.lstRaspolozivost.SelectedItem.Key, 2) = "ID" Then Me.lstDelRaspolozivost.AddItem Mid(Me.lstRaspolozivost.SelectedItem.Key, 3)
        Me.lstRaspolozivost.ListItems.Remove Me.lstRaspolozivost.SelectedItem.Index
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
      frmList.txtVrijednost.ToolTipText = "Telefonski broj davaoca usluga"
      frmList.cmdAddEdit.Caption = "Izmijeni"
      frmList.txtVrijednost.MaxLength = 15
      frmList.txtVrijednost = Me.lstTelefoni.SelectedItem.Text
      frmList.OKPressed = False
      frmList.Show 1, Me
      If frmList.OKPressed Then
        Me.lstTelefoni.SelectedItem.Text = frmList.txtVrijednost
      End If
    Case "lstZanimanja"
      frmList.Caption = "Zanimanja"
      frmList.lblNaziv.Caption = "Zanimanje:"
      frmList.cboVrijednost.Visible = True
      frmList.cboVrijednost.ToolTipText = "Zanimanje davaoca usluga"
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
      frmList.cboVrijednost.Text = Me.lstZanimanja.SelectedItem.Text
      frmList.OKPressed = False
      frmList.Show 1, Me
      If frmList.OKPressed Then
        Me.lstZanimanja.SelectedItem.Text = frmList.cboVrijednost.Text
      End If
    Case "lstRaspolozivost"
      frmList.Caption = "Raspoloživosti"
      frmList.lblNaziv.Caption = "Raspoloživost:"
      frmList.cboVrijednost.Visible = True
      frmList.cboVrijednost.ToolTipText = "Raspoloživost davaoca usluga"
      frmList.txtVrijednost.Visible = False
      frmList.cmdAddEdit.Caption = "Izmijeni"
      frmList.cboVrijednost.Clear
      If de.rsRaspolozivost.RecordCount > 0 Then
        de.rsRaspolozivost.MoveFirst
        Do Until de.rsRaspolozivost.EOF
          frmList.cboVrijednost.AddItem de.rsRaspolozivost!Naziv
          de.rsRaspolozivost.MoveNext
        Loop
      End If
      frmList.cboVrijednost.Text = Me.lstRaspolozivost.SelectedItem.Text
      frmList.OKPressed = False
      frmList.Show 1, Me
      If frmList.OKPressed Then
        Me.lstRaspolozivost.SelectedItem.Text = frmList.cboVrijednost.Text
      End If
  End Select
End Sub

Private Sub txtAdresa_GotFocus()
  SelAll
End Sub

Private Sub txtBRLK_GotFocus()
  SelAll
End Sub

Private Sub txtIme_GotFocus()
  SelAll
End Sub

Private Sub SelAll()
  Me.ActiveControl.SelStart = 0
  Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtJMB_GotFocus()
  SelAll
End Sub

Private Sub txtKomentar_GotFocus()
  SelAll
End Sub


Private Sub txtMjesto_GotFocus()
  SelAll
End Sub

Private Sub txtPrezime_GotFocus()
  SelAll
End Sub

Private Sub txtRB_GotFocus()
  SelAll
End Sub

Private Sub txtRoditelj_GotFocus()
  SelAll
End Sub
