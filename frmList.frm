VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   ControlBox      =   0   'False
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
   ScaleHeight     =   1200
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtVrijednost 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox cboVrijednost 
      Height          =   315
      ItemData        =   "frmList.frx":0000
      Left            =   120
      List            =   "frmList.frx":0002
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddEdit 
      Caption         =   "Dodaj"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblNaziv 
      Caption         =   "Telefon:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmList"
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

Private Sub cboVrijednost_GotFocus()
  SelAll
End Sub

Private Sub cmdAddEdit_Click()
  If Me.txtVrijednost.Visible Then
    If Trim(Me.txtVrijednost) = "" Then Exit Sub
  Else
    If Trim(Me.cboVrijednost) = "" Then Exit Sub
  End If
  Me.OKPressed = True
  If Me.txtVrijednost.Visible Then
    Me.txtVrijednost.SetFocus
    Me.txtVrijednost.TabIndex = 0
  Else
    Me.cboVrijednost.SetFocus
    Me.cboVrijednost.TabIndex = 0
  End If
  Me.Hide
End Sub

Private Sub cmdCancel_Click()
  If Me.txtVrijednost.Visible Then
    Me.txtVrijednost.SetFocus
    Me.txtVrijednost.TabIndex = 0
  Else
    Me.cboVrijednost.SetFocus
    Me.cboVrijednost.TabIndex = 0
  End If
  Me.Hide
End Sub

Private Sub SelAll()
  Me.ActiveControl.SelStart = 0
  Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtVrijednost_GotFocus()
  SelAll
End Sub
