VERSION 5.00
Begin VB.Form frmZR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Žiro-raèuni"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
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
   ScaleHeight     =   2670
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAddEdit 
      Caption         =   "Dodaj"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtBanka 
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   5
      ToolTipText     =   "Banka u kojoj se vodi žiro-raèun"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox cboTip 
      Height          =   315
      ItemData        =   "frmZR.frx":0000
      Left            =   120
      List            =   "frmZR.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Tip žiro-raèuna"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtBroj 
      Height          =   285
      Left            =   120
      MaxLength       =   30
      TabIndex        =   1
      ToolTipText     =   "Broj žiro-raèuna"
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Banka:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Tip:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Žiro-raèun:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmZR"
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

Private Sub cmdAddEdit_Click()
  If Trim(Me.txtBroj) = "" Then Exit Sub
  Me.txtBroj.SetFocus
  Me.OKPressed = True
  Me.Hide
End Sub

Private Sub cmdCancel_Click()
  Me.txtBroj.SetFocus
  Me.Hide
End Sub

Private Sub Form_Load()
  Me.cboTip.ListIndex = 0
End Sub

Private Sub SelAll()
  Me.ActiveControl.SelStart = 0
  Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub txtBanka_GotFocus()
  SelAll
End Sub

Private Sub txtBroj_GotFocus()
  SelAll
End Sub
