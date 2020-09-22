VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Char <=> Asc"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Convert ASCII to Character"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2895
      Begin VB.TextBox txtChr 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdChr 
         Caption         =   "Get Char"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Convert Character to ASCII"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtAsc 
         Height          =   285
         Left            =   120
         MaxLength       =   1
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAsc 
         Caption         =   "Get ASCII"
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit

Private Sub cmdAsc_Click()
On Error GoTo ErrHandler
    MsgBox Asc(txtAsc.Text), vbInformation, "Char <=> ASCII"
    txtAsc.SetFocus
    Exit Sub
ErrHandler:
    MsgBox "Invalid Code", vbExclamation, "Char <=> ASCII"
End Sub

Private Sub cmdChr_Click()
On Error GoTo ErrHandler
    MsgBox Chr(txtChr.Text), vbInformation, "Char <=> ASCII"
    txtChr.SetFocus
    Exit Sub
ErrHandler:
    MsgBox "Invalid Code", vbExclamation, "Char <=> ASCII"
End Sub


Private Sub Form_Load()
    If App.PrevInstance Then
        Unload Me
    End If
End Sub

Private Sub txtAsc_GotFocus()
    cmdAsc.Default = True
End Sub

Private Sub txtChr_GotFocus()
    cmdChr.Default = True
End Sub
