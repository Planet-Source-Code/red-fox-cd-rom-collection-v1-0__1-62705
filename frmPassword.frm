VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD-ROM Collection V1.0"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000C0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0000C0C0&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload frmPassword
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
    If txtPassword.Text = "4285116844eVeR!" Then
        Unload Me
        frmMain.Show
    Else
        MsgBox "INVALID PASSWORD", vbCritical, "CD-ROM Collection V1.0"
    End If
End Sub

