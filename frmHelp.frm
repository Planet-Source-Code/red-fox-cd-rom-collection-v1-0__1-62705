VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD-ROM Collection V1.0 Help"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbHelp 
      Height          =   3855
      Left            =   480
      TabIndex        =   2
      Top             =   5640
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6800
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmHelp.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CommadButton1 
      BackColor       =   &H0000C0C0&
      Caption         =   "MAIN"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton CommandButton2 
      BackColor       =   &H0000C0C0&
      Caption         =   "VIEW"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   4755
      Left            =   480
      Top             =   120
      Width           =   6330
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommadButton1_Click()
On Error GoTo err_comp
    Image1.Picture = LoadPicture(App.Path & "\images\main.bmp")
    Image1.Left = (frmHelp.ScaleWidth - Image1.Width) / 2
    rtbHelp.FileName = App.Path & "\help\main.rtf"
    Exit Sub
err_comp:
    MsgBox "ERROR LOADING ONE OF ITS COMPONENTS", vbCritical, "CD-ROM Collection V1.0"
    rtbHelp.Text = "ERROR LOADING ONE OF ITS COMPONENTS"
End Sub

Private Sub CommandButton2_Click()
On Error GoTo err_comp
    Image1.Picture = LoadPicture(App.Path & "\images\view.bmp")
    Image1.Left = (frmHelp.ScaleWidth - Image1.Width) / 2
    rtbHelp.FileName = App.Path & "\help\view.rtf"
    Exit Sub
err_comp:
    MsgBox "ERROR LOADING ONE OF ITS COMPONENTS", vbCritical, "CD-ROM Collection V1.0"
    rtbHelp.Text = "ERROR LOADING ONE OF ITS COMPONENTS"
End Sub

Private Sub Form_Load()
On Error GoTo err_comp
    Image1.Picture = LoadPicture(App.Path & "\images\main.bmp")
    Image1.Left = (frmHelp.ScaleWidth - Image1.Width) / 2
    rtbHelp.FileName = App.Path & "\help\main.rtf"
    Exit Sub
err_comp:
    MsgBox "ERROR LOADING ONE OF ITS COMPONENTS", vbCritical, "CD-ROM Collection V1.0"
    rtbHelp.Text = "ERROR LOADING ONE OF ITS COMPONENTS"
End Sub
