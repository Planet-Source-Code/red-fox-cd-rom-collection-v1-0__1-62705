VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD-ROM Collection V1.0"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameView 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   960
      TabIndex        =   26
      Top             =   360
      Width           =   615
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000C0C0&
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   960
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmMain.frx":0CCA
         Height          =   5175
         Left            =   360
         TabIndex        =   30
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   9128
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "frmMain.frx":0CDF
         Left            =   2760
         List            =   "frmMain.frx":0CE1
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Division:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "CD-ROM COLLECTIONS"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   27
         Top             =   240
         Width           =   8415
      End
   End
   Begin VB.Frame frameMain 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   495
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "frmMain.frx":0CE3
         Left            =   6960
         List            =   "frmMain.frx":0CE5
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CommandButton cmdBottom 
         BackColor       =   &H0000C0C0&
         Caption         =   "BOTTOM"
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
         Left            =   5280
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0000C0C0&
         Caption         =   "NEXT"
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
         Left            =   3720
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0000C0C0&
         Caption         =   "PREVIOUS"
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
         Left            =   2160
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton cmdTop 
         BackColor       =   &H0000C0C0&
         Caption         =   "TOP"
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
         Left            =   600
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0000C0C0&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5640
         Width           =   855
      End
      Begin VB.TextBox txtPCNum 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   20
         Text            =   "0"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox txtQty 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "0"
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox txtCDNum 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "0"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtCallNum 
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
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox txtAuthor 
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
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtTitle 
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
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1560
         Width           =   3855
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmMain.frx":0CE7
         Left            =   3000
         List            =   "frmMain.frx":0CE9
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   3855
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H0000C0C0&
         Caption         =   "DELETE"
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
         Left            =   5280
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelEdit 
         BackColor       =   &H0000C0C0&
         Caption         =   "EDIT"
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
         Left            =   3720
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000C0C0&
         Caption         =   "ADD"
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
         Left            =   2160
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H0000C0C0&
         Caption         =   "NEW"
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
         Left            =   600
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "by:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   7920
         TabIndex        =   31
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "PC Num.:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CD-ROM Num.:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Call Number:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Division:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CD-ROM COLLECTIONS"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Menu menuPrint 
      Caption         =   "PRINT"
      Begin VB.Menu menuPrintThis 
         Caption         =   "This Only"
      End
      Begin VB.Menu menuPrintFaculty 
         Caption         =   "Faculty"
      End
      Begin VB.Menu menuPrintStudent 
         Caption         =   "Student"
      End
   End
   Begin VB.Menu menuView 
      Caption         =   "VIEW"
      Begin VB.Menu menuViewFaculty 
         Caption         =   "Faculty"
      End
      Begin VB.Menu menuViewStudent 
         Caption         =   "Student"
      End
      Begin VB.Menu menuViewMain 
         Caption         =   "Main"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "HELP"
      Begin VB.Menu menuHelpHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu menuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'**************************************************************************************
'**                     Program name: CD-ROM Collection V1.0                         **
'**             Created for: Ateneo de Davao Univerisity Library                     **
'**                 Programmer: Junie L. Lorenzo A.K.A Red_Fox                       **
'**         Comment: A simple database using ADODC with SQL commands                 **
'**************************************************************************************
'**************************************************************************************

'Global variables = oCn  - is the ADODC connection = required when you
'                   need to connect to a ADODC
'                 = oRs  - is holder of the records = required when you need
'                   handle the data base of the ADODC
'                 = oRsPrint  - is also a holder for records but this one is used for
'                   printing purposes
'                 = sf - a string variable that will determine if you want to view all
'                   the records within the student division or in the faculty division
'                 = sfPrint - also a string that will determine if it is in faculty or
'                   student division but this time it is for printing
'                 = globalDiv - string variable for printing purposes
'Local variables  = msgVerify - to counter check the user's input if it's ok for him to
'                   proceed
'                 = i - just a counter
'                 = sTemp - will hold the field value that has enough width in the printing
'                   area
'                 = sHold - will hold the concatenation of some fields
'                 = myTextWidth - will determine  the width of the text to be print
'                 = yPrint - for determining the y-axis in printing
'                 = strLen - will get the number of characters within a string
'                 = next_div - counter for division
'                 = get_div - will get the division's name from the get_div_name function
'Functions        = get_div_name() - will determine the division's name
'                 = print_rec() - prints the output of the "CD-ROM Collection V1.0"
'                 = fill_in() - fill outs form

Option Explicit
Dim oCn As ADODB.Connection
Dim oRs As ADODB.Recordset
Dim oRsPrint As New ADODB.Recordset
Dim sf As String
Dim sfPrint As String
Dim globalDiv As String

Private Sub cmdAdd_Click() 'adding a record
Dim msgVerify As VbMsgBoxResult
    If txtTitle.Text = "" Then 'if the user didn't put any title
        txtTitle.Text = " "
    End If
    If txtAuthor.Text = "" Then 'if the user didn't put any author name
        txtAuthor.Text = " "
    End If
    If txtCallNum.Text = "" Then 'if the user didn't put any call number
        txtCallNum.Text = " "
    End If
    If txtQty.Text = "0" Or Not IsNumeric(txtQty.Text) Then 'if the user puts zero
        txtQty.Text = "1"                                   'or any invalid character
    End If                                                  'except for a number
    If Not IsNumeric(txtCDNum.Text) Then 'if the user puts zero
        txtCDNum.Text = "0"                                     'or any invalid character
    End If                                                      'except for a number
    If txtPCNum.Text = "" Then 'if the user didn't put any PC number
        txtPCNum.Text = " "
    End If
    msgVerify = MsgBox("Is this ok?", vbYesNo, "CD-ROM Collection V1.0")
    If msgVerify = vbYes Then 'verify first the inputted values
        oRs.AddNew 'add now to the database
        oRs("fldDivision") = Combo1.Text
        oRs("fldsf") = Combo2.Text
        oRs("fldTitle") = txtTitle.Text
        oRs("fldAuthor") = txtAuthor.Text
        oRs("fldCallnum") = txtCallNum.Text
        oRs("fldQty") = CInt(txtQty.Text)
        oRs("fldCDROMnum") = Val(txtCDNum.Text)
        oRs("fldPCnum") = txtPCNum.Text
        oRs.Update 'end of adding records and updating the oRs
        oRs.Close
        Set oRs = Nothing 'return to Accounting - Student division
        MsgBox "New record has been added to " & Combo1.Text, vbExclamation, "CD-ROM Collection V1.0"
        Set oRs = New ADODB.Recordset
        oRs.Open "SELECT * FROM tblCDROM WHERE fldDivision = '" & Combo1.Text & "'" & "AND fldsf = '" & _
        Combo2.Text & "' ORDER BY fldCDROMnum ASC ;", oCn, adOpenKeyset, adLockOptimistic, adCmdText
        oRs.MoveFirst
        cmdSearch.Enabled = True
        Call fill_in
        cmdNew.Enabled = True 'loads all the default command buttons
        cmdAdd.Enabled = False
        cmdCancelEdit.Caption = "EDIT"
        cmdDelete.Enabled = True
        cmdTop.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdBottom.Enabled = True
    End If
End Sub

Private Sub cmdBottom_Click()
    oRs.MoveLast 'go to the last record
    Call fill_in
    cmdPrev.Enabled = True
    cmdTop.Enabled = True
    cmdNext.Enabled = False
    cmdBottom.Enabled = False
End Sub

Private Sub cmdCancelEdit_Click() 'cancel or for editting
Dim i As Integer
Dim msgVerify As VbMsgBoxResult
    If cmdCancelEdit.Caption = "EDIT" Then 'if cmdCancelEdit.Caption = "EDIT"
        If txtTitle.Text = "" Then 'if the user didn't put any title
            txtTitle.Text = " "
        End If
        If txtAuthor.Text = "" Then 'if the user didn't put any author name
            txtAuthor.Text = " "
        End If
        If txtCallNum.Text = "" Then 'if the user didn't put any call number
            txtCallNum.Text = " "
        End If
        If txtQty.Text = "0" Or Not IsNumeric(txtQty.Text) Then 'if the user puts zero
            txtQty.Text = "1"                                   'or any invalid character
        End If                                                  'except for a number
        If Not IsNumeric(txtCDNum.Text) Then 'if the user puts zero
            txtCDNum.Text = "0"                                     'or any invalid character
        End If                                                      'except for a number
        If txtPCNum.Text = "" Then 'if the user didn't put any PC number
            txtPCNum.Text = " "
        End If
        msgVerify = MsgBox("Is this ok?", vbYesNo, "CD-ROM Collection V1.0")
        If msgVerify = vbYes Then  'verify first the inputted values
            oRs("fldDivision") = Combo1.Text 'editting now the record
            oRs("fldsf") = Combo2.Text
            oRs("fldTitle") = txtTitle.Text
            oRs("fldAuthor") = txtAuthor.Text
            oRs("fldCallnum") = txtCallNum.Text
            oRs("fldQty") = CInt(txtQty.Text)
            oRs("fldCDROMnum") = Val(txtCDNum.Text)
            oRs("fldPCnum") = txtPCNum.Text
            oRs.Update 'end of editting and updating oRs
            MsgBox "One record has been edited", vbExclamation, "CD-ROM Collection V1.0"
            oRs.Close
            Set oRs = Nothing 'return to Accounting - Student division
            Set oRs = New ADODB.Recordset
            oRs.Open "SELECT * FROM tblCDROM WHERE fldDivision = 'Accounting' AND fldsf = 'Student' ORDER BY fldCDROMnum ASC ; ", oCn, adOpenKeyset, adLockOptimistic, adCmdText
            oRs.MoveFirst
            cmdSearch.Enabled = True
            Call fill_in
            cmdNew.Enabled = True 'loads all the default command buttons
            cmdAdd.Enabled = False
            cmdCancelEdit.Caption = "EDIT"
            cmdDelete.Enabled = True
            cmdTop.Enabled = False
            cmdPrev.Enabled = False
            cmdNext.Enabled = True
            cmdBottom.Enabled = True
        End If
    End If
    If cmdCancelEdit.Caption = "CANCEL" Then 'if cmdCancelEdit.Caption = "CANCEL"
        cmdSearch.Enabled = True             'loads back to previous settings
        cmdNew.Enabled = True
        cmdAdd.Enabled = False
        cmdCancelEdit.Caption = "EDIT"
        cmdDelete.Enabled = True
        cmdTop.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdBottom.Enabled = True
        oRs.MoveFirst
        Call fill_in
    End If
    
End Sub

Private Sub cmdDelete_Click() 'deleting a record
Dim msgVerify As VbMsgBoxResult
    msgVerify = MsgBox("Are you sure you want to delete this record?", vbYesNo, "CD-ROM Collection V1.0")
    If msgVerify = vbYes Then 'verify first if it's ok to delete a record
        oRs.Delete            'deleting of record
        MsgBox "One record has been deleted", vbCritical, "CD-ROM Collection V1.0"
    End If                    'end of deleting
    oRs.Close
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset 'return to Accounting - Student division
    oRs.Open "SELECT * FROM tblCDROM WHERE fldDivision = 'Accounting' AND fldsf = 'Student' ORDER BY fldCDROMnum ASC ; ", oCn, adOpenKeyset, adLockOptimistic, adCmdText
    oRs.MoveFirst
    Call fill_in
    cmdNew.Enabled = True 'loads all the default command buttons
    cmdAdd.Enabled = False
    cmdCancelEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdTop.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = True
    cmdBottom.Enabled = True
    cmdCancelEdit.Caption = "EDIT"
End Sub

Private Sub cmdNew_Click() 'loads a default form for a new record
    cmdSearch.Enabled = False
    cmdNew.Enabled = False
    cmdAdd.Enabled = True
    cmdCancelEdit.Caption = "CANCEL"
    cmdDelete.Enabled = False
    cmdTop.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdBottom.Enabled = False
    txtTitle.Text = ""
    txtAuthor.Text = ""
    txtCallNum.Text = ""
    txtCDNum.Text = 0
    txtQty.Text = 1
    txtPCNum.Text = ""
End Sub

Private Sub cmdNext_Click() 'next record
    oRs.MoveNext
    If oRs.EOF Then
        oRs.MoveLast
        cmdNext.Enabled = False
        cmdBottom.Enabled = False
    End If
    Call fill_in
    cmdPrev.Enabled = True
    cmdTop.Enabled = True
End Sub

Private Sub cmdPrev_Click() 'previous record
    oRs.MovePrevious
    If oRs.BOF Then
        oRs.MoveFirst
        cmdPrev.Enabled = False
        cmdTop.Enabled = False
    End If
    Call fill_in
    cmdNext.Enabled = True
    cmdBottom.Enabled = True
End Sub

Private Sub cmdSearch_Click() 'search a record
    If Combo4.Text = "Author" Then 'searching for a author
        oRs.Close
        Set oRs = Nothing
        Set oRs = New ADODB.Recordset
        If txtAuthor.Text = "" Then
            txtAuthor.Text = " "
        End If
On Error GoTo err_main 'if no record/s found go to error handler
        oRs.Open "SELECT * FROM tblCDROM WHERE fldAuthor = '" & txtAuthor.Text & "'" & _
        " ORDER BY fldCDROMnum ASC ;", oCn, adOpenKeyset, adLockOptimistic, adCmdText
        oRs.MoveFirst
        Call fill_in
        cmdNew.Enabled = True
        cmdAdd.Enabled = False
        cmdCancelEdit.Caption = "EDIT"
        cmdDelete.Enabled = True
        cmdTop.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdBottom.Enabled = True
    End If
    
    If Combo4.Text = "Division" Then 'searching for a division
        oRs.Close
        Set oRs = Nothing
        Set oRs = New ADODB.Recordset
On Error GoTo err_main 'if no record/s found go to error handler
        oRs.Open "SELECT * FROM tblCDROM WHERE fldDivision = '" & Combo1.Text & "'" & "AND fldsf = '" & _
        Combo2.Text & "' ORDER BY fldCDROMnum ASC ;", oCn, adOpenKeyset, adLockOptimistic, adCmdText
        oRs.MoveFirst
        Call fill_in
        cmdNew.Enabled = True
        cmdAdd.Enabled = False
        cmdCancelEdit.Caption = "EDIT"
        cmdDelete.Enabled = True
        cmdTop.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdBottom.Enabled = True
    End If
        
    If Combo4.Text = "Call Number" Then 'searching for a call number
        oRs.Close
        Set oRs = Nothing
        Set oRs = New ADODB.Recordset
        If txtCallNum.Text = "" Then
            txtCallNum.Text = " "
        End If
On Error GoTo err_main 'if no record/s found go to error handler
        oRs.Open "SELECT * FROM tblCDROM WHERE fldCallnum = '" & txtCallNum.Text & "'" & _
        " ORDER BY fldCDROMnum ASC ;", oCn, adOpenKeyset, adLockOptimistic, adCmdText
        oRs.MoveFirst
        Call fill_in
        cmdNew.Enabled = True
        cmdAdd.Enabled = False
        cmdCancelEdit.Caption = "EDIT"
        cmdDelete.Enabled = True
        cmdTop.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdBottom.Enabled = True
    End If
    
    If Combo4.Text = "Title" Then 'searching for a title
        oRs.Close
        Set oRs = Nothing
        Set oRs = New ADODB.Recordset
        If txtTitle.Text = "" Then
            txtTitle.Text = " "
        End If
On Error GoTo err_main 'if no record/s found go to error handler
        oRs.Open "SELECT * FROM tblCDROM WHERE fldTitle = '" & txtTitle.Text & "'" & _
        " ORDER BY fldCDROMnum ASC ;", oCn, adOpenKeyset, adLockOptimistic, adCmdText
        oRs.MoveFirst
        Call fill_in
        cmdNew.Enabled = True
        cmdAdd.Enabled = False
        cmdCancelEdit.Caption = "EDIT"
        cmdDelete.Enabled = True
        cmdTop.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdBottom.Enabled = True
    End If
    Exit Sub
err_main: ' error handler
    MsgBox "No record/s found", vbCritical, "CD-ROM Collection V1.0"
    oRs.Close
    Set oRs = Nothing 'return to Accounting - Student division
    Set oRs = New ADODB.Recordset
    oRs.Open "SELECT * FROM tblCDROM WHERE fldDivision = 'Accounting' AND fldsf = 'Student' ORDER BY fldCDROMnum ASC ; ", oCn, adOpenKeyset, adLockOptimistic, adCmdText
    Call fill_in
    cmdNew.Enabled = True 'loads all the default command buttons
    cmdAdd.Enabled = False
    cmdCancelEdit.Caption = "EDIT"
    cmdDelete.Enabled = True
    cmdTop.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = True
    cmdBottom.Enabled = True
End Sub

Private Sub cmdTop_Click() 'go to the top of record
    oRs.MoveFirst
    Call fill_in
    cmdNext.Enabled = True
    cmdBottom.Enabled = True
    cmdPrev.Enabled = False
    cmdTop.Enabled = False
End Sub

Private Sub Command1_Click() 'viewing all records within a division
    oRs.Close
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
On Error GoTo err_view 'if no record within this division go to error handler
    oRs.CursorLocation = adUseClient
    oRs.Open "SELECT fldTitle, fldAuthor, fldCallnum, fldQty, fldCDROMnum, fldPCnum" & _
    " FROM tblCDROM WHERE fldDivision = '" & Combo3.Text & "'" & "AND fldsf = '" & _
    sf & "' ORDER BY fldCDROMnum ASC ;", oCn, adOpenKeyset, adLockOptimistic
    oRs.MoveFirst
    Set DataGrid1.DataSource = oRs
    Exit Sub
    
err_view: 'error handler
    MsgBox "No record/s within this division", vbCritical, "CD-ROM Collection V1.0"
    oRs.Close
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset 'return to Accounting - Student division
    oRs.CursorLocation = adUseClient
    oRs.Open "SELECT fldTitle, fldAuthor, fldCallnum, fldQty, fldCDROMnum, fldPCnum" & _
    " FROM tblCDROM WHERE fldDivision = 'Accounting'" & "AND fldsf = '" & _
    sf & "' ORDER BY fldCDROMnum ASC ;", oCn, adOpenKeyset, adLockOptimistic
    oRs.MoveFirst
    Combo3.Text = "Accounting"
    Set DataGrid1.DataSource = oRs
End Sub

Private Sub Form_Load()
Dim str_password As String
    Call load_main
    Set oCn = New ADODB.Connection
    Set oRs = New ADODB.Recordset
On Error GoTo err_out
    oCn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source= " & App.Path & ("\cd-rom.mdb") & ";Persist Security Info=False"
    oRs.Open "SELECT * FROM tblCDROM WHERE fldDivision = 'Accounting' AND fldsf = 'Student' ORDER BY fldCDROMnum ASC ; ", oCn, adOpenKeyset, adLockOptimistic, adCmdText
    oRsPrint.Open "SELECT * FROM tblCDROM", oCn, adOpenKeyset, adLockOptimistic, adCmdText
    Call fill_in
    Exit Sub
err_out:
    MsgBox "Error loading database. Program will be shut down", vbCritical, "CD-ROM Collection V1.0"
    Unload frmMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    oRsPrint.MoveFirst
    oRs.MoveFirst
    oRs.Close
    oRsPrint.Close
    oCn.Close
    Set oRs = Nothing
    Set oCn = Nothing
    Set oRsPrint = Nothing
End Sub

Private Sub menuHelpAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub menuHelpHelp_Click()
    frmHelp.Show 1
End Sub

Private Sub menuPrintFaculty_Click()
    Dim msgVerify As VbMsgBoxResult
    msgVerify = MsgBox("Print to " & Printer.DeviceName & " ?", vbYesNo, "CD-ROM Collection V1.0")
    If msgVerify = vbYes Then
        sfPrint = "Faculty"
        Call print_rec(sfPrint)
    End If
End Sub

Private Sub menuPrintStudent_Click()
    Dim msgVerify As VbMsgBoxResult
    msgVerify = MsgBox("Print to " & Printer.DeviceName & " ?", vbYesNo, "CD-ROM Collection V1.0")
    If msgVerify = vbYes Then
        sfPrint = "Student"
        Call print_rec(sfPrint)
    End If
End Sub

Private Sub menuPrintThis_Click()
    Dim msgVerify As VbMsgBoxResult
    On Error GoTo err_print
    msgVerify = MsgBox("Print to " & Printer.DeviceName & " ?", vbYesNo, "CD-ROM Collection V1.0")
    If msgVerify = vbYes Then
        Printer.ScaleMode = vbCentimeters
        Printer.FontName = "Courier New"
        Printer.FontBold = False
        Printer.FontSize = 12
        Printer.CurrentY = 0.5
        Printer.Print "Division: " & oRs("fldDivision"); Spc(2); "- " & oRs("fldsf")
        Printer.Print ""
        Printer.Print "Title: " & oRs("fldTitle")
        Printer.Print ""
        If oRs("fldQty") > 1 Then
        Printer.Print "Author: " & oRs("fldAuthor")
            Printer.Print "Call number: " & oRs("fldCallnum"); Spc(2); "[" & Str(oRs("fldQty")) & " copies]"
        Else
            Printer.Print "Call number: " & oRs("fldCallnum")
        End If
        Printer.Print ""
        Printer.Print "CD-ROM #: " & oRs("fldCDROMnum"); " PC" & oRs("fldPCnum")
        Printer.EndDoc
        Exit Sub
    End If
err_print:
    MsgBox "Sorry, printer is not available at this time", vbCritical, "CD-ROM Collection V1.0"
End Sub

Private Sub menuViewFaculty_Click()
    oRs.MoveFirst
    oRs.Close
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
    oRs.CursorLocation = adUseClient
    oRs.Open "SELECT fldTitle, fldAuthor, fldCallnum, fldQty, fldCDROMnum, fldPCnum" & _
    " FROM tblCDROM WHERE fldDivision = 'Accounting' AND fldsf = 'Faculty' ORDER BY fldCDROMnum ASC ; ", oCn, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = oRs
    Combo3.Text = "Accounting"
    sf = "Faculty"
    frameMain.Visible = False
    frameView.Visible = True
    Label16.Caption = "Faculty CD-ROM Collections"
End Sub

Private Sub menuViewMain_Click()
    Call load_main
    Set oRs = New ADODB.Recordset
    oRs.Open "SELECT * FROM tblCDROM WHERE fldDivision = 'Accounting' AND fldsf = 'Student' ORDER BY fldCDROMnum ASC ; ", oCn, adOpenKeyset, adLockOptimistic, adCmdText
    Call fill_in
    frameMain.Visible = True
    frameView.Visible = False
    menuPrintThis.Enabled = True
End Sub

Private Sub menuViewStudent_Click()
    oRs.MoveFirst
    oRs.Close
    Set oRs = Nothing
    Set oRs = New ADODB.Recordset
    oRs.CursorLocation = adUseClient
    oRs.Open "SELECT fldTitle, fldAuthor, fldCallnum, fldQty, fldCDROMnum, fldPCnum" & _
    " FROM tblCDROM WHERE fldDivision = 'Accounting' AND fldsf = 'Student' ORDER BY fldCDROMnum ASC ; ", oCn, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = oRs
    sf = "Student"
    Combo3.Text = "Accounting"
    frameMain.Visible = False
    frameView.Visible = True
    Label16.Caption = "Student CD-ROM Collections"
    menuPrintThis.Enabled = False
End Sub

Function fill_in()
    Combo1.Text = oRs("fldDivision")
    Combo2.Text = oRs("fldsf")
    txtTitle.Text = oRs("fldTitle")
    txtAuthor.Text = oRs("fldAuthor")
    txtCallNum.Text = oRs("fldCallnum")
    txtCDNum.Text = oRs("fldCDROMnum")
    txtQty.Text = oRs("fldQty")
    txtPCNum.Text = oRs("fldPCnum")
End Function

Function print_rec(fld_div As String)
Dim sTemp, sHold As String
Dim i, myTextWidth As Integer
Dim yPrint, strlen As Integer
Dim next_div As Integer
Dim get_div As String
Dim ctr_let As String
Dim i_ctr As Integer
On Error GoTo err_print
    Printer.Print ""
    next_div = 1
    i_ctr = 0
    ctr_let = "abcdefghijk"
    Do While next_div < 27
        get_div = get_div_name(next_div)
        oRsPrint.Close
        Set oRsPrint = Nothing
        Set oRsPrint = New ADODB.Recordset
On Error GoTo err_next_div
        oRsPrint.CursorLocation = adUseClient
        oRsPrint.Open "SELECT fldTitle, fldAuthor, fldCallnum, fldQty, fldCDROMnum, fldPCnum" & _
        " FROM tblCDROM WHERE fldDivision = '" & get_div & "' AND fldsf = '" & fld_div & "' ORDER BY fldCDROMnum ASC ;", oCn, adOpenKeyset, adLockOptimistic, adCmdText
        If oRsPrint.EOF Then
            GoTo err_next_div
        End If
'if division = PURE SCIENCES
        If next_div >= 14 And next_div <= 19 Then
            If i_ctr = 0 Then
                Printer.ScaleMode = vbCentimeters
                Printer.FontName = "Comic Sans MS"
                Printer.FontBold = True
                Printer.FontSize = 14
                myTextWidth = Printer.TextWidth("PURE SCIENCES")
                Printer.CurrentX = ((Printer.ScaleWidth - myTextWidth) / 2) - 0.5
                Printer.CurrentY = 1
                Printer.Print "PURE SCIENCES"
                Printer.CurrentY = Printer.CurrentY + 0.5
                Printer.CurrentX = 0.5
                yPrint = Printer.CurrentY
                i_ctr = i_ctr + 1
            End If
            If i_ctr > 0 Then
                Printer.ScaleMode = vbCentimeters
                Printer.CurrentX = 0.5
                If yPrint >= 25 Then
                    Printer.NewPage
                    yPrint = Printer.CurrentY + 0.5
                Else
                    Printer.CurrentY = yPrint
                End If
                Printer.FontName = "Comic Sans MS"
                Printer.FontBold = True
                Printer.FontSize = 11
                Printer.Print Mid(ctr_let, i_ctr, 1) & ". " & globalDiv 'sample output =
                Printer.CurrentX = 0.5                                  'a. Biology
                yPrint = Printer.CurrentY                               'where ctr_let is a
                i_ctr = i_ctr + 1                                       'repository of letters
            End If
            Printer.FontName = "Times New Roman"
            Printer.FontBold = False
            Printer.FontSize = 11
            Printer.Print "Title"
            Printer.CurrentY = yPrint
            Printer.CurrentX = 8
            Printer.Print "Author"
            Printer.CurrentY = yPrint
            Printer.CurrentX = 12
            Printer.Print "Call Number"
            Printer.CurrentY = yPrint
            Printer.CurrentX = 16.5
            Printer.Print "CD-ROM #"
            yPrint = Printer.CurrentY + 0.25
        End If
'if division = SOCIAL SCIENCES
        If next_div >= 21 And next_div <= 26 Then
            If i_ctr = 0 Then
                Printer.ScaleMode = vbCentimeters
                Printer.FontName = "Comic Sans MS"
                Printer.FontBold = True
                Printer.FontSize = 14
                myTextWidth = Printer.TextWidth("SOCIAL SCIENCES")
                Printer.CurrentX = ((Printer.ScaleWidth - myTextWidth) / 2) - 0.5
                Printer.CurrentY = 1
                Printer.Print "SOCIAL SCIENCES"
                Printer.CurrentY = Printer.CurrentY + 0.5
                Printer.CurrentX = 0.5
                yPrint = Printer.CurrentY
                i_ctr = i_ctr + 1
            End If
            If i_ctr > 0 Then
                Printer.ScaleMode = vbCentimeters
                Printer.CurrentX = 0.5
                If yPrint >= 25 Then
                    Printer.NewPage
                    yPrint = Printer.CurrentY + 0.5
                Else
                    Printer.CurrentY = yPrint
                End If
                Printer.FontName = "Comic Sans MS"
                Printer.FontBold = True
                Printer.FontSize = 11
                Printer.Print Mid(ctr_let, i_ctr, 1) & ". " & globalDiv
                Printer.CurrentX = 0.5
                yPrint = Printer.CurrentY
                i_ctr = i_ctr + 1
            End If
            Printer.FontName = "Times New Roman"
            Printer.FontBold = False
            Printer.FontSize = 11
            Printer.Print "Title"
            Printer.CurrentY = yPrint
            Printer.CurrentX = 8
            Printer.Print "Author"
            Printer.CurrentY = yPrint
            Printer.CurrentX = 12
            Printer.Print "Call Number"
            Printer.CurrentY = yPrint
            Printer.CurrentX = 16.5
            Printer.Print "CD-ROM #"
            yPrint = Printer.CurrentY + 0.25
        End If
'if division is not PURE SCIENCES or not SOCIAL SCIENCES
        If Not (next_div >= 14 And next_div <= 19) And Not (next_div >= 21 And next_div <= 26) Then
            Printer.ScaleMode = vbCentimeters
            Printer.FontName = "Comic Sans MS"
            Printer.FontBold = True
            Printer.FontSize = 14
            myTextWidth = Printer.TextWidth(globalDiv)
            Printer.CurrentX = ((Printer.ScaleWidth - myTextWidth) / 2) - 0.5
            Printer.CurrentY = 1
            Printer.Print globalDiv
            Printer.CurrentY = Printer.CurrentY + 0.5
            Printer.CurrentX = 0.5
            Printer.FontName = "Times New Roman"
            Printer.FontBold = False
            Printer.FontSize = 11
            Printer.Print "Title"
            Printer.CurrentY = 2.5
            Printer.CurrentX = 8
            Printer.Print "Author"
            Printer.CurrentY = 2.5
            Printer.CurrentX = 12
            Printer.Print "Call Number"
            Printer.CurrentY = 2.5
            Printer.CurrentX = 16.5
            Printer.Print "CD-ROM #"
            yPrint = Printer.CurrentY + 0.5
        End If
'printing now of the records
        Do While Not oRsPrint.EOF
'for title
            i = 1
            sTemp = ""
            strlen = Len(oRsPrint("fldTitle"))
            Do While Printer.TextWidth(sTemp) + 0.5 < 7.75 And Len(sTemp) < strlen
                sTemp = sTemp & Mid(oRsPrint("fldTitle"), i, 1)
                i = i + 1
            Loop
            Printer.CurrentY = yPrint
            Printer.CurrentX = 0.5
            Printer.Print sTemp
'for author
            i = 1
            sTemp = ""
            strlen = Len(oRsPrint("fldAuthor"))
            Do While Printer.TextWidth(sTemp) + 8 + Printer.CurrentX < 11.75 And Len(sTemp) < strlen
                sTemp = sTemp & Mid(oRsPrint("fldAuthor"), i, 1)
                i = i + 1
            Loop
            Printer.CurrentY = yPrint
            Printer.CurrentX = 8
            Printer.Print sTemp
'for call number
            i = 1
            sTemp = ""
            sHold = ""
            If oRsPrint("fldQty") > 1 Then
                sHold = oRsPrint("fldCallnum") & " [" & Str(oRsPrint("fldQty")) & " copies]"
            Else
                sHold = oRsPrint("fldCallnum")
            End If
            strlen = Len(sHold)
            Do While Printer.TextWidth(sTemp) + 12 < 16.25 And Len(sTemp) < strlen
                sTemp = sTemp & Mid(sHold, i, 1)
                i = i + 1
            Loop
            Printer.CurrentY = yPrint
            Printer.CurrentX = 12
            Printer.Print sTemp
'for CD-ROM #
            i = 1
            sTemp = ""
            sHold = ""
            If oRsPrint("fldCDROMnum") = 0 Then
                sHold = "PC" & oRsPrint("fldPCnum")
            Else
                sHold = oRsPrint("fldCDROMnum") & " PC" & oRsPrint("fldPCnum")
            End If
            strlen = Len(sHold)
            Do While Printer.TextWidth(sTemp) + 16.5 < Printer.ScaleWidth And Len(sTemp) < strlen
                sTemp = sTemp & Mid(sHold, i, 1)
                i = i + 1
            Loop
            Printer.CurrentY = yPrint
            Printer.CurrentX = 16.5
            Printer.Print sTemp
            yPrint = Printer.CurrentY + 0.25
            oRsPrint.MoveNext
            If yPrint >= 25 Then
                Printer.NewPage
                yPrint = Printer.CurrentY + 0.5
            End If
        Loop
        If Not (next_div >= 14 And next_div <= 19) And Not (next_div >= 21 And next_div <= 26) Then
            If yPrint < 25 Then
                Printer.NewPage
            End If
        End If
err_next_div:
        next_div = next_div + 1
        If next_div = 20 Then
            If yPrint < 26 Then
                Printer.NewPage
            End If
            i_ctr = 0
        End If
    Loop
    Printer.EndDoc
    MsgBox "Finished spooling to printer", vbExclamation, "CD-ROM Collection V1.0"
    Exit Function
    
err_print:
    MsgBox "Sorry, printer is not available at this time", vbCritical, "CD-ROM Collection V1.0"
    oRsPrint.Close
    Set oRsPrint = Nothing
    Set oRsPrint = New ADODB.Recordset
    oRsPrint.Open "SELECT * FROM tblCDROM", oCn, adOpenKeyset, adLockOptimistic, adCmdText
    oRsPrint.MoveFirst
End Function

Function get_div_name(div_num As Integer) As String
    Select Case div_num
        Case 1: get_div_name = "Accounting": globalDiv = "ACCOUNTING"
        Case 2: get_div_name = "Anthropology": globalDiv = "ANTHROPOLOGY"
        Case 3: get_div_name = "Architecture": globalDiv = "ARCHITECTURE"
        Case 4: get_div_name = "Arts and Mass Comm": globalDiv = "ARTS & MASS COMMUNICATION"
        Case 5: get_div_name = "ADB": globalDiv = "ASIAN DEVELOPMENT BANK"
        Case 6: get_div_name = "Computer Science and I.T.": globalDiv = "COMPUTER SCIENCE, INFORMATION TECHNOLOGY & TECHNOLOGY"
        Case 7: get_div_name = "Engineering": globalDiv = "ENGINEERING"
        Case 8: get_div_name = "History": globalDiv = "HISTORY"
        Case 9: get_div_name = "Logic and Philosophy": globalDiv = "LOGIC & PHILOSOPHY"
        Case 10: get_div_name = "Management and Business": globalDiv = "MANAGEMENT & BUSINESS"
        Case 11: get_div_name = "Music": globalDiv = "MUSIC"
        Case 12: get_div_name = "Nursing": globalDiv = "NURSING"
        Case 13: get_div_name = "Psychology": globalDiv = "PSYCHOLOGY/EDUCATIONAL PSYCHOLOGY"
        Case 14: get_div_name = "Astronomy": globalDiv = "Astronomy"
        Case 15: get_div_name = "Biology": globalDiv = "Biology"
        Case 16: get_div_name = "Chemistry": globalDiv = "Chemistry"
        Case 17: get_div_name = "Earth Sciences": globalDiv = "Earth Sciences"
        Case 18: get_div_name = "Mathematics": globalDiv = "Mathematics"
        Case 19: get_div_name = "Physics": globalDiv = "Physics"
        Case 20: get_div_name = "Religion": globalDiv = "RELIGION"
        Case 21: get_div_name = "Economics": globalDiv = "Economics"
        Case 22: get_div_name = "Education": globalDiv = "Education"
        Case 23: get_div_name = "Language and Literature": globalDiv = "Language & Literature"
        Case 24: get_div_name = "Law": globalDiv = "Law"
        Case 25: get_div_name = "Sociology": globalDiv = "Sociology"
        Case 26: get_div_name = "Statistics": globalDiv = "Statistics"
    End Select
End Function
