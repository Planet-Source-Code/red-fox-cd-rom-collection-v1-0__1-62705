Attribute VB_Name = "Module1"
Option Explicit

Function load_main()
    With frmMain
        .Height = 8340
        .Width = 10995
        .cmdNew.Enabled = True
        .cmdAdd.Enabled = False
        .cmdCancelEdit.Enabled = True
        .cmdDelete.Enabled = True
        .cmdTop.Enabled = False
        .cmdPrev.Enabled = False
        .cmdNext.Enabled = True
        .cmdBottom.Enabled = True
        .cmdCancelEdit.Caption = "EDIT"
    End With
    With frmMain.frameMain
        .Visible = True
        .Height = 7095
        .Left = 720
        .Top = 240
        .Width = 9495
    End With
    With frmMain.frameView
        .Visible = False
        .Height = 7095
        .Left = 720
        .Top = 240
        .Width = 9495
    End With
    With frmMain.Combo1
        .AddItem ("Accounting")
        .AddItem ("Anthropology")
        .AddItem ("Architecture")
        .AddItem ("Arts and Mass Comm")
        .AddItem ("ADB")
        .AddItem ("Computer Science and I.T.")
        .AddItem ("Engineering")
        .AddItem ("History")
        .AddItem ("Logic and Philosophy")
        .AddItem ("Management and Business")
        .AddItem ("Music")
        .AddItem ("Nursing")
        .AddItem ("Psychology")
        .AddItem ("Astronomy")
        .AddItem ("Biology")
        .AddItem ("Chemistry")
        .AddItem ("Earth Sciences")
        .AddItem ("Mathematics")
        .AddItem ("Physics")
        .AddItem ("Religion")
        .AddItem ("Economics")
        .AddItem ("Education")
        .AddItem ("Language and Literature")
        .AddItem ("Law")
        .AddItem ("Sociology")
        .AddItem ("Statistics")
        .Text = "Accounting"
    End With
    With frmMain.Combo4
        .AddItem ("Division")
        .AddItem ("Call Number")
        .AddItem ("Author")
        .AddItem ("Title")
        .Text = "Division"
    End With
    With frmMain.Combo2
        .AddItem ("Faculty")
        .AddItem ("Student")
        .Text = "Student"
    End With
    With frmMain.Combo3
        .AddItem ("Accounting")
        .AddItem ("Anthropology")
        .AddItem ("Architecture")
        .AddItem ("Arts and Mass Comm")
        .AddItem ("ADB")
        .AddItem ("Computer Science and I.T.")
        .AddItem ("Engineering")
        .AddItem ("History")
        .AddItem ("Logic and Philosophy")
        .AddItem ("Management and Business")
        .AddItem ("Music")
        .AddItem ("Nursing")
        .AddItem ("Psychology")
        .AddItem ("Astronomy")
        .AddItem ("Biology")
        .AddItem ("Chemistry")
        .AddItem ("Earth Sciences")
        .AddItem ("Mathematics")
        .AddItem ("Physics")
        .AddItem ("Religion")
        .AddItem ("Economics")
        .AddItem ("Education")
        .AddItem ("Language and Literature")
        .AddItem ("Law")
        .AddItem ("Sociology")
        .AddItem ("Statistics")
        .Text = "Accounting"
    End With
End Function

