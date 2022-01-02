VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "brute065 v2.0"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   Icon            =   "mdcrack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame9 
      Caption         =   "Resume File"
      Height          =   2115
      Left            =   0
      TabIndex        =   40
      Top             =   3000
      Width           =   4215
      Begin VB.Frame Frame10 
         Height          =   1290
         Left            =   75
         TabIndex        =   43
         Top             =   750
         Visible         =   0   'False
         Width           =   4065
         Begin VB.TextBox txtupto 
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   1725
            TabIndex        =   46
            Top             =   825
            Width           =   2190
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Info"
            Height          =   495
            Left            =   3075
            TabIndex        =   45
            Top             =   225
            Width           =   840
         End
         Begin VB.CommandButton cmdcreate 
            Caption         =   "Create a Resume file (not done yet)"
            Height          =   495
            Left            =   75
            TabIndex        =   44
            Top             =   225
            Width           =   2925
         End
         Begin VB.Label Label3 
            Caption         =   "Text to Resume From"
            Height          =   315
            Left            =   75
            TabIndex        =   47
            Top             =   900
            Width           =   1665
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete resume file"
         Height          =   390
         Left            =   2175
         TabIndex        =   42
         Top             =   300
         Width           =   1965
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Resume previous session"
         Height          =   390
         Left            =   75
         TabIndex        =   41
         Top             =   300
         Width           =   1965
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Brute065 by hanicraft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   0
      TabIndex        =   23
      Top             =   6240
      Width           =   4815
      Begin VB.Label Label4 
         Caption         =   "To use resume files place everthing in c:\tmp       "
         Height          =   420
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   3690
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Hash"
      Height          =   2925
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   4890
      Begin VB.OptionButton optcmd 
         Caption         =   "Dont Start a new console"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Width           =   3495
      End
      Begin VB.OptionButton optcmd 
         Caption         =   "Keep Console window open (command.com 95/98)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   3975
      End
      Begin VB.OptionButton optcmd 
         Caption         =   "Keep Console window open (cmd.exe 2K/XP)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Value           =   -1  'True
         Width           =   3855
      End
      Begin VB.CommandButton cmdexe 
         Caption         =   "BruteForce"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   480
         TabIndex        =   24
         Top             =   2250
         Width           =   3990
      End
      Begin VB.TextBox txtpass 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         TabIndex        =   17
         Text            =   "202cb962ac59075b964b07152d234b70"
         Top             =   225
         Width           =   4740
      End
      Begin VB.Label lblwarn 
         Caption         =   "WARNING if a computed hash is found then the window will close and you will not see the result"
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   480
         TabIndex        =   38
         Top             =   1800
         Visible         =   0   'False
         Width           =   3975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      Height          =   7260
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   4890
      Begin VB.CheckBox Check1 
         Caption         =   "Check all collisions"
         Height          =   390
         Left            =   240
         TabIndex        =   39
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Frame Frame7 
         Caption         =   "If you already know part of the password (leave blank for none)"
         Height          =   690
         Left            =   150
         TabIndex        =   18
         Top             =   6375
         Width           =   4665
         Begin VB.TextBox txte 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2850
            TabIndex        =   20
            Top             =   300
            Width           =   1290
         End
         Begin VB.TextBox txtb 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   975
            TabIndex        =   19
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   315
            Left            =   2475
            TabIndex        =   22
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Beginning"
            Height          =   390
            Left            =   150
            TabIndex        =   21
            Top             =   300
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Verbose Options (enabled=slower)"
         Height          =   1245
         Left            =   2175
         TabIndex        =   11
         Top             =   4875
         Width           =   2640
         Begin VB.OptionButton optv 
            Caption         =   "None (disabled)"
            ForeColor       =   &H00C00000&
            Height          =   390
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   180
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optv 
            Caption         =   "Verbose Extra MD5 (enabled)"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   2415
         End
         Begin VB.OptionButton optv 
            Caption         =   "Verbose (enabled)"
            Height          =   390
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.TextBox txtsize 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         TabIndex        =   10
         Top             =   4500
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Hash Algorithm"
         Height          =   1245
         Left            =   150
         TabIndex        =   6
         Top             =   4875
         Width           =   1890
         Begin VB.OptionButton opth 
            Caption         =   "MD4 Hash"
            Height          =   390
            Index           =   0
            Left            =   150
            TabIndex        =   9
            Top             =   525
            Width           =   1365
         End
         Begin VB.OptionButton opth 
            Caption         =   "MD5 hash"
            ForeColor       =   &H00C00000&
            Height          =   390
            Index           =   1
            Left            =   150
            TabIndex        =   8
            Top             =   225
            Value           =   -1  'True
            Width           =   1365
         End
         Begin VB.OptionButton opth 
            Caption         =   "NTLM1 Hash"
            Height          =   390
            Index           =   2
            Left            =   150
            TabIndex        =   7
            Top             =   825
            Width           =   1365
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1365
         Left            =   180
         TabIndex        =   4
         Top             =   2820
         Width           =   4665
         Begin VB.CheckBox chkfastout 
            Caption         =   "Fast Write to file (not human readable)"
            Height          =   252
            Left            =   150
            TabIndex        =   15
            Top             =   975
            Width           =   3240
         End
         Begin VB.TextBox txtout 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   150
            TabIndex        =   5
            Top             =   600
            Width           =   3540
         End
         Begin VB.Label Label5 
            Caption         =   "Output computed Hash to file (leave blank for none)"
            Height          =   390
            Left            =   180
            TabIndex        =   26
            Top             =   300
            Width           =   4440
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Custom Charset"
         Height          =   2490
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   4665
         Begin VB.Frame fmChar 
            Height          =   1515
            Left            =   150
            TabIndex        =   27
            Top             =   900
            Width           =   4440
            Begin VB.CommandButton Command5 
               Caption         =   "Clear"
               Enabled         =   0   'False
               Height          =   315
               Left            =   75
               TabIndex        =   33
               Top             =   525
               Width           =   615
            End
            Begin VB.CommandButton cmdother 
               Caption         =   "!-)"
               Enabled         =   0   'False
               Height          =   240
               Left            =   3225
               TabIndex        =   32
               Top             =   225
               Width           =   765
            End
            Begin VB.CommandButton cmdup 
               Caption         =   "A-Z"
               Enabled         =   0   'False
               Height          =   240
               Left            =   2400
               TabIndex        =   31
               Top             =   225
               Width           =   765
            End
            Begin VB.CommandButton cmdnum 
               Caption         =   "0-9"
               Enabled         =   0   'False
               Height          =   240
               Left            =   1575
               TabIndex        =   30
               Top             =   225
               Width           =   765
            End
            Begin VB.CommandButton cmdlow 
               Caption         =   "a-z"
               Enabled         =   0   'False
               Height          =   240
               Left            =   750
               TabIndex        =   29
               Top             =   225
               Width           =   765
            End
            Begin VB.TextBox txtchars 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   915
               Left            =   750
               MultiLine       =   -1  'True
               TabIndex        =   28
               Top             =   525
               Width           =   3540
            End
         End
         Begin VB.OptionButton opt 
            Caption         =   "Default (a-zA-Z0-9"
            ForeColor       =   &H00C00000&
            Height          =   390
            Index           =   0
            Left            =   150
            TabIndex        =   3
            Top             =   225
            Value           =   -1  'True
            Width           =   1740
         End
         Begin VB.OptionButton opt 
            Caption         =   "Custom"
            Height          =   390
            HelpContextID   =   3
            Index           =   1
            Left            =   150
            TabIndex        =   2
            Top             =   525
            Width           =   1740
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Minimum Password Size (blank for default)"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   4560
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim apppath, hash, siz, outhash, verbose, findall, cmd

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        findall = " -a"
    Else
        findall = ""
    End If
End Sub

Private Sub chkfastout_Click()
    If txtout.Text <> "" And chkfastout.Value = 1 Then
        fastout = " -F"
    Else
        fastout = ""
    End If
End Sub

Private Sub cmdcreate_Click()
'this doesnt work yet, probably because mdcrack.exe sets file permissions

'optdata = " 0 0 none 0 0 0 none none 1"
'Open apppath & ".mdcrack.resume" For Output As #1
 '   Print #1, txtpass.Text & " " & txtupto.Text & " 0 1 " & txtchars.Text & optdata
'Close #1
End Sub

Private Sub cmdexe_Click()
    'determine whether it is a custom charset
    If txtchars.Text = "" Then
        chars = ""
    Else
        chars = " -s " & txtchars.Text
    End If
    
    If txtb.Text = "" Then
        bn = ""
    Else
        bn = " -b " & txtb.Text
    End If
    If txte.Text = "" Then
        en = ""
    Else
        en = " -e " & txte.Text
    End If
    
    'execute command
    exec = Shell(cmd & apppath & "mdcrack.exe" & hash & verbose & siz & outhash & fastout & findall & en & bn & chars & " " & txtpass.Text, vbNormalFocus)
End Sub

Private Sub cmdlow_Click()
    txtchars.Text = txtchars.Text & "abcdefghijklmnopqrstuvwxyz"
End Sub

Private Sub cmdnum_Click()
    txtchars.Text = txtchars.Text & "0123456789"
End Sub

Private Sub cmdother_Click()
    txtchars.Text = txtchars.Text & "~!@#$%^&*()"
End Sub

Private Sub cmdup_Click()
    txtchars.Text = txtchars.Text & "ABCDEFGHIJKLMNOPQRSTUWXYZ"
End Sub

Private Sub Command1_Click()
    MsgBox "* Overrides any previous resume file" & vbNewLine & _
    "* Resumes from 'Text to Resume From' Field" & vbNewLine & _
    "* Uses MD5 hash from 'Hash' field" & vbNewLine & _
    "* Uses Charset from 'Custom Charset' Field" & vbNewLine & _
    "* Currently experimental so those are all the options you get" & vbNewLine, vbInformation, "Resume File Generator"
End Sub

Private Sub Command3_Click()
    sure = MsgBox("Are you sure", vbCritical + vbYesNo)
    If sure = vbYes Then
        exec = Shell(cmd & apppath & "mdcrack.exe -d", vbNormalFocus)
    End If
End Sub

Private Sub Command4_Click()
    exec = Shell(cmd & apppath & "mdcrack.exe", vbNormalFocus)
End Sub

Private Sub Command5_Click()
    txtchars.Text = ""
End Sub

Private Sub Form_Load()
'initalise variables
    apppath = App.Path & "\"
    hash = " -M MD5"
    cmd = "cmd /K "
End Sub


Private Sub opt_Click(Index As Integer)
'preset charsets or custom or none
    opt(0).ForeColor = vbBlack
    opt(1).ForeColor = vbBlack
    Select Case Index
    Case 0
        txtchars.Text = ""
        txtchars.Enabled = False
        cmdlow.Enabled = False
        cmdup.Enabled = False
        cmdnum.Enabled = False
        cmdother.Enabled = False
        Command5.Enabled = False
    Case 1
        txtchars.Enabled = True
        cmdlow.Enabled = True
        cmdup.Enabled = True
        cmdnum.Enabled = True
        cmdother.Enabled = True
        Command5.Enabled = True
    End Select
    opt(Index).ForeColor = RGB(0, 0, 200)
End Sub

Private Sub optcmd_Click(Index As Integer)
    'type of command interpter
    'if none is used then mdcrack will just close the console afterwards
    Select Case Index
    Case 0:
        cmd = "cmd /K " 'windows 2000/XP
        lblwarn.Visible = False
    Case 1:
        cmd = "command.com /K " 'windows 95/98
        lblwarn.Visible = False
    Case 2:
        cmd = ""
        lblwarn.Visible = True
    End Select
End Sub

Private Sub opth_Click(Index As Integer)
    'hash type
    opth(0).ForeColor = vbBlack
    opth(1).ForeColor = vbBlack
    opth(2).ForeColor = vbBlack
    Select Case Index
    Case 0
        hash = " -M MD5"
    Case 1
        hash = " -M MD4"
    Case 2
        hash = " -M NTLM1"
    End Select
    opth(Index).ForeColor = RGB(0, 0, 200)
End Sub

Private Sub optv_Click(Index As Integer)
    'verbose mode
    optv(0).ForeColor = vbBlack
    optv(1).ForeColor = vbBlack
    optv(2).ForeColor = vbBlack
    Select Case Index
    Case 0
        verbose = ""
    Case 1
        verbose = " -v"
    Case 2
        verbose = " -V"
    End Select
    optv(Index).ForeColor = RGB(0, 0, 200)
End Sub

Private Sub txtout_Change()
    'fastwrite
    If txtout.Text = "" Then
        outhash = ""
    Else
        outhash = " -W " & txtout.Text
    End If
End Sub

Private Sub txtsize_Change()
    If txtsize <> "" Then
        siz = " -S " & txtsize.Text
        'simple message notifies user of long password
        If Val(txtsize.Text) > 8 Then MsgBox "Passwords longer than 8 take a very long time", vbInformation
    Else
        siz = ""
    End If
End Sub
