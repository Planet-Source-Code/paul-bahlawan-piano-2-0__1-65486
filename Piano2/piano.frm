VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmPiano 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Armin Piano"
   ClientHeight    =   2235
   ClientLeft      =   3195
   ClientTop       =   4335
   ClientWidth     =   10290
   Icon            =   "piano.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   10290
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   70
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   68
      Left            =   9570
      Style           =   1  'Graphical
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   66
      Left            =   9330
      Style           =   1  'Graphical
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   71
      Left            =   9930
      Style           =   1  'Graphical
      TabIndex        =   79
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   69
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   67
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   77
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   65
      Left            =   9210
      Style           =   1  'Graphical
      TabIndex        =   76
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.Timer tmrWelcome 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   1080
   End
   Begin VB.CommandButton cmdRec 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Record"
      Height          =   270
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   375
      Width           =   675
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Play"
      Height          =   270
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   120
      Width           =   675
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Save"
      Height          =   270
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   375
      Width           =   675
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Load"
      Height          =   270
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   120
      Width           =   675
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -30
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Armin Piano"
      Filter          =   "*.apo|*.apo"
   End
   Begin VB.Timer tmrRec 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   840
      Top             =   1080
   End
   Begin VB.Timer tmrPlayBack 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   1080
   End
   Begin MSComctlLib.Slider sldVol 
      Height          =   300
      Left            =   5430
      TabIndex        =   65
      Top             =   345
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      LargeChange     =   25
      Max             =   127
      SelStart        =   127
      TickStyle       =   3
      TickFrequency   =   10
      Value           =   127
      TextPosition    =   1
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   61
      Left            =   8625
      Style           =   1  'Graphical
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   63
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   49
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   51
      Left            =   7185
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   54
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   56
      Left            =   7905
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   58
      Left            =   8145
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   46
      Left            =   6465
      Style           =   1  'Graphical
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   44
      Left            =   6225
      Style           =   1  'Graphical
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   42
      Left            =   5985
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   39
      Left            =   5505
      Style           =   1  'Graphical
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   37
      Left            =   5265
      Style           =   1  'Graphical
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   34
      Left            =   4785
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   32
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   30
      Left            =   4305
      Style           =   1  'Graphical
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   27
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   25
      Left            =   3585
      Style           =   1  'Graphical
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   22
      Left            =   3105
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   20
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   18
      Left            =   2625
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   15
      Left            =   2145
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   13
      Left            =   1905
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   10
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   8
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   6
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   3
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00000000&
      Height          =   855
      Index           =   1
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   64
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   62
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   60
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   59
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   57
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   55
      Left            =   7770
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   53
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   52
      Left            =   7290
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   50
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   48
      Left            =   6810
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   47
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   45
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   43
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   41
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   40
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   38
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   36
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   35
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   33
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   31
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   29
      Left            =   4170
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   28
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   26
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   24
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   23
      Left            =   3210
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   21
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   19
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   17
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   16
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   14
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   12
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   11
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   9
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   7
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   5
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   4
      Left            =   570
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   2
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton pKey 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   750
      Width           =   255
   End
   Begin MSComctlLib.Slider sldPitch 
      Height          =   300
      Left            =   6320
      TabIndex        =   67
      Top             =   345
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      LargeChange     =   12
      Min             =   -12
      Max             =   12
      TickStyle       =   3
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider sldInst 
      Height          =   300
      Left            =   7200
      TabIndex        =   68
      Top             =   345
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   529
      _Version        =   393216
      LargeChange     =   1
      Max             =   128
      TickStyle       =   3
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider sldKeyboard 
      Height          =   300
      Left            =   4550
      TabIndex        =   84
      Top             =   360
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   4
      SelStart        =   2
      TickStyle       =   3
      TickFrequency   =   10
      Value           =   2
      TextPosition    =   1
   End
   Begin VB.Shape shpMiddleC 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1800
      Top             =   720
      Width           =   120
   End
   Begin VB.Label lblPed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ped"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   195
      TabIndex        =   86
      Top             =   520
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard"
      Height          =   225
      Index           =   1
      Left            =   4550
      TabIndex        =   85
      Top             =   120
      Width           =   825
   End
   Begin VB.Label lblInstrument 
      Alignment       =   2  'Center
      Caption         =   "Instrument"
      Height          =   255
      Left            =   7320
      TabIndex        =   83
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Piano"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Index           =   1
      Left            =   1560
      TabIndex        =   73
      Top             =   0
      Width           =   1380
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Armin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   72
      Top             =   0
      Width           =   1410
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pitch"
      Height          =   225
      Index           =   2
      Left            =   6320
      TabIndex        =   69
      Top             =   120
      Width           =   825
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      Height          =   225
      Index           =   0
      Left            =   5430
      TabIndex        =   66
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmPiano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' P I A N O  by Armin Niki
'' Original code:
'' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64928&lngWId=1
''
''======================================================================================
'' Update - Apr 16 2006 by Paul Bahlawan
''  - ADD Instrument selection
''  - ADD record key-up event and use that during playback
''  - Remove "use keyboard" button & make keyboard always active
''  - Adjust the notes played when using keyboard (was 2 4 6 8... should be 1 3 5 6 8...)
''  - Make piano-keys a control array to simplify code
''  - Properly declare all variables & add Option Explicit
''  - Much code clean up & comments too!
''  - Make Play & Record buttons toggle on/off
''  - Make/add an icon
''  - Make Fur Elise demo
''
'' Update - Apr 20 2006 by Paul Bahlawan
''  - ADD display all keys currently being played
''
'' Update - Apr 29 2011 by Paul Bahlawan
''  - Implement new keyboard mapping scheme
''  - ADD Instrument list based on http://www.midi.org/techspecs/gm1sound.php
''  - Add Sustain via Shift key
''  - Remove channel select
''  - Change pitch range to be +/- 1 octave
''
'' Update - May 10 2011 by Paul Bahlawan
''  - Change pitch indicator
''  - Add another keyboard map
''  - Extend Fur Elise, Make Ode to Joy and Twinkle Twinkle

Option Explicit

Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Private hmidi As Long
Private baseNote As Long
Private channel As Long
Private velocity As Long
Private lNote As Long
Private Playin() As String
Private playinc As Long
Private timers As Long
Private rec As String
Private KeyMap(255) As Long

'Toggle record on/off
Private Sub cmdRec_Click()
    If tmrPlayBack.Enabled Then cmdPlay_Click ' Stop playback
    
    If tmrRec.Enabled Then
        'stop recording
        tmrRec.Enabled = False
        cmdRec.BackColor = &HC0C0FF
    Else
        'start recording
        rec = ""
        timers = 0
        tmrRec.Enabled = True
        cmdRec.BackColor = &H4040FF
    End If
End Sub

'Toggle playback on/off
Private Sub cmdPlay_Click()
Dim x As Long

    If tmrRec.Enabled Then cmdRec_Click 'Stop recording
    
    If tmrPlayBack.Enabled Then
        'stop playback
        tmrPlayBack.Enabled = False
        cmdPlay.BackColor = &HC0FFC0
        Sustain False
        For x = 1 To 71 '(stop all notes)
            domusicstop x
        Next x
    Else
        'start playback
        Playin = Split(rec, " ")
        playinc = 0
        tmrPlayBack.Interval = 50
        tmrPlayBack.Enabled = True
        cmdPlay.BackColor = &H10FF10
    End If
End Sub

'Save recorded music to disk
Private Sub cmdSave_Click()
Dim ff As Long

    If tmrRec.Enabled Then cmdRec_Click 'Stop Recording
    If tmrPlayBack.Enabled Then cmdPlay_Click 'Stop playback
    
    If rec = "" Then Exit Sub 'nothing to save
    
    CommonDialog1.ShowSave
    
    If Not CommonDialog1.FileName = "" Then
        ff = FreeFile
        Open CommonDialog1.FileName For Binary Access Write As #ff
        Put #ff, , rec
        Close #ff
    End If
End Sub

'Load music from disk
Private Sub cmdLoad_Click()
Dim ff As Long
Dim temp As Long

    If tmrRec.Enabled Then cmdRec_Click 'Stop recording
    If tmrPlayBack.Enabled Then cmdPlay_Click 'Stop playback
    
    CommonDialog1.ShowOpen
    
    ff = FreeFile
    If Not CommonDialog1.FileName = "" Then
        rec = ""
        Open CommonDialog1.FileName For Input As ff
        rec = Input(LOF(ff), ff)
        Close ff
    End If
End Sub

'Play piano via KEYBOARD
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim note As Long

    note = KeyMap(KeyCode)
    If note > 0 Then
        If Not note = lNote And note Then
            domusic note
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim note As Long
    
    note = KeyMap(KeyCode)
    If note > 0 Then
        domusicstop (note)
    End If
End Sub

'Play piano via MOUSE
Private Sub pKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    domusic Index + 1
End Sub

Private Sub pKey_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    domusicstop Index + 1
End Sub

'Start a note
Private Sub domusic(mNote As Long)
Dim midimsg As Long
    
    If mNote = 88 Then
        Sustain True
    Else
        'Play the note
        midimsg = &H90 + channel + ((baseNote + mNote) * &H100) + (velocity * &H10000)
        midiOutShortMsg hmidi, midimsg
        'Hi-light key being played
        pKey(mNote - 1).BackColor = vbRed
    End If
    
    'Record the key-down event
    If tmrRec.Enabled Then
        rec = rec & mNote & "x" & timers & " "
        timers = 0
    End If
    lNote = mNote
    
End Sub

'Stop a note
Private Sub domusicstop(mNote As Long)
Dim midimsg As Long

    If mNote = 88 Then
        Sustain False
    Else
        'Stop the note
        midimsg = &H80 + ((baseNote + mNote) * &H100) + channel
        midiOutShortMsg hmidi, midimsg
        
        'Un-hi-light released key
        If pKey(mNote - 1).Tag = "1" Then
            pKey(mNote - 1).BackColor = vbWhite
        Else
            pKey(mNote - 1).BackColor = vbBlack
        End If
    End If
    
    'Record the key-up event
    If tmrRec.Enabled Then
        rec = rec & -mNote & "x" & timers & " "
        timers = 0
    End If
    If mNote = lNote Then lNote = 0
End Sub


Private Sub Form_Activate()
Dim rc As Long
Dim curDevice As Long
Dim x As Long

    'Open MIDI device
    midiOutClose hmidi
    rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
    If (rc <> 0) Then
        MsgBox "Couldn't open midi device - Error #" & rc
    End If
    
    CommonDialog1.InitDir = App.Path
    
    'Set initial parameters
    sldKeyboard_Change
    sldVol_Scroll
    sldPitch_Scroll
    sldInst_Scroll
    velocity = 127
    
    tmrWelcome.Enabled = True
End Sub


Private Sub Form_Terminate()
    midiOutClose hmidi
End Sub

Private Sub Form_Unload(Cancel As Integer)
    midiOutClose hmidi
End Sub

Private Sub lblTitle_Click(Index As Integer)
    tmrWelcome.Enabled = True
End Sub

'Change the instrument
Private Sub sldInst_Scroll()
Dim midimsg As Long
Dim sel As Long

    sel = sldInst.Value
    lblInstrument.Caption = LoadResString(sel)
    
    If sel = 128 Then
        'Percussion
        channel = 9
    Else
        'All other instruments
        channel = 0
        midimsg = (sel * &H100) + &HC0 + channel 'Program change
        midiOutShortMsg hmidi, midimsg
    End If
End Sub

'Select Keyboard Mapping
Private Sub sldKeyboard_Change()
Dim temp() As String
Dim x As Long

    For x = 300 To 347
        temp = Split(LoadResString(x), ",")
        KeyMap(CLng(temp(0))) = CLng(temp(sldKeyboard.Value))
    Next x
    KeyMap(16) = 88 'Shift = Sustain
End Sub

'Change the pitch
Private Sub sldPitch_Scroll()
    baseNote = 23 + sldPitch.Value
    
    'Show middle C
    shpMiddleC.Left = pKey(36 - sldPitch).Left + pKey(36 - sldPitch).Width / 2 - 60
End Sub

'Change the Volume
Private Sub sldVol_Scroll()
    velocity = sldVol.Value
End Sub

'Activate/Deactivate Sustain
Private Sub Sustain(Activate As Boolean)
    If Activate Then
        midiOutShortMsg hmidi, (&HB0 + channel + &H4000 + &H7F0000)
        lblPed.Visible = True
    Else
        midiOutShortMsg hmidi, (&HB0 + channel + &H4000)
        lblPed.Visible = False
    End If
End Sub

'Play back a recording
Private Sub tmrPlayBack_Timer()
Dim getnote() As String
Dim temp As Long

    On Error GoTo Errs
    
    getnote = Split(Playin(playinc), "x")
    
    temp = getnote(0)
    If temp < 0 Then
        'key-up
        domusicstop Abs(temp)
    Else
        'key-down
        domusic temp
    End If
    
    playinc = playinc + 1
    getnote = Split(Playin(playinc), "x")
    
    temp = getnote(1) * 50
    If temp = 0 Then 'a 0 means another event happens at the exact same time, so do it now!
        tmrPlayBack_Timer
        Exit Sub
    End If
    
    tmrPlayBack.Enabled = False
    tmrPlayBack.Interval = temp + 50
    tmrPlayBack.Enabled = True
    Exit Sub

Errs:
    cmdPlay_Click 'Stop playback
End Sub

'used during recording
Private Sub tmrRec_Timer()
    timers = timers + 1
End Sub

'Play welcome tune
Private Sub tmrWelcome_Timer()
Static pdemo As Long

    If pdemo > 1 Then domusicstop pdemo - 7
    If pdemo > 64 Then
        pdemo = 0
        tmrWelcome.Enabled = False
        Exit Sub
    End If
    
    domusic pdemo + 5
    pdemo = pdemo + 12
End Sub
