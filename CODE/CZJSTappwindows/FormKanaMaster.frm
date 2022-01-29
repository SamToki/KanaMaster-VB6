VERSION 5.00
Begin VB.Form FormKanaMaster 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KanaMaster"
   ClientHeight    =   11670
   ClientLeft      =   45
   ClientTop       =   750
   ClientWidth     =   18600
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "FormKanaMaster.frx":0000
   LinkTopic       =   "FormKanaMaster"
   MaxButton       =   0   'False
   MouseIcon       =   "FormKanaMaster.frx":23D2
   MousePointer    =   99  'Custom
   ScaleHeight     =   11670
   ScaleWidth      =   18600
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdStop 
      Caption         =   "停止(&O)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4200
      MouseIcon       =   "FormKanaMaster.frx":2524
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   210
      Width           =   1380
   End
   Begin VB.Timer TimerResultAnimation 
      Interval        =   1
      Left            =   13545
      Top             =   1155
   End
   Begin VB.Timer TimerPopupAnimation 
      Interval        =   1
      Left            =   5775
      Top             =   8190
   End
   Begin VB.Timer TimerNumberAnimation 
      Interval        =   50
      Left            =   17955
      Top             =   630
   End
   Begin VB.Timer TimerArrowAnimation 
      Interval        =   1
      Left            =   3570
      Top             =   1470
   End
   Begin VB.Timer TimerToastAnimation 
      Interval        =   1
      Left            =   13860
      Top             =   5145
   End
   Begin VB.Timer TimerSakuraAnimation 
      Interval        =   1
      Left            =   630
      Top             =   1785
   End
   Begin VB.Timer TimerProgressbarAnimation 
      Interval        =   1
      Left            =   13125
      Top             =   1155
   End
   Begin VB.Timer TimerTimer 
      Interval        =   90
      Left            =   16800
      Top             =   5985
   End
   Begin VB.TextBox TextboxInput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00709000&
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   5775
      MaxLength       =   1
      MouseIcon       =   "FormKanaMaster.frx":2676
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "键盘输入框"
      Top             =   210
      Width           =   435
   End
   Begin VB.CommandButton CmdOption3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   12390
      MouseIcon       =   "FormKanaMaster.frx":27C8
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   8820
      Width           =   5370
   End
   Begin VB.CommandButton CmdOption1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   840
      MouseIcon       =   "FormKanaMaster.frx":291A
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   8820
      Width           =   5370
   End
   Begin VB.CommandButton CmdOption2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   6615
      MouseIcon       =   "FormKanaMaster.frx":2A6C
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   8820
      Width           =   5370
   End
   Begin VB.CommandButton CmdBackToHome 
      Cancel          =   -1  'True
      Caption         =   "返回主页(&B)"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      MouseIcon       =   "FormKanaMaster.frx":2BBE
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   210
      Width           =   1380
   End
   Begin VB.CommandButton CmdStartPauseResume 
      Caption         =   "开始(&S)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2835
      MouseIcon       =   "FormKanaMaster.frx":2D10
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   210
      Width           =   1380
   End
   Begin VB.Timer TimerClock 
      Interval        =   500
      Left            =   2415
      Top             =   420
   End
   Begin VB.Label LabelAccuracy 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "---.--%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00009000&
      Height          =   960
      Left            =   14700
      TabIndex        =   28
      Top             =   5145
      Width           =   3165
   End
   Begin VB.Label LabelAverageReactionTime 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "-.---s"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000B000&
      Height          =   960
      Left            =   14700
      MouseIcon       =   "FormKanaMaster.frx":2E62
      MousePointer    =   99  'Custom
      TabIndex        =   26
      ToolTipText     =   "「均速」指平均反应用时。"
      Top             =   3675
      Width           =   2640
   End
   Begin VB.Label LabelTimeLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "-.-"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000D000&
      Height          =   960
      Left            =   14700
      MouseIcon       =   "FormKanaMaster.frx":2FB4
      MousePointer    =   99  'Custom
      TabIndex        =   23
      ToolTipText     =   "指示剩余时间。单位：秒"
      Top             =   2205
      Width           =   1275
   End
   Begin VB.Label LabelCurrentDifficulty 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "-.-"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000D000&
      Height          =   540
      Left            =   15960
      MouseIcon       =   "FormKanaMaster.frx":3106
      MousePointer    =   99  'Custom
      TabIndex        =   24
      ToolTipText     =   "指示当前时限。若启用了「缓慢缩短时限」，此数字可能会变化。单位：秒"
      Top             =   2625
      Width           =   645
   End
   Begin VB.Label LabelComboCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "----x"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000050B0&
      Height          =   960
      Left            =   1470
      MouseIcon       =   "FormKanaMaster.frx":3258
      MousePointer    =   99  'Custom
      TabIndex        =   20
      ToolTipText     =   "当前连击数(Combo)"
      Top             =   5145
      Width           =   2430
   End
   Begin VB.Label LabelBestComboCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "----x"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000050B0&
      Height          =   540
      Left            =   2625
      MouseIcon       =   "FormKanaMaster.frx":33AA
      MousePointer    =   99  'Custom
      TabIndex        =   21
      ToolTipText     =   "最大连击数(Best Combo)"
      Top             =   5985
      Width           =   1275
   End
   Begin VB.Label LabelMissCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "----x"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000070D0&
      Height          =   960
      Left            =   1470
      TabIndex        =   18
      Top             =   3675
      Width           =   2430
   End
   Begin VB.Label LabelHP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "--.-"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000090FF&
      Height          =   960
      Left            =   2205
      TabIndex        =   16
      Top             =   2205
      Width           =   1695
   End
   Begin VB.Label LabelLightIndicatorModAU 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "AU"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   330
      Left            =   8505
      MouseIcon       =   "FormKanaMaster.frx":34FC
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "(Auto)"
      Top             =   210
      Width           =   435
   End
   Begin VB.Label LabelLightIndicatorModAP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "AP"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   330
      Left            =   7980
      MouseIcon       =   "FormKanaMaster.frx":364E
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "(Autopilot)"
      Top             =   210
      Width           =   435
   End
   Begin VB.Label LabelLightIndicatorModNF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "NF"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   330
      Left            =   7455
      MouseIcon       =   "FormKanaMaster.frx":37A0
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "(No-Fail)"
      Top             =   210
      Width           =   435
   End
   Begin VB.Label LabelLightIndicatorModPF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "PF"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   330
      Left            =   6930
      MouseIcon       =   "FormKanaMaster.frx":38F2
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "(Perfect)"
      Top             =   210
      Width           =   435
   End
   Begin VB.Label LabelLightIndicatorModSD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "SD"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   330
      Left            =   6405
      MouseIcon       =   "FormKanaMaster.frx":3A44
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "(Sudden Death)"
      Top             =   210
      Width           =   435
   End
   Begin VB.Label LabelReadyArrow3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1485
      Left            =   14490
      TabIndex        =   39
      Top             =   7350
      Width           =   1170
   End
   Begin VB.Label LabelReadyArrow2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1485
      Left            =   8715
      TabIndex        =   38
      Top             =   7350
      Width           =   1170
   End
   Begin VB.Label LabelReadyArrow1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1485
      Left            =   2940
      TabIndex        =   37
      Top             =   7350
      Width           =   1170
   End
   Begin VB.Label LabelPopup1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "-------"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   840
      TabIndex        =   41
      Top             =   7665
      Width           =   5370
   End
   Begin VB.Label LabelToast 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Toast提示信息"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   4305
      TabIndex        =   40
      Top             =   4620
      Width           =   9990
   End
   Begin VB.Label LabelAccuracyTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "准确度"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   540
      Left            =   14490
      MouseIcon       =   "FormKanaMaster.frx":3B96
      MousePointer    =   99  'Custom
      TabIndex        =   27
      ToolTipText     =   "(Accuracy)"
      Top             =   4935
      Width           =   1380
   End
   Begin VB.Label LabelPopup2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "+---"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   16695
      TabIndex        =   14
      Top             =   1050
      Width           =   1695
   End
   Begin VB.Label LabelStartArrow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "↑"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1485
      Left            =   2950
      TabIndex        =   36
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label LabelMissCountTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "失误"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   540
      Left            =   3150
      MouseIcon       =   "FormKanaMaster.frx":3CE8
      MousePointer    =   99  'Custom
      TabIndex        =   17
      ToolTipText     =   "(Miss)"
      Top             =   3465
      Width           =   960
   End
   Begin VB.Label LabelHPTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "血量"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   540
      Left            =   3150
      MouseIcon       =   "FormKanaMaster.frx":3E3A
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "(HP)"
      Top             =   1995
      Width           =   960
   End
   Begin VB.Label LabelTimeElapsed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "--' --'' -"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D07000&
      Height          =   330
      Left            =   11235
      MouseIcon       =   "FormKanaMaster.frx":3F8C
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "计时器（可能不准确）"
      Top             =   210
      Width           =   1590
   End
   Begin VB.Line LineSakura5 
      BorderColor     =   &H00FF90FF&
      BorderWidth     =   5
      X1              =   735
      X2              =   735
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSakura4 
      BorderColor     =   &H00FF90FF&
      BorderWidth     =   5
      X1              =   630
      X2              =   630
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSakura3 
      BorderColor     =   &H00FF90FF&
      BorderWidth     =   5
      X1              =   525
      X2              =   525
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSakura2 
      BorderColor     =   &H00FF90FF&
      BorderWidth     =   5
      X1              =   420
      X2              =   420
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Line LineSakura1 
      BorderColor     =   &H00FF90FF&
      BorderWidth     =   5
      X1              =   315
      X2              =   315
      Y1              =   1575
      Y2              =   1890
   End
   Begin VB.Label LabelClock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1680
      MouseIcon       =   "FormKanaMaster.frx":40DE
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "时钟"
      Top             =   270
      Width           =   1065
   End
   Begin VB.Label LabelScore 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "--------"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   14385
      MouseIcon       =   "FormKanaMaster.frx":4230
      MousePointer    =   99  'Custom
      TabIndex        =   13
      ToolTipText     =   "得分(Score)"
      Top             =   105
      Width           =   4005
   End
   Begin VB.Label LabelAverageReactionTimeTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "均速"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   540
      Left            =   14490
      MouseIcon       =   "FormKanaMaster.frx":4382
      MousePointer    =   99  'Custom
      TabIndex        =   25
      ToolTipText     =   "(Average Reaction Time)"
      Top             =   3465
      Width           =   960
   End
   Begin VB.Label LabelTotalCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "----x"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D07000&
      Height          =   330
      Left            =   12915
      MouseIcon       =   "FormKanaMaster.frx":44D4
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "计数器"
      Top             =   210
      Width           =   1065
   End
   Begin VB.Label LabelProgress 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "---.--%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF9000&
      Height          =   435
      Left            =   11760
      MouseIcon       =   "FormKanaMaster.frx":4626
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "游戏进度"
      Top             =   525
      Width           =   2220
   End
   Begin VB.Shape ShapeTimeLeftProgressbar 
      BackColor       =   &H0000D000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   6000
      Left            =   13965
      Top             =   2205
      Width           =   120
   End
   Begin VB.Shape ShapeHPProgressbar 
      BackColor       =   &H000090FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   6000
      Left            =   4515
      Top             =   2205
      Width           =   120
   End
   Begin VB.Label LabelTimeLeftTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "时限"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   540
      Left            =   14490
      MouseIcon       =   "FormKanaMaster.frx":4778
      MousePointer    =   99  'Custom
      TabIndex        =   22
      ToolTipText     =   "(Time Left)"
      Top             =   1995
      Width           =   960
   End
   Begin VB.Label LabelComboCountTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "连击"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   540
      Left            =   3150
      MouseIcon       =   "FormKanaMaster.frx":48CA
      MousePointer    =   99  'Custom
      TabIndex        =   19
      ToolTipText     =   "(Combo)"
      Top             =   4935
      Width           =   960
   End
   Begin VB.Label LabelOption3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   14490
      MouseIcon       =   "FormKanaMaster.frx":4A1C
      MousePointer    =   99  'Custom
      TabIndex        =   35
      ToolTipText     =   "右边选项的键位"
      Top             =   10815
      Width           =   1170
   End
   Begin VB.Label LabelOption2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   8715
      MouseIcon       =   "FormKanaMaster.frx":4B6E
      MousePointer    =   99  'Custom
      TabIndex        =   34
      ToolTipText     =   "中间选项的键位"
      Top             =   10815
      Width           =   1170
   End
   Begin VB.Shape ShapeTimeLeftBottombar 
      BackColor       =   &H00F0F0F0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   6210
      Left            =   13965
      Top             =   1995
      Width           =   120
   End
   Begin VB.Shape ShapeHPBottombar 
      BackColor       =   &H00F0F0F0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   6210
      Left            =   4515
      Top             =   1995
      Width           =   120
   End
   Begin VB.Label LabelOption1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   2940
      MouseIcon       =   "FormKanaMaster.frx":4CC0
      MousePointer    =   99  'Custom
      TabIndex        =   33
      ToolTipText     =   "左边选项的键位"
      Top             =   10815
      Width           =   1170
   End
   Begin VB.Label LabelDashboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   320.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6600
      Left            =   5260
      TabIndex        =   29
      Top             =   1800
      Width           =   8070
   End
   Begin VB.Shape ShapeLightIndicatorOption1 
      BackColor       =   &H00707070&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   1905
      Left            =   840
      Top             =   8820
      Width           =   5370
   End
   Begin VB.Shape ShapeLightIndicatorOption2 
      BackColor       =   &H00707070&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   1905
      Left            =   6615
      Top             =   8820
      Width           =   5370
   End
   Begin VB.Shape ShapeLightIndicatorOption3 
      BackColor       =   &H00707070&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   1905
      Left            =   12390
      Top             =   8820
      Width           =   5370
   End
   Begin VB.Shape ShapeProgressProgressbar 
      BackColor       =   &H00FF9000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   630
      Top             =   1050
      Width           =   13140
   End
   Begin VB.Shape ShapeProgressBottombar 
      BackColor       =   &H00F0F0F0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   630
      Top             =   1050
      Width           =   13350
   End
   Begin VB.Menu Menu 
      Caption         =   "菜单(&M)"
      Begin VB.Menu MenuStartPauseResume 
         Caption         =   "开始(&S)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuStop 
         Caption         =   "停止(&O)"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu Menu1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuChooseOption1 
         Caption         =   "选择左边选项(&1)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuChooseOption2 
         Caption         =   "选择中间选项(&2)"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MenuChooseOption3 
         Caption         =   "选择右边选项(&3)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Menu2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuBackToHome 
         Caption         =   "返回主页(&B)"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "FormKanaMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

'Declare Game...
Public gamestatus As Integer  '0-Stopped, 3-Ready, 1-Ongoing, 2-Cooldown, 4-Paused.

Public gameprogress As Double  'Unit: %. Range: 0.00~100.00
Public gameclear As Boolean
Public gamequestionrepeatedtimescount As Double
Public gametimeelapsed As Double  'Unit: sec.

Public gamehp As Single
Public gameminimumhp As Single
Public gametimeleft As Single  'Unit: sec.
Public gamecurrentdifficulty As Single  'Unit: sec.
Public gameaveragereactiontime As Single  'Unit: sec.

Public gametotalcount As Long
Public gamecombocount As Long
Public gamebestcombocount As Integer
Public gameperfectcount As Integer
Public gamegreatcount As Integer
Public gamegoodcount As Integer
Public gamemisscount As Integer

Public gamescore As Long  'Range: 0~99,999,999
Public gameaccuracy As Double  'Unit: %. Range: 0.00~100.00
Public gameranking As String

Public lotterytotal As Integer
Public lotterynumber As Integer

Public lotteryquestion As String
Public lotteryquestionlocationX As Integer
Public lotteryquestionlocationY As Integer
Public questiondata As Variant  '(1 To 11, 1 To 16)
Public questionrepeatedtimesdata As Variant  '(1 To 11, 1 To 16)

Public correspondinganswer As String
Public lotteryanswerlocationX As Integer
Public lotteryanswerlocationY As Integer
Public answerdata As Variant  '(1 To 11, 1 To 16)

Public correctoption As Integer
Public chosenoption As Integer

'Declare Display...
Public toastanimationtime As Integer  'Range: 0~1,000.
Public toastanimationtarget As Long  'Range: 0 or 960.
Public popupanimationtarget1 As Long  'Range: 0 or 960.
Public popupanimationtarget2 As Long  'Range: 0 or 750.
Public gameprogressanimationtarget As Long  'Range: 0~13,350.
Public gamehpanimationtarget As Long  'Range: 0~6,210.
Public gametimeleftprogressbaranimationtarget As Long  'Range: 0~6,210.
Public sakuracurrentangle As Single  'Range: -180.000~180.000. Note: 90.000 means straight up.
Public sakuracurrentangle2 As Single
Public sakuracurrentangle3 As Single
Public sakuracurrentangle4 As Single
Public sakuracurrentangle5 As Single
Public sakuracurrentspeed As Single  'Range: 0.00~10.00
Public sakuratargetspeed As Single  'Range: 0.00~10.00. Note: The maximum speed is based on the current difficulty.
Public gamescorenumberanimationcurrent As Long  'Range: 0~99,999,999.
Public arrowanimationtime As Single  'Range: 0~6.28
Public resultanimationtime As Integer  'Range: 0~200.

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Sub Form_Load()
        'Initialize game...
        gamestatus = 0

        gameprogress = 0
        gameclear = False
        gamequestionrepeatedtimescount = 0
        gametimeelapsed = 0

        gamehp = 0
        gameminimumhp = 50
        gametimeleft = 0
        gamecurrentdifficulty = 0
        gameaveragereactiontime = 0

        gametotalcount = 0
        gamecombocount = 0
        gamebestcombocount = 0
        gameperfectcount = 0
        gamegreatcount = 0
        gamegoodcount = 0
        gamemisscount = 0

        gamescore = 0
        gameaccuracy = 0
        gameranking = "??"

        lotterytotal = 0
        lotterynumber = 0

        lotteryquestion = "??"
        questiondata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                             Array("!!", "あ", "か", "さ", "た", "な", "は", "ま", "や", "ら", "わ", "ん", "が", "ざ", "だ", "ば", "ぱ"), _
                             Array("!!", "い", "き", "し", "ち", "に", "ひ", "み", "--", "り", "--", "--", "ぎ", "じ", "ぢ", "び", "ぴ"), _
                             Array("!!", "う", "く", "す", "つ", "ぬ", "ふ", "む", "ゆ", "る", "--", "--", "ぐ", "ず", "づ", "ぶ", "ぷ"), _
                             Array("!!", "え", "け", "せ", "て", "ね", "へ", "め", "--", "れ", "--", "--", "げ", "ぜ", "で", "べ", "ぺ"), _
                             Array("!!", "お", "こ", "そ", "と", "の", "ほ", "も", "よ", "ろ", "を", "--", "ご", "ぞ", "ど", "ぼ", "ぽ"), _
 _
                             Array("!!", "ア", "カ", "サ", "タ", "ナ", "ハ", "マ", "ヤ", "ラ", "ワ", "ン", "ガ", "ザ", "ダ", "バ", "パ"), _
                             Array("!!", "イ", "キ", "シ", "チ", "ニ", "ヒ", "ミ", "--", "リ", "--", "--", "ギ", "ジ", "ヂ", "ビ", "ピ"), _
                             Array("!!", "ウ", "ク", "ス", "ツ", "ヌ", "フ", "ム", "ユ", "ル", "--", "ヴ", "グ", "ズ", "ヅ", "ブ", "プ"), _
                             Array("!!", "エ", "ケ", "セ", "テ", "ネ", "ヘ", "メ", "--", "レ", "--", "--", "ゲ", "ゼ", "デ", "ベ", "ペ"), _
                             Array("!!", "オ", "コ", "ソ", "ト", "ノ", "ホ", "モ", "ヨ", "ロ", "ヲ", "--", "ゴ", "ゾ", "ド", "ボ", "ポ"), _
 _
                             Array("!!", "ゐ", "ゑ", "ヰ", "ヱ", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--") _
                             )
        questionrepeatedtimesdata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
 _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
 _
                                          Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                          )

        correspondinganswer = "??"
        answerdata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                           Array("!!", "a", "ka", "sa", "ta", "na", "ha", "ma", "ya", "ra", "wa", "n", "ga", "za", "da", "ba", "pa"), _
                           Array("!!", "i", "ki", "shi", "chi", "ni", "hi", "mi", "--", "ri", "--", "--", "gi", "ji", "ji", "bi", "pi"), _
                           Array("!!", "u", "ku", "su", "tsu", "nu", "fu", "mu", "yu", "ru", "--", "--", "gu", "zu", "zu", "bu", "pu"), _
                           Array("!!", "e", "ke", "se", "te", "ne", "he", "me", "--", "re", "--", "--", "ge", "ze", "de", "be", "pe"), _
                           Array("!!", "o", "ko", "so", "to", "no", "ho", "mo", "yo", "ro", "wo", "--", "go", "zo", "do", "bo", "po"), _
 _
                           Array("!!", "a", "ka", "sa", "ta", "na", "ha", "ma", "ya", "ra", "wa", "n", "ga", "za", "da", "ba", "pa"), _
                           Array("!!", "i", "ki", "shi", "chi", "ni", "hi", "mi", "--", "ri", "--", "--", "gi", "ji", "ji", "bi", "pi"), _
                           Array("!!", "u", "ku", "su", "tsu", "nu", "fu", "mu", "yu", "ru", "--", "v", "gu", "zu", "zu", "bu", "pu"), _
                           Array("!!", "e", "ke", "se", "te", "ne", "he", "me", "--", "re", "--", "--", "ge", "ze", "de", "be", "pe"), _
                           Array("!!", "o", "ko", "so", "to", "no", "ho", "mo", "yo", "ro", "wo", "--", "go", "zo", "do", "bo", "po"), _
 _
                           Array("!!", "wi", "we", "wi", "we", "wo", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--") _
                           )

        correctoption = 0
        chosenoption = 0

        'Initialize display...
        toastanimationtime = 0
        toastanimationtarget = 0
        popupanimationtarget1 = 0
        popupanimationtarget2 = 0
        gameprogressanimationtarget = 0
        gamehpanimationtarget = 0
        gametimeleftprogressbaranimationtarget = 0
        sakuracurrentangle = 90
        sakuracurrentspeed = 0
        sakuratargetspeed = 0
        gamescorenumberanimationcurrent = 0
        arrowanimationtime = 0
        resultanimationtime = 0
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    'CMD Menu...
    Public Sub MenuBackToHome_Click()
        If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Back.wav"
        gamestatus = 0: Call GameStatusRefresher
        Me.Hide: FormMainWindow.Show
        FormMainWindow.WindowState = 0

        'Hide Result window...
        FormResult.windowanimationtargetleft = (Screen.Width / 2)
        FormResult.windowanimationtargettop = (Screen.Height / 2)
        FormResult.windowanimationtargetwidth = 0
        FormResult.windowanimationtargetheight = 0
    End Sub
    Public Sub CmdBackToHome_Click()
        Call MenuBackToHome_Click
    End Sub

    'CMD Game...
    Public Sub MenuStartPauseResume_Click()
        Select Case gamestatus
            Case 0  'Status: Stopped...
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterReady.wav"
                gamestatus = 3  'Into: Ready...
            Case 3  'Status: Ready...
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterPause.wav"
                gamestatus = 4  'Into: Paused...
            Case 1  'Status: Ongoing...
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterPause.wav"
                gamestatus = 4  'Into: Paused...
            Case 2  'Status: Cooldown...
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterPause.wav"
                gamestatus = 4  'Into: Paused...
            Case 4  'Status: Paused...
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterPause.wav"
                gamestatus = 2  'Into: Cooldown...
            Case Else
                MsgBox "错误：Variable gamestatus is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End Select
        Call GameStatusRefresher: TextboxInput.SetFocus

        Me.WindowState = 0

        'Hide Result window...
        FormResult.windowanimationtargetleft = (Screen.Width / 2)
        FormResult.windowanimationtargettop = (Screen.Height / 2)
        FormResult.windowanimationtargetwidth = 0
        FormResult.windowanimationtargetheight = 0
    End Sub
    Public Sub CmdStartPauseResume_Click()
        Call MenuStartPauseResume_Click
    End Sub
    Public Sub MenuStop_Click()
        If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterStop.wav"
        LabelToast.Caption = "游戏停止": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
        gamestatus = 0: Call GameStatusRefresher: TextboxInput.SetFocus
    End Sub
    Public Sub CmdStop_Click()
        Call MenuStop_Click
    End Sub

    Public Sub MenuChooseOption1_Click()
        TextboxInput.SetFocus
        chosenoption = 1
        Select Case gamestatus
            Case 0
                LabelToast.Caption = "请先开始游戏": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            Case 3
                LabelToast.Caption = "准备中": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            Case 1
                Call GameRespondent
            Case 2
                LabelToast.Caption = "换题还在CD中": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            Case 4
                LabelToast.Caption = "请先继续游戏": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
        End Select
    End Sub
    Public Sub CmdOption1_Click()
        Call MenuChooseOption1_Click
    End Sub
    Public Sub MenuChooseOption2_Click()
        TextboxInput.SetFocus
        chosenoption = 2
        Select Case gamestatus
            Case 0
                LabelToast.Caption = "请先开始游戏": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            Case 3
                LabelToast.Caption = "准备中": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            Case 1
                Call GameRespondent
            Case 2
                LabelToast.Caption = "换题还在CD中": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            Case 4
                LabelToast.Caption = "请先继续游戏": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
        End Select
    End Sub
    Public Sub CmdOption2_Click()
        Call MenuChooseOption2_Click
    End Sub
    Public Sub MenuChooseOption3_Click()
        TextboxInput.SetFocus
        chosenoption = 3
        Select Case gamestatus
            Case 0
                LabelToast.Caption = "请先开始游戏": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            Case 3
                LabelToast.Caption = "准备中": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            Case 1
                Call GameRespondent
            Case 2
                LabelToast.Caption = "换题还在CD中": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            Case 4
                LabelToast.Caption = "请先继续游戏": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
        End Select
    End Sub
    Public Sub CmdOption3_Click()
        Call MenuChooseOption3_Click
    End Sub

    Public Sub TextboxInput_Change()
        Select Case TextboxInput.Text
            Case FormMainWindow.setkeyboardoption(1)
                Call MenuChooseOption1_Click
            Case FormMainWindow.setkeyboardoption(2)
                Call MenuChooseOption2_Click
            Case FormMainWindow.setkeyboardoption(3)
                Call MenuChooseOption3_Click
            Case ""
                Exit Sub
            Case Else
                LabelToast.Caption = "按键错误! 请检查键位": LabelToast.BackColor = &HF0F0F0: LabelToast.ForeColor = &H0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
        End Select

        TextboxInput.Text = ""
    End Sub

'[] TIMERS []

    Public Sub TimerClock_Timer()
        LabelClock.Caption = Format((Hour(Time)), "00") & ":" & Format((Minute(Time)), "00") & ":" & Format((Second(Time)), "00")
    End Sub

    Public Sub TimerTimer_Timer()
        'Game clear judgement...
        If gameclear = True Then GoTo TimerTimer_SkipGameClearJudgement_

        'Progress...
        Select Case FormMainWindow.setprogressmode
            Case 1
                gamequestionrepeatedtimescount = 0
                For forloop1 = 1 To 11
                    For forloop2 = 1 To 16
                        gamequestionrepeatedtimescount = gamequestionrepeatedtimescount + questionrepeatedtimesdata(forloop1)(forloop2) / FormMainWindow.setrepeatedtimes
                    Next
                Next
                If FormMainWindow.settotalquestion <> 0 Then gameprogress = (gamequestionrepeatedtimescount / FormMainWindow.settotalquestion) * 100
            Case 2
                gameprogress = (gametimeelapsed / (FormMainWindow.setspecifiedtime * 60)) * 100
        End Select
        LabelProgress.Caption = Format(gameprogress, "0.00") & "%"
        gameprogressanimationtarget = gameprogress / 100 * 13350
        If gameprogressanimationtarget < 0 Then gameprogressanimationtarget = 0
        If gameprogressanimationtarget > 13350 Then gameprogressanimationtarget = 13350

        'Current Difficulty...
        Select Case FormMainWindow.setincreasedifficultygraduallyswitch
            Case True
                If gameprogress < FormMainWindow.setreachnormaldifficultyat Then
                    gamecurrentdifficulty = FormMainWindow.setinitialdifficulty - (FormMainWindow.setinitialdifficulty - FormMainWindow.setnormaldifficulty) * (gameprogress / FormMainWindow.setreachnormaldifficultyat)
                Else
                    gamecurrentdifficulty = FormMainWindow.setnormaldifficulty
                End If
            Case False
                gamecurrentdifficulty = FormMainWindow.setnormaldifficulty
        End Select
        LabelCurrentDifficulty.Caption = Format(gamecurrentdifficulty, "0.0")

        'Time Elapsed, HP, and Time Left...
        Select Case gamestatus
            Case 3
                'New game initialization...
                gameprogress = 0: gameclear = False: gamequestionrepeatedtimescount = 0: gametimeelapsed = 0: gameaveragereactiontime = 0: gametotalcount = 0: gamecombocount = 0: gamebestcombocount = 0: gameperfectcount = 0: gamegreatcount = 0: gamegoodcount = 0: gamemisscount = 0: gamescore = 0: gameaccuracy = 0
                lotterytotal = 0: lotterynumber = 0: lotteryquestion = "??": correspondinganswer = "??": correctoption = 0: chosenoption = 0
                questionrepeatedtimesdata = Array(Array("!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!", "!!"), _
 _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
 _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
 _
                                                  Array("!!", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                                  )
                gamehp = gamehp + 0.1
                gametimeleft = gametimeleft + 0.1
                LabelDashboard.Caption = Format(Int(4 - gametimeleft), "0")
                LabelHP.Caption = Format(gamehp / 3 * 50, "0.0")
                LabelTimeLeft.Caption = Format(gametimeleft / 3 * gamecurrentdifficulty, "0.0")

                gamehpanimationtarget = gamehp / 3 * 6210
                If gamehpanimationtarget < 0 Then gamehpanimationtarget = 0
                If gamehpanimationtarget > 6210 Then gamehpanimationtarget = 6210
                gametimeleftprogressbaranimationtarget = gametimeleft / 3 * 6210
                If gametimeleftprogressbaranimationtarget < 0 Then gametimeleftprogressbaranimationtarget = 0
                If gametimeleftprogressbaranimationtarget > 6210 Then gametimeleftprogressbaranimationtarget = 6210
                'Hide popups...
                popupanimationtarget1 = 0: popupanimationtarget2 = 0

                If gametimeleft >= 3 Then
                    gamehp = 50
                    gametimeleft = gamecurrentdifficulty
                    Call GameQuestioner: GoTo TimerTimer_SkipSelectCaseGameStatus_
                End If

                ShapeLightIndicatorOption1.BackColor = &H707070
                ShapeLightIndicatorOption2.BackColor = &H707070
                ShapeLightIndicatorOption3.BackColor = &H707070
            Case 1
                gametimeelapsed = gametimeelapsed + 0.1
                gamehp = gamehp - 0.1
                gametimeleft = gametimeleft - 0.1
                If gamehp > 50 Then gamehp = 50
                If gamehp < gameminimumhp Then gameminimumhp = gamehp

                LabelHP.Caption = Format(gamehp, "0.0")
                LabelTimeLeft.Caption = Format(gametimeleft, "0.0")

                gamehpanimationtarget = gamehp / 50 * 6210
                If gamehpanimationtarget < 0 Then gamehpanimationtarget = 0
                If gamehpanimationtarget > 6210 Then gamehpanimationtarget = 6210
                gametimeleftprogressbaranimationtarget = gametimeleft / gamecurrentdifficulty * 6210
                If gametimeleftprogressbaranimationtarget < 0 Then gametimeleftprogressbaranimationtarget = 0
                If gametimeleftprogressbaranimationtarget > 6210 Then gametimeleftprogressbaranimationtarget = 6210

                'Time up...
                If gametimeleft < 0 Then
                    chosenoption = 4: Call GameRespondent: GoTo TimerTimer_SkipSelectCaseGameStatus_
                End If

                'Mod-AU...
                If FormMainWindow.setmodauswitch = True Then
                    If gametimeleft / gamecurrentdifficulty <= 0.7 Then
                        chosenoption = correctoption: Call GameRespondent: GoTo TimerTimer_SkipSelectCaseGameStatus_
                    End If
                End If
            Case 2
                gametimeelapsed = gametimeelapsed + 0.1
                gamehp = gamehp - 0.1
                gametimeleft = gametimeleft + 0.1
                If gamehp > 50 Then gamehp = 50
                If gamehp < gameminimumhp Then gameminimumhp = gamehp

                LabelHP.Caption = Format(gamehp, "0.0")
                LabelTimeLeft.Caption = Format(gametimeleft / FormMainWindow.setcooldown * gamecurrentdifficulty, "0.0")

                gamehpanimationtarget = gamehp / 50 * 6210
                If gamehpanimationtarget < 0 Then gamehpanimationtarget = 0
                If gamehpanimationtarget > 6210 Then gamehpanimationtarget = 6210
                gametimeleftprogressbaranimationtarget = gametimeleft / FormMainWindow.setcooldown * 6210
                If gametimeleftprogressbaranimationtarget < 0 Then gametimeleftprogressbaranimationtarget = 0
                If gametimeleftprogressbaranimationtarget > 6210 Then gametimeleftprogressbaranimationtarget = 6210

                'Time up...
                If gametimeleft >= FormMainWindow.setcooldown Then
                    'Hide toast and popups...
                    toastanimationtime = 50: popupanimationtarget1 = 0: popupanimationtarget2 = 0
                    Call GameQuestioner: GoTo TimerTimer_SkipSelectCaseGameStatus_
                End If
            Case 4
                GoTo TimerTimer_SkipSelectCaseGameStatus_
            Case 0
                GoTo TimerTimer_SkipSelectCaseGameStatus_
            Case Else
                MsgBox "错误：Variable gamestatus is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End Select

TimerTimer_SkipSelectCaseGameStatus_:

        LabelTimeElapsed.Caption = (Format(Int(gametimeelapsed / 60), "00")) & "' " & (Format((Int(gametimeelapsed) Mod 60), "00")) & "'' " & (Format((gametimeelapsed * 10 Mod 10), "0"))

        'Clear!...
        If gameprogress >= 100 Then
            gameprogress = 100

            If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterClear.wav"

            'CL/FC/AP Achievement...
            If gamemisscount > 0 Then
                LabelToast.Caption = "挑战成功": LabelToast.BackColor = &H90FFFF: LabelToast.ForeColor = &H90D0&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
                FormResult.LabelAchievement.Visible = False
            Else
                If gamegoodcount > 0 Then
                    LabelToast.Caption = "FULL COMBO": LabelToast.BackColor = &HD0FFD0: LabelToast.ForeColor = &H9000&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
                    FormResult.LabelAchievement.Caption = "FULL COMBO": FormResult.LabelAchievement.BackColor = &HD0FFD0: FormResult.LabelAchievement.ForeColor = &H9000&: FormResult.LabelAchievement.Visible = True
                Else
                    If gamegreatcount > 0 Then
                        LabelToast.Caption = "FULL COMBO+": LabelToast.BackColor = &HD0FFD0: LabelToast.ForeColor = &H9000&: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
                        FormResult.LabelAchievement.Caption = "FULL COMBO+": FormResult.LabelAchievement.BackColor = &HD0FFD0: FormResult.LabelAchievement.ForeColor = &H9000&: FormResult.LabelAchievement.Visible = True
                    Else
                        LabelToast.Caption = "ALL PERFECT": LabelToast.BackColor = &HFFF0C0: LabelToast.ForeColor = &HFF9000: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
                        FormResult.LabelAchievement.Caption = "ALL PERFECT": FormResult.LabelAchievement.BackColor = &HFFF0C0: FormResult.LabelAchievement.ForeColor = &HFF9000: FormResult.LabelAchievement.Visible = True
                    End If
                End If
            End If

            'Result...
            FormResult.LabelScore.Caption = Format(gamescore, "00000000")
            FormResult.LabelAccuracy.Caption = Format(gameaccuracy, "0.00") & "%"
            'Ranking...
            If gameaccuracy = 100 Then
                FormResult.LabelRanking.Caption = "SS": FormResult.LabelRanking.ForeColor = &HFFFF&
            Else
                If gameaccuracy >= 95 And gamemisscount = 0 Then
                    FormResult.LabelRanking.Caption = "S+": FormResult.LabelRanking.ForeColor = &HFF9000
                Else
                    If gameaccuracy >= 90 And gamemisscount = 0 Then
                        FormResult.LabelRanking.Caption = "S": FormResult.LabelRanking.ForeColor = &HFF5000
                    Else
                        If gameaccuracy >= 95 Then
                            FormResult.LabelRanking.Caption = "A+": FormResult.LabelRanking.ForeColor = &HD000&
                        Else
                            If gameaccuracy >= 90 Then
                                FormResult.LabelRanking.Caption = "A": FormResult.LabelRanking.ForeColor = &H9000&
                            Else
                                If gameaccuracy >= 80 Then
                                    FormResult.LabelRanking.Caption = "B": FormResult.LabelRanking.ForeColor = &H90FF&
                                Else
                                    If gameaccuracy >= 60 Then
                                        FormResult.LabelRanking.Caption = "C": FormResult.LabelRanking.ForeColor = &HFF&
                                    Else
                                        FormResult.LabelRanking.Caption = "D": FormResult.LabelRanking.ForeColor = &H90&
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            FormResult.LabelPerfectCount.Caption = gameperfectcount & "x"
            FormResult.LabelGreatCount.Caption = gamegreatcount & "x"
            FormResult.LabelGoodCount.Caption = gamegoodcount & "x"
            FormResult.LabelMissCount.Caption = gamemisscount & "x"
            FormResult.LabelTotalCount.Caption = gametotalcount & "x"
            FormResult.LabelBestComboCount.Caption = gamebestcombocount & "x"
            FormResult.LabelAverageReactionTime.Caption = Format(gameaveragereactiontime, "0.000") & "s"
            FormResult.LabelMinimumHP.Caption = Format(gameminimumhp, "0.0")
            FormResult.LabelTimeElapsed.Caption = (Format(Int(gametimeelapsed / 60), "00")) & "' " & (Format((Int(gametimeelapsed) Mod 60), "00")) & "'' " & (Format((gametimeelapsed * 10 Mod 10), "0"))
            If FormMainWindow.setmodsdswitch = True Then
                FormResult.LabelLightIndicatorModSD.BackColor = &HFF9000
                FormResult.LabelLightIndicatorModSD.ForeColor = &HFFFFFF
            Else
                FormResult.LabelLightIndicatorModSD.BackColor = &H707070
                FormResult.LabelLightIndicatorModSD.ForeColor = &HB0B0B0
            End If
            If FormMainWindow.setmodpfswitch = True Then
                FormResult.LabelLightIndicatorModPF.BackColor = &HFF9000
                FormResult.LabelLightIndicatorModPF.ForeColor = &HFFFFFF
            Else
                FormResult.LabelLightIndicatorModPF.BackColor = &H707070
                FormResult.LabelLightIndicatorModPF.ForeColor = &HB0B0B0
            End If
            If FormMainWindow.setmodnfswitch = True Then
                FormResult.LabelLightIndicatorModNF.BackColor = &HFF9000
                FormResult.LabelLightIndicatorModNF.ForeColor = &HFFFFFF
            Else
                FormResult.LabelLightIndicatorModNF.BackColor = &H707070
                FormResult.LabelLightIndicatorModNF.ForeColor = &HB0B0B0
            End If
            If FormMainWindow.setmodapswitch = True Then
                FormResult.LabelLightIndicatorModAP.BackColor = &HFF9000
                FormResult.LabelLightIndicatorModAP.ForeColor = &HFFFFFF
            Else
                FormResult.LabelLightIndicatorModAP.BackColor = &H707070
                FormResult.LabelLightIndicatorModAP.ForeColor = &HB0B0B0
            End If
            If FormMainWindow.setmodauswitch = True Then
                FormResult.LabelLightIndicatorModAU.BackColor = &HFF9000
                FormResult.LabelLightIndicatorModAU.ForeColor = &HFFFFFF
            Else
                FormResult.LabelLightIndicatorModAU.BackColor = &H707070
                FormResult.LabelLightIndicatorModAU.ForeColor = &HB0B0B0
            End If
            resultanimationtime = 0
            FormResult.CmdRetry.Enabled = False: FormResult.CmdBackToHome.Enabled = False
            gameclear = True
            MenuBackToHome.Enabled = False: CmdBackToHome.Enabled = False: MenuStartPauseResume.Enabled = False: CmdStartPauseResume.Enabled = False: FormKanaMaster.MenuStop.Enabled = False: FormKanaMaster.CmdStop.Enabled = False

            gamestatus = 0: Call GameStatusRefresher
        End If

        'Mod-NF...
        If FormMainWindow.setmodnfswitch = True Then
            If gamehp < 0 Then gamehp = 0
        End If

        'Fail...
        If gamehp < 0 Then
            If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterStop.wav"
            LabelToast.Caption = "挑战失败": LabelToast.BackColor = &HD0D0FF: LabelToast.ForeColor = &HFF: LabelToast.Visible = True: toastanimationtime = 0: toastanimationtarget = 960
            gamestatus = 0: Call GameStatusRefresher
        End If

TimerTimer_SkipGameClearJudgement_:

        'Display refresh...
        LabelTimeElapsed.Caption = (Format(Int(gametimeelapsed / 60), "00")) & "' " & (Format((Int(gametimeelapsed) Mod 60), "00")) & "'' " & (Format((gametimeelapsed * 10 Mod 10), "0"))
        LabelTotalCount.Caption = gametotalcount & "x"
        LabelProgress.Caption = Format(gameprogress, "0.00") & "%"
        'DISABLED LINE: LabelScore.Caption = Format(gamescorenumberanimationcurrent, "00000000")
        'DISABLED LINE: LabelHP.Caption = Format(gamehp, "0.0")
        LabelMissCount.Caption = gamemisscount & "x"
        LabelComboCount.Caption = gamecombocount & "x"
        LabelBestComboCount.Caption = gamebestcombocount & "x"
        'DISABLED LINE: LabelTimeLeft.Caption = Format(gametimeleft, "0.0")
        LabelCurrentDifficulty.Caption = Format(gamecurrentdifficulty, "0.0")
        LabelAverageReactionTime.Caption = Format(gameaveragereactiontime, "0.000") & "s"
        LabelAccuracy.Caption = Format(gameaccuracy, "0.00") & "%"
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] GAME ENGINE []

    Public Sub GameStatusRefresher()
        'Prevent MainWindow active during gameplay...
        If gamestatus = 0 Then
            FormMainWindow.Enabled = True
        Else
            FormMainWindow.Enabled = False
            FormMainWindow.Hide
        End If
        'Skip when game clear...
        If gameclear = True Then Exit Sub

        Select Case gamestatus
            Case 0
                'Reset something...
                gamehp = 0: gameminimumhp = 50: gametimeleft = 0
                MenuStartPauseResume.Caption = "开始(&S)": MenuStartPauseResume.Enabled = True: MenuStop.Enabled = False
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "开始(&S)": CmdStop.Enabled = False
                CmdOption1.Caption = "": CmdOption2.Caption = "": CmdOption3.Caption = "": LabelDashboard.Caption = ""
            Case 3
                'Variables like gameprogress will be reset in TimerTimer_Timer...
                MenuStartPauseResume.Caption = "准备中": MenuStartPauseResume.Enabled = False: MenuStop.Enabled = True
                CmdStartPauseResume.Enabled = False: CmdStartPauseResume.Caption = "准备中": CmdStop.Enabled = True
                CmdOption1.Caption = "": CmdOption2.Caption = "": CmdOption3.Caption = ""
            Case 1
                MenuStartPauseResume.Caption = "暂停(&P)": MenuStartPauseResume.Enabled = True: MenuStop.Enabled = True
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "暂停(&P)": CmdStop.Enabled = True
            Case 2
                MenuStartPauseResume.Caption = "暂停(&P)": MenuStartPauseResume.Enabled = True: MenuStop.Enabled = True
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "暂停(&P)": CmdStop.Enabled = True
            Case 4
                gametimeleft = 0
                MenuStartPauseResume.Caption = "继续(&U)": MenuStartPauseResume.Enabled = True: MenuStop.Enabled = True
                CmdStartPauseResume.Enabled = True: CmdStartPauseResume.Caption = "继续(&U)": CmdStop.Enabled = True
                CmdOption1.Caption = "": CmdOption2.Caption = "": CmdOption3.Caption = "": LabelDashboard.Caption = "?"
            Case Else
                MsgBox "错误：Variable gamestatus is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End Select
    End Sub

    Public Sub RandomNumberGenerator()
        Randomize
        lotterynumber = Int((lotterytotal + 1) * Rnd)
        While lotterynumber = 0
            Randomize
            lotterynumber = Int((lotterytotal + 1) * Rnd)
        Wend
    End Sub

    Public Sub GameQuestioner()
        If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterQuestion.wav"

        gamestatus = 1: Call GameStatusRefresher: gametimeleft = gamecurrentdifficulty
        If gameprogress >= 100 Then Exit Sub

        'Clear contents...
        LabelDashboard.Caption = ""
        CmdOption1.Caption = ""
        CmdOption2.Caption = ""
        CmdOption3.Caption = ""
        ShapeLightIndicatorOption1.BackColor = &H707070
        ShapeLightIndicatorOption2.BackColor = &H707070
        ShapeLightIndicatorOption3.BackColor = &H707070

        'Step 1/4: Question...
        lotterytotal = 0: lotterynumber = 0: lotteryquestionlocationX = 0: lotteryquestionlocationY = 0
        Do Until Not (questionrepeatedtimesdata(lotteryquestionlocationX)(lotteryquestionlocationY) >= FormMainWindow.setrepeatedtimes Or questiondata(lotteryquestionlocationX)(lotteryquestionlocationY) = "!!" Or questiondata(lotteryquestionlocationX)(lotteryquestionlocationY) = "--")  'Prevent selecting a repeated block or an empty block...
            lotterytotal = 11: lotterynumber = 0: Do Until FormMainWindow.setkanarangeswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
            lotteryquestionlocationX = lotterynumber
            lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
            lotteryquestionlocationY = lotterynumber
        Loop
        lotteryquestion = questiondata(lotteryquestionlocationX)(lotteryquestionlocationY)
        correspondinganswer = answerdata(lotteryquestionlocationX)(lotteryquestionlocationY)
        LabelDashboard.Caption = lotteryquestion

        'Step 2/4: Correct option...
        lotterytotal = 3: lotterynumber = 0: Call RandomNumberGenerator: correctoption = lotterynumber
        Select Case correctoption
            Case 1
                CmdOption1.Caption = correspondinganswer
            Case 2
                CmdOption2.Caption = correspondinganswer
            Case 3
                CmdOption3.Caption = correspondinganswer
            Case Else
                MsgBox "错误：Variable correctoption is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End Select

        'Step 3/4: Other option 1...
        Select Case correctoption
            Case 1
                lotterytotal = 0: lotterynumber = 0: lotteryanswerlocationX = 0: lotteryanswerlocationY = 0
                Do Until Not (answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = CmdOption1.Caption Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "!!" Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                    lotterytotal = 11: lotterynumber = 0: Do Until FormMainWindow.setkanarangeswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                    lotteryanswerlocationX = lotterynumber
                    lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                    lotteryanswerlocationY = lotterynumber
                Loop
                CmdOption2.Caption = answerdata(lotteryanswerlocationX)(lotteryanswerlocationY)
            Case 2
                lotterytotal = 0: lotterynumber = 0: lotteryanswerlocationX = 0: lotteryanswerlocationY = 0
                Do Until Not (answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = CmdOption2.Caption Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "!!" Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                    lotterytotal = 11: lotterynumber = 0: Do Until FormMainWindow.setkanarangeswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                    lotteryanswerlocationX = lotterynumber
                    lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                    lotteryanswerlocationY = lotterynumber
                Loop
                CmdOption1.Caption = answerdata(lotteryanswerlocationX)(lotteryanswerlocationY)
            Case 3
                lotterytotal = 0: lotterynumber = 0: lotteryanswerlocationX = 0: lotteryanswerlocationY = 0
                Do Until Not (answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = CmdOption3.Caption Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "!!" Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                    lotterytotal = 11: lotterynumber = 0: Do Until FormMainWindow.setkanarangeswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                    lotteryanswerlocationX = lotterynumber
                    lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                    lotteryanswerlocationY = lotterynumber
                Loop
                CmdOption1.Caption = answerdata(lotteryanswerlocationX)(lotteryanswerlocationY)
        End Select

        'Step 4/4: Other option 2...
        Select Case correctoption
            Case 1
                lotterytotal = 0: lotterynumber = 0: lotteryanswerlocationX = 0: lotteryanswerlocationY = 0
                Do Until Not (answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = CmdOption1.Caption Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = CmdOption2.Caption Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "!!" Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                    lotterytotal = 11: lotterynumber = 0: Do Until FormMainWindow.setkanarangeswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                    lotteryanswerlocationX = lotterynumber
                    lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                    lotteryanswerlocationY = lotterynumber
                Loop
                CmdOption3.Caption = answerdata(lotteryanswerlocationX)(lotteryanswerlocationY)
            Case 2
                lotterytotal = 0: lotterynumber = 0: lotteryanswerlocationX = 0: lotteryanswerlocationY = 0
                Do Until Not (answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = CmdOption2.Caption Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = CmdOption1.Caption Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "!!" Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                    lotterytotal = 11: lotterynumber = 0: Do Until FormMainWindow.setkanarangeswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                    lotteryanswerlocationX = lotterynumber
                    lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                    lotteryanswerlocationY = lotterynumber
                Loop
                CmdOption3.Caption = answerdata(lotteryanswerlocationX)(lotteryanswerlocationY)
            Case 3
                lotterytotal = 0: lotterynumber = 0: lotteryanswerlocationX = 0: lotteryanswerlocationY = 0
                Do Until Not (answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = CmdOption3.Caption Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = CmdOption1.Caption Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "!!" Or answerdata(lotteryanswerlocationX)(lotteryanswerlocationY) = "--")  'Prevent selecting a repeated block or an empty block...
                    lotterytotal = 11: lotterynumber = 0: Do Until FormMainWindow.setkanarangeswitch(lotterynumber) = True: Call RandomNumberGenerator: Loop  'Prevent selecting a disabled part...
                    lotteryanswerlocationX = lotterynumber
                    lotterytotal = 16: lotterynumber = 0: Call RandomNumberGenerator
                    lotteryanswerlocationY = lotterynumber
                Loop
                CmdOption2.Caption = answerdata(lotteryanswerlocationX)(lotteryanswerlocationY)
        End Select

        'Mod-AP...
        If FormMainWindow.setmodapswitch = True Then
            Select Case correctoption
                Case 1
                    ShapeLightIndicatorOption1.BackColor = &HFF00&
                    ShapeLightIndicatorOption2.BackColor = &H707070
                    ShapeLightIndicatorOption3.BackColor = &H707070
                Case 2
                    ShapeLightIndicatorOption1.BackColor = &H707070
                    ShapeLightIndicatorOption2.BackColor = &HFF00&
                    ShapeLightIndicatorOption3.BackColor = &H707070
                Case 3
                    ShapeLightIndicatorOption1.BackColor = &H707070
                    ShapeLightIndicatorOption2.BackColor = &H707070
                    ShapeLightIndicatorOption3.BackColor = &HFF00&
                Case Else
                    MsgBox "错误：Variable correctoption is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
            End Select
        End If
    End Sub

    Public Sub GameRespondent()
        'Light indicators...
        Select Case chosenoption
            Case 1
                ShapeLightIndicatorOption1.BackColor = &HFFFFFF
                ShapeLightIndicatorOption2.BackColor = &H707070
                ShapeLightIndicatorOption3.BackColor = &H707070
            Case 2
                ShapeLightIndicatorOption1.BackColor = &H707070
                ShapeLightIndicatorOption2.BackColor = &HFFFFFF
                ShapeLightIndicatorOption3.BackColor = &H707070
            Case 3
                ShapeLightIndicatorOption1.BackColor = &H707070
                ShapeLightIndicatorOption2.BackColor = &H707070
                ShapeLightIndicatorOption3.BackColor = &HFFFFFF
            Case 4
                ShapeLightIndicatorOption1.BackColor = &H707070
                ShapeLightIndicatorOption2.BackColor = &H707070
                ShapeLightIndicatorOption3.BackColor = &H707070
            Case Else
                MsgBox "错误：Variable chosenoption is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End Select

        'Switch game status...
        gamestatus = 2: Call GameStatusRefresher

        'Statistics...
        gametotalcount = gametotalcount + 1
        If gametotalcount > 9999 Then gametotalcount = 9999
        gameaveragereactiontime = (gameaveragereactiontime * (gametotalcount - 1) + (gamecurrentdifficulty - gametimeleft)) / gametotalcount

        'Judgement...
        Select Case correctoption
            Case 1
                LabelPopup1.Left = 840
            Case 2
                LabelPopup1.Left = 6615
            Case 3
                LabelPopup1.Left = 12390
            Case Else
                MsgBox "错误：Variable correctoption is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End Select

        If chosenoption = correctoption Then
            'HP restore...
            gamehp = gamehp + (gamecurrentdifficulty - gametimeleft) * 5
            If gamehp > 50 Then gamehp = 50
            If gamehp < gameminimumhp Then gameminimumhp = gamehp

            'Combo count...
            gamecombocount = gamecombocount + 1
            If gamecombocount > 9999 Then gamecombocount = 9999
            If gamebestcombocount < gamecombocount Then gamebestcombocount = gamecombocount
            If gamebestcombocount > 9999 Then gamebestcombocount = 9999
            If FormMainWindow.setgamemode = 1 Then questionrepeatedtimesdata(lotteryquestionlocationX)(lotteryquestionlocationY) = questionrepeatedtimesdata(lotteryquestionlocationX)(lotteryquestionlocationY) + 1

            'Accuracy...
            If (gametimeleft / gamecurrentdifficulty) >= 0.6 Then
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterPerfect.wav"
                LabelPopup1.Caption = "Perfect": LabelPopup1.BackColor = &HFFF0C0: LabelPopup1.ForeColor = &HFF9000: LabelPopup1.Visible = True: popupanimationtarget1 = 960
                gameperfectcount = gameperfectcount + 1
                If gameperfectcount > 9999 Then gameperfectcount = 9999
                gameaccuracy = (gameaccuracy * (gametotalcount - 1) + 100) / gametotalcount
            End If
            If (gametimeleft / gamecurrentdifficulty) < 0.6 And (gametimeleft / gamecurrentdifficulty) >= 0.2 Then
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterGreat.wav"
                LabelPopup1.Caption = "Great": LabelPopup1.BackColor = &HD0FFD0: LabelPopup1.ForeColor = &H9000&: LabelPopup1.Visible = True: popupanimationtarget1 = 960
                gamegreatcount = gamegreatcount + 1
                If gamegreatcount > 9999 Then gamegreatcount = 9999
                gameaccuracy = (gameaccuracy * (gametotalcount - 1) + 60) / gametotalcount
            End If
            If (gametimeleft / gamecurrentdifficulty) < 0.2 And (gametimeleft / gamecurrentdifficulty) >= 0 Then
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterGood.wav"
                LabelPopup1.Caption = "Good": LabelPopup1.BackColor = &H90FFFF: LabelPopup1.ForeColor = &H90D0&: LabelPopup1.Visible = True: popupanimationtarget1 = 960
                gamegoodcount = gamegoodcount + 1
                If gamegoodcount > 9999 Then gamegoodcount = 9999
                gameaccuracy = (gameaccuracy * (gametotalcount - 1) + 20) / gametotalcount
            End If

            'Mod-PF...
            If FormMainWindow.setmodpfswitch = True Then
                If gamegreatcount > 0 Or gamegoodcount > 0 Or gamemisscount > 0 Then
                    gamehp = -1
                End If
            End If

            'Score...
            If (gametimeleft / gamecurrentdifficulty) >= 0.6 Then
                LabelPopup2.Caption = "+300": LabelPopup2.BackColor = &HFFF0C0: LabelPopup2.ForeColor = &HFF9000: LabelPopup2.Visible = True: popupanimationtarget2 = 750
                gamescore = gamescore + gamecombocount * 300
            End If
            If (gametimeleft / gamecurrentdifficulty) < 0.6 And (gametimeleft / gamecurrentdifficulty) >= 0.2 Then
                LabelPopup2.Caption = "+100": LabelPopup2.BackColor = &HD0FFD0: LabelPopup2.ForeColor = &H9000&: LabelPopup2.Visible = True: popupanimationtarget2 = 750
                gamescore = gamescore + gamecombocount * 100
            End If
            If (gametimeleft / gamecurrentdifficulty) < 0.2 And (gametimeleft / gamecurrentdifficulty) >= 0 Then
                LabelPopup2.Caption = "+50": LabelPopup2.BackColor = &H90FFFF: LabelPopup2.ForeColor = &H90D0&: LabelPopup2.Visible = True: popupanimationtarget2 = 750
                gamescore = gamescore + gamecombocount * 50
            End If
            If gamescore > 99999999 Then gamescore = 99999999
        Else
            'HP drain...
            gamehp = gamehp - FormMainWindow.setmistakehpdrain
            If gamehp > 50 Then gamehp = 50
            If gamehp < gameminimumhp Then gameminimumhp = gamehp

            'Combo reset... But do not reset the best combo (gamebestcombocount)...
            gamecombocount = 0

            'Accuracy...
            If chosenoption = 4 Then
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterTimeUpMiss.wav"
            Else
                If FormMainWindow.soundswitch = True Then FormMainWindow.WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterWrongChoiceMiss.wav"
            End If
            LabelPopup1.Caption = "×": LabelPopup1.BackColor = &HD0D0FF: LabelPopup1.ForeColor = &HFF: LabelPopup1.Visible = True: popupanimationtarget1 = 960
            gamemisscount = gamemisscount + 1
            If gamemisscount > 9999 Then gamemisscount = 9999
            gameaccuracy = (gameaccuracy * (gametotalcount - 1) + 0) / gametotalcount

            'Score...
            'Do nothing here!

            'Mod-SD...
            If FormMainWindow.setmodsdswitch = True Then
                gamehp = -1
            End If

            'Mod-PF...
            If FormMainWindow.setmodpfswitch = True Then
                If gamegreatcount > 0 Or gamegoodcount > 0 Or gamemisscount > 0 Then
                    gamehp = -1
                End If
            End If
        End If

        gametimeleft = 0: Call TimerTimer_Timer
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerToastAnimation_Timer()
        'Timer...
        toastanimationtime = toastanimationtime + 1
        If toastanimationtime >= 50 Then
            toastanimationtime = 0
            toastanimationtarget = 0
        End If

        If LabelToast.Height = toastanimationtarget Then Exit Sub

        Select Case FormMainWindow.setanimationswitch
            Case True
                If LabelToast.Height > toastanimationtarget Then LabelToast.Height = LabelToast.Height - Abs(LabelToast.Height - toastanimationtarget) / 4
                If LabelToast.Height < toastanimationtarget Then LabelToast.Height = LabelToast.Height + Abs(LabelToast.Height - toastanimationtarget) / 4
                If Abs(LabelToast.Height - toastanimationtarget) < 20 Then LabelToast.Height = toastanimationtarget
                LabelToast.Top = 5100 - LabelToast.Height / 2
            Case False
                LabelToast.Height = toastanimationtarget
                LabelToast.Top = 5100 - LabelToast.Height / 2
        End Select

        If toastanimationtarget = 0 And LabelToast.Height < 20 Then LabelToast.Visible = False
    End Sub

    Public Sub TimerPopupAnimation_Timer()
        Select Case FormMainWindow.setanimationswitch
            Case True
                If LabelPopup1.Height = popupanimationtarget1 Then GoTo TimerPopupAnimation_Skip1_
                If LabelPopup1.Height > popupanimationtarget1 Then LabelPopup1.Height = LabelPopup1.Height - Abs(LabelPopup1.Height - popupanimationtarget1) / 4
                If LabelPopup1.Height < popupanimationtarget1 Then LabelPopup1.Height = LabelPopup1.Height + Abs(LabelPopup1.Height - popupanimationtarget1) / 4
                If Abs(LabelPopup1.Height - popupanimationtarget1) < 20 Then LabelPopup1.Height = popupanimationtarget1
                LabelPopup1.Top = 8625 - LabelPopup1.Height
TimerPopupAnimation_Skip1_:
                If LabelPopup2.Height = popupanimationtarget2 Then GoTo TimerPopupAnimation_Skip2_
                If LabelPopup2.Height > popupanimationtarget2 Then LabelPopup2.Height = LabelPopup2.Height - Abs(LabelPopup2.Height - popupanimationtarget2) / 4
                If LabelPopup2.Height < popupanimationtarget2 Then LabelPopup2.Height = LabelPopup2.Height + Abs(LabelPopup2.Height - popupanimationtarget2) / 4
                If Abs(LabelPopup2.Height - popupanimationtarget2) < 20 Then LabelPopup2.Height = popupanimationtarget2
                LabelPopup2.Top = 1800 - LabelPopup2.Height
TimerPopupAnimation_Skip2_:
            Case False
                If LabelPopup1.Height = popupanimationtarget1 Then GoTo TimerPopupAnimation_Skip3_
                LabelPopup1.Height = popupanimationtarget1
                LabelPopup1.Top = 8625 - LabelPopup1.Height
TimerPopupAnimation_Skip3_:
                If LabelPopup2.Height = popupanimationtarget2 Then GoTo TimerPopupAnimation_Skip4_
                LabelPopup2.Height = popupanimationtarget2
                LabelPopup2.Top = 1800 - LabelPopup2.Height
TimerPopupAnimation_Skip4_:
        End Select

        If popupanimationtarget1 = 0 And LabelPopup1.Height < 20 Then LabelPopup1.Visible = False
        If popupanimationtarget2 = 0 And LabelPopup2.Height < 20 Then LabelPopup2.Visible = False
    End Sub

    Public Sub TimerProgressbarAnimation_Timer()
        Select Case FormMainWindow.setanimationswitch
            Case True
                If ShapeProgressProgressbar.Width = gameprogressanimationtarget Then GoTo TimerProgressbarAnimation_Skip1_
                If ShapeProgressProgressbar.Width > gameprogressanimationtarget Then ShapeProgressProgressbar.Width = ShapeProgressProgressbar.Width - Abs(ShapeProgressProgressbar.Width - gameprogressanimationtarget) / 4
                If ShapeProgressProgressbar.Width < gameprogressanimationtarget Then ShapeProgressProgressbar.Width = ShapeProgressProgressbar.Width + Abs(ShapeProgressProgressbar.Width - gameprogressanimationtarget) / 4
                If Abs(ShapeProgressProgressbar.Width - gameprogressanimationtarget) < 10 Then ShapeProgressProgressbar.Width = gameprogressanimationtarget
TimerProgressbarAnimation_Skip1_:
                If ShapeHPProgressbar.Height = gamehpanimationtarget Then GoTo TimerProgressbarAnimation_Skip2_
                If ShapeHPProgressbar.Height > gamehpanimationtarget Then ShapeHPProgressbar.Height = ShapeHPProgressbar.Height - Abs(ShapeHPProgressbar.Height - gamehpanimationtarget) / 4
                If ShapeHPProgressbar.Height < gamehpanimationtarget Then ShapeHPProgressbar.Height = ShapeHPProgressbar.Height + Abs(ShapeHPProgressbar.Height - gamehpanimationtarget) / 4
                If Abs(ShapeHPProgressbar.Height - gamehpanimationtarget) < 10 Then ShapeHPProgressbar.Height = gamehpanimationtarget
                ShapeHPProgressbar.Top = 8205 - ShapeHPProgressbar.Height
TimerProgressbarAnimation_Skip2_:
                If ShapeTimeLeftProgressbar.Height = gametimeleftprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip3_
                If ShapeTimeLeftProgressbar.Height > gametimeleftprogressbaranimationtarget Then ShapeTimeLeftProgressbar.Height = ShapeTimeLeftProgressbar.Height - Abs(ShapeTimeLeftProgressbar.Height - gametimeleftprogressbaranimationtarget) / 4
                If ShapeTimeLeftProgressbar.Height < gametimeleftprogressbaranimationtarget Then ShapeTimeLeftProgressbar.Height = ShapeTimeLeftProgressbar.Height + Abs(ShapeTimeLeftProgressbar.Height - gametimeleftprogressbaranimationtarget) / 4
                If Abs(ShapeTimeLeftProgressbar.Height - gametimeleftprogressbaranimationtarget) < 10 Then ShapeTimeLeftProgressbar.Height = gametimeleftprogressbaranimationtarget
                ShapeTimeLeftProgressbar.Top = 8205 - ShapeTimeLeftProgressbar.Height
TimerProgressbarAnimation_Skip3_:
            Case False
                If ShapeProgressProgressbar.Width = gameprogressanimationtarget Then GoTo TimerProgressbarAnimation_Skip4_
                ShapeProgressProgressbar.Width = gameprogressanimationtarget
TimerProgressbarAnimation_Skip4_:
                If ShapeHPProgressbar.Height = gamehpanimationtarget Then GoTo TimerProgressbarAnimation_Skip5_
                ShapeHPProgressbar.Height = gamehpanimationtarget
                ShapeHPProgressbar.Top = 8205 - ShapeHPProgressbar.Height
TimerProgressbarAnimation_Skip5_:
                If ShapeTimeLeftProgressbar.Height = gametimeleftprogressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip6_
                ShapeTimeLeftProgressbar.Height = gametimeleftprogressbaranimationtarget
                ShapeTimeLeftProgressbar.Top = 8205 - ShapeTimeLeftProgressbar.Height
TimerProgressbarAnimation_Skip6_:
        End Select
    End Sub

    Public Sub TimerSakuraAnimation_Timer()
        If (gamestatus = 1 Or gamestatus = 2 Or gamestatus = 3) Then
            sakuratargetspeed = 6 - gamecurrentdifficulty
        Else
            sakuratargetspeed = 0
        End If

        'Locate (630+ShapeProgressProgressbar.Width, 1110) ...
        LineSakura1.X1 = 630 + ShapeProgressProgressbar.Width: LineSakura1.Y1 = 1110
        LineSakura2.X1 = 630 + ShapeProgressProgressbar.Width: LineSakura2.Y1 = 1110
        LineSakura3.X1 = 630 + ShapeProgressProgressbar.Width: LineSakura3.Y1 = 1110
        LineSakura4.X1 = 630 + ShapeProgressProgressbar.Width: LineSakura4.Y1 = 1110
        LineSakura5.X1 = 630 + ShapeProgressProgressbar.Width: LineSakura5.Y1 = 1110

        'Make flower (Length set to 250) ...
        sakuracurrentangle2 = sakuracurrentangle - 360 / 5 * 1
        sakuracurrentangle3 = sakuracurrentangle - 360 / 5 * 2
        sakuracurrentangle4 = sakuracurrentangle - 360 / 5 * 3
        sakuracurrentangle5 = sakuracurrentangle - 360 / 5 * 4
        While sakuracurrentangle2 < -180: sakuracurrentangle2 = sakuracurrentangle2 + 360: Wend
        While sakuracurrentangle3 < -180: sakuracurrentangle3 = sakuracurrentangle3 + 360: Wend
        While sakuracurrentangle4 < -180: sakuracurrentangle4 = sakuracurrentangle4 + 360: Wend
        While sakuracurrentangle5 < -180: sakuracurrentangle5 = sakuracurrentangle5 + 360: Wend

        LineSakura1.X2 = LineSakura1.X1 + 250 * Cos(3.14 / 180 * sakuracurrentangle)
        LineSakura1.Y2 = LineSakura1.Y1 - 250 * Sin(3.14 / 180 * sakuracurrentangle)
        LineSakura2.X2 = LineSakura2.X1 + 250 * Cos(3.14 / 180 * sakuracurrentangle2)
        LineSakura2.Y2 = LineSakura2.Y1 - 250 * Sin(3.14 / 180 * sakuracurrentangle2)
        LineSakura3.X2 = LineSakura3.X1 + 250 * Cos(3.14 / 180 * sakuracurrentangle3)
        LineSakura3.Y2 = LineSakura3.Y1 - 250 * Sin(3.14 / 180 * sakuracurrentangle3)
        LineSakura4.X2 = LineSakura4.X1 + 250 * Cos(3.14 / 180 * sakuracurrentangle4)
        LineSakura4.Y2 = LineSakura4.Y1 - 250 * Sin(3.14 / 180 * sakuracurrentangle4)
        LineSakura5.X2 = LineSakura5.X1 + 250 * Cos(3.14 / 180 * sakuracurrentangle5)
        LineSakura5.Y2 = LineSakura5.Y1 - 250 * Sin(3.14 / 180 * sakuracurrentangle5)

        'Prevent constant blinking...
        If (sakuratargetspeed = 0 And sakuracurrentspeed = 0) Then Exit Sub

        'Spin...
        Select Case FormMainWindow.setanimationswitch
            Case True
                sakuracurrentangle = sakuracurrentangle - sakuracurrentspeed
                If sakuracurrentangle <= -180 Then sakuracurrentangle = sakuracurrentangle + 360
            Case False
                sakuracurrentangle = 90
        End Select

        'Adjust spinning speed...
        If sakuracurrentspeed < sakuratargetspeed Then sakuracurrentspeed = sakuracurrentspeed + 0.1
        If sakuracurrentspeed > sakuratargetspeed Then sakuracurrentspeed = sakuracurrentspeed - 0.05
        If sakuracurrentspeed < 0 Then sakuracurrentspeed = 0
        If sakuracurrentspeed > 6 Then sakuracurrentspeed = 6
    End Sub

    Public Sub TimerNumberAnimation_Timer()
        Select Case FormMainWindow.setanimationswitch
            Case True
                If gamescorenumberanimationcurrent = gamescore Then GoTo TimerNumberAnimation_Skip1_
                If gamescorenumberanimationcurrent > gamescore Then gamescorenumberanimationcurrent = gamescorenumberanimationcurrent - Abs(gamescorenumberanimationcurrent - gamescore) / 3
                If gamescorenumberanimationcurrent < gamescore Then gamescorenumberanimationcurrent = gamescorenumberanimationcurrent + Abs(gamescorenumberanimationcurrent - gamescore) / 3
                If Abs(gamescorenumberanimationcurrent - gamescore) < 2 Then gamescorenumberanimationcurrent = gamescore
TimerNumberAnimation_Skip1_:
            Case False
                If gamescorenumberanimationcurrent = gamescore Then GoTo TimerNumberAnimation_Skip2_
                gamescorenumberanimationcurrent = gamescore
TimerNumberAnimation_Skip2_:
        End Select
        LabelScore.Caption = Format(gamescorenumberanimationcurrent, "00000000")
    End Sub

    Public Sub TimerArrowAnimation_Timer()
        'Timer...
        arrowanimationtime = arrowanimationtime + 0.15
        If arrowanimationtime >= 6.28 Then arrowanimationtime = 0

        'Start Arrow (Yellow)...
        If (gamestatus = 0 Or gamestatus = 4) And gameclear = False Then
            LabelStartArrow.Visible = True
            LabelStartArrow.Top = 525 + 200 * Sin(arrowanimationtime)
        Else
            LabelStartArrow.Visible = False
        End If

        'Ready Arrow (Red) (3x)...
        If gamestatus = 3 And gametimeleft <= 2 Then
            If Abs(arrowanimationtime * 6 / 6.28) Mod 2 = 0 Then
                LabelReadyArrow1.Visible = True: LabelReadyArrow2.Visible = True: LabelReadyArrow3.Visible = True
            Else
                LabelReadyArrow1.Visible = False: LabelReadyArrow2.Visible = False: LabelReadyArrow3.Visible = False
            End If
        Else
            LabelReadyArrow1.Visible = False: LabelReadyArrow2.Visible = False: LabelReadyArrow3.Visible = False
        End If
    End Sub

    Public Sub TimerResultAnimation_Timer()
        If gameclear = False Then Exit Sub

        'Timer...
        resultanimationtime = resultanimationtime + 1
        If resultanimationtime > 200 Then resultanimationtime = 200

        'Step 1/6: Initialize and wait...
        If resultanimationtime = 20 Then
            FormResult.TimerWindowAnimation.Enabled = False
            FormResult.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
            FormResult.ShapeEdge.Width = 0
            FormResult.ShapeEdge.Height = 0
            FormResult.windowanimationtargetleft = (Screen.Width / 2) - (18690 / 2)
            FormResult.windowanimationtargettop = (Screen.Height / 2) - (12435 / 2)
            FormResult.windowanimationtargetwidth = 18690
            FormResult.windowanimationtargetheight = 12345
            FormResult.Hide
        End If

        'Step 2/6: Prepare title...
        If resultanimationtime = 90 Then
            FormResult.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
            FormResult.ShapeEdge.Width = 0
            FormResult.ShapeEdge.Height = 0
            FormResult.Show
        End If
        If resultanimationtime >= 90 And resultanimationtime < 100 Then
            FormResult.Top = (Screen.Height / 2) - (((1 - (3 / 4) ^ ((resultanimationtime - 90) / 10 * 20)) * 1360) / 2)
            FormResult.Height = ((1 - (3 / 4) ^ ((resultanimationtime - 90) / 10 * 20)) * 1360)
            FormResult.ShapeEdge.Height = ((1 - (3 / 4) ^ ((resultanimationtime - 90) / 10 * 20)) * 1360)
        End If

        'Step 3/6: Show title...
        If resultanimationtime = 100 Then
            FormResult.Move (Screen.Width / 2), (Screen.Height / 2) - (1360 / 2), 0, 1360
            FormResult.ShapeEdge.Width = 0
            FormResult.ShapeEdge.Height = 1360
            FormResult.LabelTitle1.Left = 420
            FormResult.LabelTitle2.Left = 3675
            FormResult.LabelTitle3.Left = 7350
            FormResult.LabelTitle4.Left = 11025
            FormResult.LabelTitle5.Left = 14385
            FormResult.LabelTitle6.Left = 17535
        End If
        If resultanimationtime >= 100 And resultanimationtime < 130 Then
            FormResult.Left = (Screen.Width / 2) - (((1 - (3 / 4) ^ ((resultanimationtime - 100) / 30 * 20)) * 18690) / 2)
            FormResult.Width = ((1 - (3 / 4) ^ ((resultanimationtime - 100) / 30 * 20)) * 18690)
            FormResult.ShapeEdge.Width = ((1 - (3 / 4) ^ ((resultanimationtime - 100) / 30 * 20)) * 18690)
        End If

        'Step 4/6: Gather title letters...
        If resultanimationtime = 130 Then
            FormResult.Move (Screen.Width / 2) - (18690 / 2), (Screen.Height / 2) - (1360 / 2), 18690, 1360
            FormResult.ShapeEdge.Width = 18690
            FormResult.ShapeEdge.Height = 1360
        End If
        If resultanimationtime >= 130 And resultanimationtime < 150 Then
            FormResult.LabelTitle1.Left = 420 + (1 - (3 / 4) ^ ((resultanimationtime - 130) / 20 * 20)) * Abs(7770 - 420)
            FormResult.LabelTitle2.Left = 3675 + (1 - (3 / 4) ^ ((resultanimationtime - 130) / 20 * 20)) * Abs(8400 - 3675)
            FormResult.LabelTitle3.Left = 7350 + (1 - (3 / 4) ^ ((resultanimationtime - 130) / 20 * 20)) * Abs(9030 - 7350)
            FormResult.LabelTitle4.Left = 11025 - (1 - (3 / 4) ^ ((resultanimationtime - 130) / 20 * 20)) * Abs(9660 - 11025)
            FormResult.LabelTitle5.Left = 14385 - (1 - (3 / 4) ^ ((resultanimationtime - 130) / 20 * 20)) * Abs(10290 - 14385)
            FormResult.LabelTitle6.Left = 17535 - (1 - (3 / 4) ^ ((resultanimationtime - 130) / 20 * 20)) * Abs(10815 - 17535)
        End If

        'Step 5/6: Show the entire window...
        If resultanimationtime = 150 Then
            FormResult.LabelTitle1.Left = 7770
            FormResult.LabelTitle2.Left = 8400
            FormResult.LabelTitle3.Left = 9030
            FormResult.LabelTitle4.Left = 9660
            FormResult.LabelTitle5.Left = 10290
            FormResult.LabelTitle6.Left = 10815
        End If
        If resultanimationtime >= 150 And resultanimationtime < 180 Then
            FormResult.Top = (Screen.Height / 2) - ((1360 + (1 - (3 / 4) ^ ((resultanimationtime - 150) / 30 * 20)) * (12435 - 1360)) / 2)
            FormResult.Height = (1360 + (1 - (3 / 4) ^ ((resultanimationtime - 150) / 30 * 20)) * (12435 - 1360))
            FormResult.ShapeEdge.Height = (1360 + (1 - (3 / 4) ^ ((resultanimationtime - 150) / 30 * 20)) * (12435 - 1360))
        End If

        'Step 6/6: Finish...
        If resultanimationtime = 180 Then
            FormResult.TimerWindowAnimation.Enabled = True
            FormResult.Move (Screen.Width / 2) - (18690 / 2), (Screen.Height / 2) - (12435 / 2), 18690, 12435
            FormResult.ShapeEdge.Width = 18690
            FormResult.ShapeEdge.Height = 12435
            FormResult.windowanimationtargetleft = (Screen.Width / 2) - (18690 / 2)
            FormResult.windowanimationtargettop = (Screen.Height / 2) - (12435 / 2)
            FormResult.windowanimationtargetwidth = 18690
            FormResult.windowanimationtargetheight = 12435

            FormResult.CmdRetry.Enabled = True: FormResult.CmdBackToHome.Enabled = True
        End If
    End Sub
