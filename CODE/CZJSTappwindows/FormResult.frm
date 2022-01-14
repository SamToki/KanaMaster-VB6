VERSION 5.00
Begin VB.Form FormResult 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   0  'None
   Caption         =   "KanaMaster"
   ClientHeight    =   12435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18690
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
   Icon            =   "FormResult.frx":0000
   LinkTopic       =   "FormResult"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormResult.frx":23D2
   MousePointer    =   99  'Custom
   ScaleHeight     =   12435
   ScaleWidth      =   18690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdRetry 
      Caption         =   "再玩一次!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   14385
      MouseIcon       =   "FormResult.frx":2524
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   8820
      Width           =   3690
   End
   Begin VB.CommandButton CmdBackToHome 
      Cancel          =   -1  'True
      Caption         =   "返回主页"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   14385
      MouseIcon       =   "FormResult.frx":2676
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   10395
      Width           =   3690
   End
   Begin VB.Timer TimerWindowAnimation 
      Interval        =   1
      Left            =   18375
      Top             =   12075
   End
   Begin VB.Label LabelRanking 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   255.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4635
      Left            =   12180
      MouseIcon       =   "FormResult.frx":27C8
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "评级标准详见 主页→帮助。"
      Top             =   945
      Width           =   6105
   End
   Begin VB.Label LabelLightIndicatorModAU 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "AU"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   495
      Left            =   17430
      MouseIcon       =   "FormResult.frx":291A
      MousePointer    =   99  'Custom
      TabIndex        =   35
      ToolTipText     =   "(Auto)"
      Top             =   5775
      Width           =   645
   End
   Begin VB.Label LabelLightIndicatorModPF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "PF"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   495
      Left            =   15225
      MouseIcon       =   "FormResult.frx":2A6C
      MousePointer    =   99  'Custom
      TabIndex        =   32
      ToolTipText     =   "(Perfect)"
      Top             =   5775
      Width           =   645
   End
   Begin VB.Label LabelLightIndicatorModNF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "NF"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   495
      Left            =   15960
      MouseIcon       =   "FormResult.frx":2BBE
      MousePointer    =   99  'Custom
      TabIndex        =   33
      ToolTipText     =   "(No-Fail)"
      Top             =   5775
      Width           =   645
   End
   Begin VB.Label LabelLightIndicatorModAP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "AP"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   495
      Left            =   16695
      MouseIcon       =   "FormResult.frx":2D10
      MousePointer    =   99  'Custom
      TabIndex        =   34
      ToolTipText     =   "(Autopilot)"
      Top             =   5775
      Width           =   645
   End
   Begin VB.Label LabelLightIndicatorModSD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00707070&
      Caption         =   "SD"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B0B0&
      Height          =   495
      Left            =   14490
      MouseIcon       =   "FormResult.frx":2E62
      MousePointer    =   99  'Custom
      TabIndex        =   31
      ToolTipText     =   "(Sudden Death)"
      Top             =   5775
      Width           =   645
   End
   Begin VB.Label LabelTimeElapsed 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "--' --'' -"
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
      Left            =   9660
      TabIndex        =   30
      Top             =   6090
      Width           =   3900
   End
   Begin VB.Label LabelTimeElapsedTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "时长"
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
      Left            =   9030
      TabIndex        =   29
      Top             =   5775
      Width           =   960
   End
   Begin VB.Label LabelMinimumHP 
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
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   5460
      TabIndex        =   28
      Top             =   10500
      Width           =   2640
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
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   5460
      MouseIcon       =   "FormResult.frx":2FB4
      MousePointer    =   99  'Custom
      TabIndex        =   26
      ToolTipText     =   "「均速」指平均反应用时。"
      Top             =   9030
      Width           =   2640
   End
   Begin VB.Label LabelBestComboCount 
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
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   5460
      TabIndex        =   24
      Top             =   7560
      Width           =   2640
   End
   Begin VB.Label LabelTotalCount 
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
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   5460
      TabIndex        =   22
      Top             =   6090
      Width           =   2640
   End
   Begin VB.Label LabelMinimumHPTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "最低血量"
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
      Left            =   4830
      TabIndex        =   27
      Top             =   10185
      Width           =   1800
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
      Left            =   4830
      MouseIcon       =   "FormResult.frx":3106
      MousePointer    =   99  'Custom
      TabIndex        =   25
      ToolTipText     =   "(Average Reaction Time)"
      Top             =   8715
      Width           =   960
   End
   Begin VB.Label LabelBestComboCountTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "最大连击数"
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
      Left            =   4830
      MouseIcon       =   "FormResult.frx":3258
      MousePointer    =   99  'Custom
      TabIndex        =   23
      ToolTipText     =   "(Best Combo)"
      Top             =   7245
      Width           =   2220
   End
   Begin VB.Label LabelTotalCountTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "总数"
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
      Left            =   4830
      TabIndex        =   21
      Top             =   5775
      Width           =   960
   End
   Begin VB.Label LabelMissCount 
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
      ForeColor       =   &H000000FF&
      Height          =   960
      Left            =   1260
      TabIndex        =   20
      Top             =   10500
      Width           =   2640
   End
   Begin VB.Label LabelGoodCount 
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
      ForeColor       =   &H000090D0&
      Height          =   960
      Left            =   1260
      TabIndex        =   18
      Top             =   9030
      Width           =   2640
   End
   Begin VB.Label LabelGreatCount 
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
      ForeColor       =   &H00009000&
      Height          =   960
      Left            =   1260
      TabIndex        =   16
      Top             =   7560
      Width           =   2640
   End
   Begin VB.Label LabelPerfectCount 
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
      ForeColor       =   &H00FF9000&
      Height          =   960
      Left            =   1260
      TabIndex        =   14
      Top             =   6090
      Width           =   2640
   End
   Begin VB.Label LabelMissCountTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Miss"
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
      Left            =   630
      TabIndex        =   19
      Top             =   10185
      Width           =   1065
   End
   Begin VB.Label LabelGoodCountTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Good"
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
      Left            =   630
      TabIndex        =   17
      Top             =   8715
      Width           =   1275
   End
   Begin VB.Label LabelGreatCountTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Great"
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
      Left            =   630
      TabIndex        =   15
      Top             =   7245
      Width           =   1275
   End
   Begin VB.Label LabelPerfectCountTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Perfect"
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
      Left            =   630
      TabIndex        =   13
      Top             =   5775
      Width           =   1590
   End
   Begin VB.Label LabelRankingTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "评级"
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
      Left            =   17115
      MouseIcon       =   "FormResult.frx":33AA
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "(Ranking)"
      Top             =   1575
      Width           =   960
   End
   Begin VB.Label LabelAchievement 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "-------- --------"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   7875
      TabIndex        =   10
      Top             =   3780
      Width           =   4110
   End
   Begin VB.Label LabelAccuracy 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "---.--%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2010
      Left            =   1260
      TabIndex        =   9
      Top             =   3570
      Width           =   6420
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
      Left            =   630
      MouseIcon       =   "FormResult.frx":34FC
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "(Accuracy)"
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Label LabelScoreTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "得分"
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
      Left            =   630
      MouseIcon       =   "FormResult.frx":364E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "(Score)"
      Top             =   1575
      Width           =   960
   End
   Begin VB.Label LabelScore 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "--------"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   63.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Left            =   1260
      TabIndex        =   7
      Top             =   1890
      Width           =   5265
   End
   Begin VB.Label LabelTitle6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D000D0&
      Height          =   960
      Left            =   17535
      TabIndex        =   5
      Top             =   315
      Width           =   750
   End
   Begin VB.Label LabelTitle5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF9000&
      Height          =   960
      Left            =   14385
      TabIndex        =   4
      Top             =   210
      Width           =   750
   End
   Begin VB.Label LabelTitle4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000D000&
      Height          =   960
      Left            =   11025
      TabIndex        =   3
      Top             =   210
      Width           =   750
   End
   Begin VB.Label LabelTitle3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   960
      Left            =   7350
      TabIndex        =   2
      Top             =   210
      Width           =   750
   End
   Begin VB.Label LabelTitle2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000090FF&
      Height          =   960
      Left            =   3675
      TabIndex        =   1
      Top             =   315
      Width           =   750
   End
   Begin VB.Label LabelTitle1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Left            =   420
      TabIndex        =   0
      Top             =   210
      Width           =   750
   End
   Begin VB.Shape ShapeEdge 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   12435
      Left            =   0
      Top             =   0
      Width           =   18690
   End
End
Attribute VB_Name = "FormResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Public windowanimationtargetleft As Integer
Public windowanimationtargettop As Integer
Public windowanimationtargetwidth As Integer
Public windowanimationtargetheight As Integer

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    Public Sub CmdRetry_Click()
        Select Case FormMainWindow.setgamemode
            Case 1
                FormKanaMaster.gameclear = False
                FormKanaMaster.MenuBackToHome.Enabled = True: FormKanaMaster.CmdBackToHome.Enabled = True: FormKanaMaster.MenuStartPauseResume.Enabled = True: FormKanaMaster.CmdStartPauseResume.Enabled = True: FormKanaMaster.MenuStop.Enabled = True: FormKanaMaster.CmdStop.Enabled = True
                Call FormKanaMaster.MenuStartPauseResume_Click
            'Case 2
                '?????
            'Case 3
                '?????
            'Case 4
                '?????
            Case Else
                MsgBox "错误：Variable setgamemode is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End Select
    End Sub
    Public Sub CmdBackToHome_Click()
        Select Case FormMainWindow.setgamemode
            Case 1
                FormKanaMaster.gameclear = False
                FormKanaMaster.MenuBackToHome.Enabled = True: FormKanaMaster.CmdBackToHome.Enabled = True: FormKanaMaster.MenuStartPauseResume.Enabled = True: FormKanaMaster.CmdStartPauseResume.Enabled = True: FormKanaMaster.MenuStop.Enabled = True: FormKanaMaster.CmdStop.Enabled = True
                Call FormKanaMaster.MenuStartPauseResume_Click
                Call FormKanaMaster.TimerTimer_Timer
                Call FormKanaMaster.MenuStop_Click
                Call FormKanaMaster.MenuBackToHome_Click
            'Case 2
                '?????
            'Case 3
                '?????
            'Case 4
                '?????
            Case Else
                MsgBox "错误：Variable setgamemode is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End Select
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerWindowAnimation_Timer()
        If ((Me.Left = windowanimationtargetleft) And (Me.Top = windowanimationtargettop) And (Me.Width = windowanimationtargetwidth) And (Me.Height = windowanimationtargetheight)) Then Exit Sub

        Select Case FormMainWindow.setanimationswitch
            Case True
                If Me.Left > windowanimationtargetleft Then Me.Left = Me.Left - Abs(Me.Left - windowanimationtargetleft) / 4
                If Me.Left < windowanimationtargetleft Then Me.Left = Me.Left + Abs(Me.Left - windowanimationtargetleft) / 4
                If Me.Top > windowanimationtargettop Then Me.Top = Me.Top - Abs(Me.Top - windowanimationtargettop) / 4
                If Me.Top < windowanimationtargettop Then Me.Top = Me.Top + Abs(Me.Top - windowanimationtargettop) / 4
                If Me.Width > windowanimationtargetwidth Then Me.Width = Me.Width - Abs(Me.Width - windowanimationtargetwidth) / 4
                If Me.Width < windowanimationtargetwidth Then Me.Width = Me.Width + Abs(Me.Width - windowanimationtargetwidth) / 4
                If Me.Height > windowanimationtargetheight Then Me.Height = Me.Height - Abs(Me.Height - windowanimationtargetheight) / 4
                If Me.Height < windowanimationtargetheight Then Me.Height = Me.Height + Abs(Me.Height - windowanimationtargetheight) / 4
                If Abs(Me.Left - windowanimationtargetleft) < 10 Then Me.Left = windowanimationtargetleft
                If Abs(Me.Top - windowanimationtargettop) < 10 Then Me.Top = windowanimationtargettop
                If Abs(Me.Width - windowanimationtargetwidth) < 10 Then Me.Width = windowanimationtargetwidth
                If Abs(Me.Height - windowanimationtargetheight) < 10 Then Me.Height = windowanimationtargetheight
            Case False
                Me.Move windowanimationtargetleft, windowanimationtargettop, windowanimationtargetwidth, windowanimationtargetheight
        End Select

        If windowanimationtargetheight = 0 And Me.Height < 10 Then Me.Hide
    End Sub
