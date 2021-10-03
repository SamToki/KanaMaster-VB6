VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormMainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "假名征服者(KanaMaster) v0.52chs"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   15345
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
   Icon            =   "FormMainWindow.frx":0000
   LinkTopic       =   "FormMainWindow"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMainWindow.frx":23D2
   MousePointer    =   99  'Custom
   ScaleHeight     =   10515
   ScaleWidth      =   15345
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame FramePresets 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "预设"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   210
      TabIndex        =   7
      Top             =   2940
      Width           =   9045
      Begin VB.CommandButton CmdPresetAllKana 
         Caption         =   "全部假名"
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
         Left            =   7770
         MouseIcon       =   "FormMainWindow.frx":2524
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton CmdPresetOnlyKatakana 
         Caption         =   "仅片假名"
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
         Left            =   6720
         MouseIcon       =   "FormMainWindow.frx":2676
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton CmdPresetOnlyHiragana 
         Caption         =   "仅平假名"
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
         Left            =   5670
         MouseIcon       =   "FormMainWindow.frx":27C8
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton CmdPresetDefaultKanaRange 
         Caption         =   "默认范围"
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
         Left            =   4620
         MouseIcon       =   "FormMainWindow.frx":291A
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton CmdPresetMaster 
         Caption         =   "マスター"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3360
         MouseIcon       =   "FormMainWindow.frx":2A6C
         MousePointer    =   99  'Custom
         TabIndex        =   11
         ToolTipText     =   "Dungeon Master?"
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton CmdPresetHard 
         Caption         =   "困难"
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
         Left            =   2310
         MouseIcon       =   "FormMainWindow.frx":2BBE
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton CmdPresetNormal 
         Caption         =   "普通"
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
         Left            =   1260
         MouseIcon       =   "FormMainWindow.frx":2D10
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton CmdPresetEasy 
         Caption         =   "简单"
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
         MouseIcon       =   "FormMainWindow.frx":2E62
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   420
         Width           =   1065
      End
   End
   Begin VB.Timer TimerSettingsRefresher 
      Interval        =   100
      Left            =   14700
      Top             =   840
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "开始!"
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
      Left            =   12600
      MouseIcon       =   "FormMainWindow.frx":2FB4
      MousePointer    =   99  'Custom
      TabIndex        =   80
      Top             =   8925
      Width           =   2535
   End
   Begin VB.CommandButton CmdSoundSwitch 
      Caption         =   "音效: 关"
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
      Left            =   12810
      MouseIcon       =   "FormMainWindow.frx":3106
      MousePointer    =   99  'Custom
      TabIndex        =   82
      ToolTipText     =   "注意：开启音效可能会影响流畅度。"
      Top             =   210
      Width           =   1065
   End
   Begin VB.CommandButton CmdEXIT 
      Cancel          =   -1  'True
      Caption         =   "退出"
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
      Left            =   14070
      MouseIcon       =   "FormMainWindow.frx":3258
      MousePointer    =   99  'Custom
      TabIndex        =   83
      ToolTipText     =   "See you next time~"
      Top             =   210
      Width           =   1065
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12180
      MouseIcon       =   "FormMainWindow.frx":33AA
      MousePointer    =   99  'Custom
      TabIndex        =   81
      ToolTipText     =   "帮助"
      Top             =   210
      Width           =   435
   End
   Begin VB.Frame FrameFonts 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "字体"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1800
      Left            =   7560
      TabIndex        =   73
      Top             =   8505
      Width           =   4005
      Begin VB.CommandButton CmdFontsApply 
         Caption         =   "应用"
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
         Left            =   2835
         MouseIcon       =   "FormMainWindow.frx":34FC
         MousePointer    =   99  'Custom
         TabIndex        =   79
         Top             =   315
         Width           =   960
      End
      Begin VB.TextBox TextboxFontsRomajiFont 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1575
         MousePointer    =   3  'I-Beam
         TabIndex        =   78
         Text            =   "Microsoft Sans Serif"
         ToolTipText     =   "推荐：Microsoft Sans Serif，筑紫，思源。"
         Top             =   1260
         Width           =   2220
      End
      Begin VB.TextBox TextboxFontsKanaFont 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1575
         MousePointer    =   3  'I-Beam
         TabIndex        =   76
         Text            =   "MS PGothic"
         ToolTipText     =   "推荐：MS PGothic，MS PMincho，筑紫，思源，冬青，Shin-Go，教科书体。"
         Top             =   840
         Width           =   2220
      End
      Begin VB.CheckBox CheckboxFontsSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "自定义字体"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":364E
         MousePointer    =   99  'Custom
         TabIndex        =   74
         ToolTipText     =   "注意：设定不兼容的字体可能导致乱码。"
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label LabelFontsRomajiFont 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "罗马字字体:"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   315
         TabIndex        =   77
         Top             =   1260
         Width           =   1125
      End
      Begin VB.Label LabelFontsKanaFont 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "假名字体:"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   315
         TabIndex        =   75
         Top             =   840
         Width           =   1125
      End
   End
   Begin VB.Frame FrameKanaRange 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "范围"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1275
      Left            =   210
      TabIndex        =   16
      Top             =   4095
      Width           =   5895
      Begin VB.CheckBox CheckboxKanaRange11 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "生僻字: ゐゑヰヱ"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3885
         MouseIcon       =   "FormMainWindow.frx":37A0
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   735
         Width           =   1800
      End
      Begin VB.CheckBox CheckboxKanaRange10 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "オ"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormMainWindow.frx":38F2
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   735
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox CheckboxKanaRange09 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "エ"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2415
         MouseIcon       =   "FormMainWindow.frx":3A44
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   735
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox CheckboxKanaRange08 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "ウ"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1680
         MouseIcon       =   "FormMainWindow.frx":3B96
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   735
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox CheckboxKanaRange07 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "イ"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   945
         MouseIcon       =   "FormMainWindow.frx":3CE8
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   735
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox CheckboxKanaRange06 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "ア"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":3E3A
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   735
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox CheckboxKanaRange05 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "お"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3150
         MouseIcon       =   "FormMainWindow.frx":3F8C
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   315
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox CheckboxKanaRange04 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "え"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2415
         MouseIcon       =   "FormMainWindow.frx":40DE
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   315
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox CheckboxKanaRange03 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "う"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1680
         MouseIcon       =   "FormMainWindow.frx":4230
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   315
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox CheckboxKanaRange02 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "い"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   945
         MouseIcon       =   "FormMainWindow.frx":4382
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   315
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox CheckboxKanaRange01 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "あ"
         BeginProperty Font 
            Name            =   "MS PGothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":44D4
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   315
         Value           =   1  'Checked
         Width           =   645
      End
   End
   Begin VB.Frame FrameGameMode 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   210
      TabIndex        =   1
      Top             =   945
      Width           =   9045
      Begin VB.CommandButton CmdGameModeRomajisu 
         Caption         =   "Romaji-su!"
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
         Left            =   6825
         MouseIcon       =   "FormMainWindow.frx":4626
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "在 RomajiMaster 模式的基础上，按钮出现在随机位置，且只能用鼠标点击。"
         Top             =   420
         Width           =   1980
      End
      Begin VB.CommandButton CmdGameModeKanasu 
         Caption         =   "Kana-su!"
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
         Left            =   4620
         MouseIcon       =   "FormMainWindow.frx":4778
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "在 KanaMaster 模式的基础上，按钮出现在随机位置，且只能用鼠标点击。"
         Top             =   420
         Width           =   1980
      End
      Begin VB.CommandButton CmdGameModeRomajiMaster 
         Caption         =   "RomajiMaster"
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
         Left            =   2415
         MouseIcon       =   "FormMainWindow.frx":48CA
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "与 KanaMaster 模式相反，玩家须正确选择与罗马字对应的假名。"
         Top             =   420
         Width           =   1980
      End
      Begin VB.CommandButton CmdGameModeKanaMaster 
         Caption         =   "KanaMaster"
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
         MouseIcon       =   "FormMainWindow.frx":4A1C
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "玩家须正确选择与假名对应的罗马字。"
         Top             =   420
         Width           =   1980
      End
      Begin VB.Shape ShapeLightGameModeRomajisu 
         BackColor       =   &H00707070&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFD030&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF9000&
         Height          =   525
         Left            =   6780
         Top             =   375
         Width           =   2070
      End
      Begin VB.Shape ShapeLightGameModeKanasu 
         BackColor       =   &H00707070&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFD030&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF9000&
         Height          =   525
         Left            =   4575
         Top             =   375
         Width           =   2070
      End
      Begin VB.Shape ShapeLightGameModeRomajiMaster 
         BackColor       =   &H00707070&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFD030&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FF9000&
         Height          =   525
         Left            =   2370
         Top             =   375
         Width           =   2070
      End
      Begin VB.Shape ShapeLightGameModeKanaMaster 
         BackColor       =   &H00707070&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFD030&
         FillColor       =   &H00FF9000&
         FillStyle       =   0  'Solid
         Height          =   525
         Left            =   165
         Top             =   375
         Width           =   2070
      End
   End
   Begin VB.Frame FrameKeyboard 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "键位"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1800
      Left            =   210
      TabIndex        =   60
      Top             =   8505
      Width           =   4005
      Begin VB.TextBox TextboxKeyboardOption3 
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
         Left            =   3255
         MaxLength       =   1
         MouseIcon       =   "FormMainWindow.frx":4B6E
         MousePointer    =   99  'Custom
         TabIndex        =   66
         ToolTipText     =   "指定右边选项的键位。"
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox TextboxKeyboardOption2 
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
         Left            =   1995
         MaxLength       =   1
         MouseIcon       =   "FormMainWindow.frx":4CC0
         MousePointer    =   99  'Custom
         TabIndex        =   64
         ToolTipText     =   "指定中间选项的键位。"
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox TextboxKeyboardOption1 
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
         Left            =   735
         MaxLength       =   1
         MouseIcon       =   "FormMainWindow.frx":4E12
         MousePointer    =   99  'Custom
         TabIndex        =   62
         ToolTipText     =   "指定左边选项的键位。"
         Top             =   420
         Width           =   435
      End
      Begin VB.Label LabelKeyboardOption3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   2730
         TabIndex        =   65
         Top             =   420
         Width           =   495
      End
      Begin VB.Label LabelKeyboardOption2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1470
         TabIndex        =   63
         Top             =   420
         Width           =   495
      End
      Begin VB.Label LabelKeyboardOption1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   210
         TabIndex        =   61
         Top             =   420
         Width           =   495
      End
      Begin VB.Label LabelKeyboardNote 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "提示：在前两种模式下，您始终可以使用 F6/F7/F8；后两种模式只能用鼠标操作。"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   630
         Left            =   210
         TabIndex        =   67
         Top             =   945
         Width           =   3600
      End
   End
   Begin VB.Frame FrameDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "显示"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1800
      Left            =   4410
      TabIndex        =   68
      Top             =   8505
      Width           =   2955
      Begin VB.CheckBox CheckboxDisplaySpinningSakura 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "旋转樱花"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1575
         MouseIcon       =   "FormMainWindow.frx":4F64
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   1155
         Value           =   1  'Checked
         Width           =   1170
      End
      Begin VB.CheckBox CheckboxDisplaySmoothAnimations 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "平滑动画"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":50B6
         MousePointer    =   99  'Custom
         TabIndex        =   71
         ToolTipText     =   "不勾选此选项将会关闭窗口缩放动画、提示信息缩放动画、进度条平滑动画、数字跳动效果。"
         Top             =   1155
         Value           =   1  'Checked
         Width           =   1170
      End
      Begin VB.CheckBox CheckboxDisplayHideUnnecessaryInformation 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "隐藏部分信息"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":5208
         MousePointer    =   99  'Custom
         TabIndex        =   70
         ToolTipText     =   "隐藏出题板周围的大部分数字信息，例如倒计时与准确度。"
         Top             =   735
         Width           =   1590
      End
      Begin VB.CheckBox CheckboxDisplayReduceContrast 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "降低出题板对比度"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":535A
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   315
         Width           =   2010
      End
   End
   Begin VB.Frame FrameDifficulty 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "时限"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3480
      Left            =   6300
      TabIndex        =   37
      Top             =   4095
      Width           =   6105
      Begin VB.CheckBox CheckboxDifficultyIncreaseDifficultyGradually 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "逐渐缩短时限"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":54AC
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   945
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.HScrollBar HScrollDifficultyMistakeHPDrain 
         Height          =   330
         LargeChange     =   45
         Left            =   2310
         Max             =   50
         Min             =   5
         MouseIcon       =   "FormMainWindow.frx":55FE
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   2730
         Value           =   10
         Width           =   3585
      End
      Begin VB.HScrollBar HScrollDifficultyCooldown 
         Height          =   330
         LargeChange     =   18
         Left            =   2310
         Max             =   20
         Min             =   2
         MouseIcon       =   "FormMainWindow.frx":5750
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   2100
         Value           =   10
         Width           =   3585
      End
      Begin VB.HScrollBar HScrollDifficultyReachNormalDifficultyAt 
         Height          =   330
         LargeChange     =   95
         Left            =   3990
         Max             =   100
         Min             =   5
         MouseIcon       =   "FormMainWindow.frx":58A2
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   1470
         Value           =   50
         Width           =   1905
      End
      Begin VB.HScrollBar HScrollDifficultyInitialDifficulty 
         Height          =   330
         LargeChange     =   48
         Left            =   3990
         Max             =   50
         Min             =   2
         MouseIcon       =   "FormMainWindow.frx":59F4
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   1050
         Value           =   30
         Width           =   1905
      End
      Begin VB.HScrollBar HScrollDifficultyNormalDifficulty 
         Height          =   330
         LargeChange     =   28
         Left            =   1260
         Max             =   30
         Min             =   2
         MouseIcon       =   "FormMainWindow.frx":5B46
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   420
         Value           =   20
         Width           =   4635
      End
      Begin VB.Label LabelDifficultyMistakeHPDrain 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "失误扣血:"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":5C98
         MousePointer    =   99  'Custom
         TabIndex        =   50
         ToolTipText     =   "(HP Drain)"
         Top             =   2730
         Width           =   975
      End
      Begin VB.Label LabelDifficultyCooldown 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "冷却:"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":5DEA
         MousePointer    =   99  'Custom
         TabIndex        =   47
         ToolTipText     =   "(CD)"
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label LabelDifficultyReachNormalDifficultyAt 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "完成至此进度:"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1470
         TabIndex        =   44
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label LabelDifficultyInitialDifficulty 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "初始时限:"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1890
         TabIndex        =   41
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label LabelDifficultyNormalDifficultyIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   210
         TabIndex        =   38
         Top             =   420
         Width           =   915
      End
      Begin VB.Label LabelDifficultyInitialDifficultyIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2940
         TabIndex        =   42
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label LabelDifficultyReachNormalDifficultyAtIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2940
         TabIndex        =   45
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label LabelDifficultyCooldownIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1260
         TabIndex        =   48
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label LabelDifficultyMistakeHPDrainIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1260
         TabIndex        =   51
         Top             =   2730
         Width           =   915
      End
   End
   Begin VB.Frame FrameMods 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Mods"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   12600
      TabIndex        =   53
      Top             =   4095
      Width           =   2535
      Begin VB.CheckBox CheckboxModsPF 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "PF: 完美主义"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":5F3C
         MousePointer    =   99  'Custom
         TabIndex        =   55
         ToolTipText     =   "(Perfect)"
         Top             =   735
         Width           =   2220
      End
      Begin VB.CheckBox CheckboxModsNF 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "NF: 不死之身"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":608E
         MousePointer    =   99  'Custom
         TabIndex        =   56
         ToolTipText     =   "(No-Fail)"
         Top             =   1155
         Width           =   2220
      End
      Begin VB.CheckBox CheckboxModsAU 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "AU: 自动演示"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":61E0
         MousePointer    =   99  'Custom
         TabIndex        =   58
         ToolTipText     =   "(Auto)"
         Top             =   1995
         Width           =   2220
      End
      Begin VB.CheckBox CheckboxModsSD 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "SD: 失误即死"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":6332
         MousePointer    =   99  'Custom
         TabIndex        =   54
         ToolTipText     =   "(Sudden Death)"
         Top             =   315
         Width           =   2220
      End
      Begin VB.CheckBox CheckboxModsAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "AP: 指示答案"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":6484
         MousePointer    =   99  'Custom
         TabIndex        =   57
         ToolTipText     =   "(Autopilot)"
         Top             =   1575
         Width           =   2220
      End
   End
   Begin VB.Frame FrameProgressMode 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "进度"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2115
      Left            =   210
      TabIndex        =   28
      Top             =   5460
      Width           =   5895
      Begin VB.OptionButton RadioboxProgressModeKana 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "遍历范围内所有假名（抽选过且答对过所有假名）"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":65D6
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   315
         Value           =   -1  'True
         Width           =   4740
      End
      Begin VB.OptionButton RadioboxProgressModeTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0D0D0&
         Caption         =   "固定时长"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   210
         MouseIcon       =   "FormMainWindow.frx":6728
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1155
         Width           =   1275
      End
      Begin VB.HScrollBar HScrollProgressModeSpecifiedTime 
         Height          =   330
         LargeChange     =   29
         Left            =   2835
         Max             =   30
         Min             =   1
         MouseIcon       =   "FormMainWindow.frx":687A
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   1575
         Value           =   3
         Width           =   2850
      End
      Begin VB.HScrollBar HScrollProgressModeRepeatedTimes 
         Height          =   330
         LargeChange     =   9
         Left            =   2835
         Max             =   10
         Min             =   1
         MouseIcon       =   "FormMainWindow.frx":69CC
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   735
         Value           =   1
         Width           =   2850
      End
      Begin VB.Label LabelProgressModeSpecifiedTime 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "设定时间:"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   525
         TabIndex        =   34
         Top             =   1575
         Width           =   915
      End
      Begin VB.Label LabelProgressModeRepeatedTimes 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "重复次数:"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   525
         TabIndex        =   31
         Top             =   735
         Width           =   915
      End
      Begin VB.Label LabelProgressModeRepeatedTimesIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1575
         TabIndex        =   32
         Top             =   735
         Width           =   1125
      End
      Begin VB.Label LabelProgressModeSpecifiedTimeIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "Microsoft YaHei UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1575
         TabIndex        =   35
         Top             =   1575
         Width           =   1125
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   435
      Left            =   14070
      TabIndex        =   84
      Top             =   840
      Visible         =   0   'False
      Width           =   435
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   767
      _cy             =   767
   End
   Begin VB.Shape ShapeLightSoundSwitch 
      BackColor       =   &H00707070&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFD030&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF9000&
      Height          =   525
      Left            =   12765
      Top             =   165
      Width           =   1155
   End
   Begin VB.Label LabelTitle3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "偏好设定"
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
      Height          =   645
      Left            =   210
      TabIndex        =   59
      Top             =   7770
      Width           =   2220
   End
   Begin VB.Label LabelTitle2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "调整难度"
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
      Height          =   645
      Left            =   210
      TabIndex        =   6
      Top             =   2205
      Width           =   2220
   End
   Begin VB.Label LabelTitle1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "选择模式"
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
      Height          =   645
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   2220
   End
   Begin VB.Menu Menu 
      Caption         =   "菜单"
      Begin VB.Menu MenuStart 
         Caption         =   "开始!"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Menu2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuHelp 
         Caption         =   "帮助"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenuSoundSwitch 
         Caption         =   "音效"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuEXIT 
         Caption         =   "退出"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Menu1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "关于"
         Begin VB.Menu MenuAboutGitHub 
            Caption         =   "GitHub @SamToki"
         End
         Begin VB.Menu MenuAboutLicense 
            Caption         =   "Released under license GNU GPL v3"
         End
         Begin VB.Menu MenuAboutCopyright 
            Caption         =   "TM && (C) 2015-2021 SAM TOKI STUDIO"
         End
      End
   End
End
Attribute VB_Name = "FormMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === INFORMATION ===
'
'  SAM TOKI STUDIO
'  This is a .frm source code file.
'
'  KanaMaster
'
'  Powered by Sam Toki
'  Version: v0.52chs
'  Date:    2021/08/25 (Wed)
'  History: First version v0.10 was built on 2020/03/18.
'
'  WARNING: Commercial use of this computer software is strictly prohibited.
'           Open source license:      GNU GPL v3
'           Creative Commons license: CC BY-NC 3.0
'
'  Copyright: TM & (C) 2015-2021 SAM TOKI STUDIO. All rights reserved.
'             SAM TOKI STUDIO is a trademark of CZJ Software Technologies (CZJST) Inc. in the P.R.C and other countries.
'
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === NOTES FOR REFERENCE ===
'
'  ...
'
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Option Explicit

'Declare Menu...
Public soundswitch As Boolean

'Declare Settings...
Public setgamemode As Integer  '1-KanaMaster, 2-RomajiMaster, 3-Kanasu, 4-Romajisu.

Public setkanarangeswitch As Variant  '(1 to 11)
Public settotalquestion As Integer

Public setprogressmode As Integer  '1-Kana, 2-Time.
Public setrepeatedtimes As Integer  'Range: 1~10.
Public setspecifiedtime As Integer  'Unit: min. Range: 1~30min.

Public setnormaldifficulty As Single  'Unit: sec. Range: 0.2~5.0
Public setincreasedifficultygraduallyswitch As Boolean
Public setinitialdifficulty As Single  'Unit: sec. Range: 0.2~5.0
Public setreachnormaldifficultyat As Double  'Unit: %. Range: 5.00~100.00
Public setcooldown As Single  'Unit: sec. Range: 0.2~2.0
Public setmistakehpdrain As Single  'Range: 5.0~50.0

Public setmodsdswitch As Boolean
Public setmodpfswitch As Boolean
Public setmodnfswitch As Boolean
Public setmodapswitch As Boolean
Public setmodauswitch As Boolean

Public setkeyboardoption As Variant  '(1 to 3)

Public setreducecontrast As Boolean
Public sethideunnecessaryinfo As Boolean
Public setanimationswitch As Boolean
Public setspinningsakuraswitch As Boolean

Public setfontswitch As Boolean

'Declare Others...
Public forloop1 As Integer
Public forloop2 As Integer
Public forloop3 As Integer

'Declare Dialog...
Public answer

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Sub Form_Load()
        'Initialize Menu...
        soundswitch = False

        'Initialize Settings...
        setgamemode = 1

        setkanarangeswitch = Array("!!", True, True, True, True, True, True, True, True, True, True, False)

        setprogressmode = 1
        setrepeatedtimes = 1
        setspecifiedtime = 3

        setnormaldifficulty = 2
        setincreasedifficultygraduallyswitch = True
        setinitialdifficulty = 3
        setreachnormaldifficultyat = 50
        setcooldown = 1
        setmistakehpdrain = 1

        setmodsdswitch = False
        setmodpfswitch = False
        setmodnfswitch = False
        setmodapswitch = False
        setmodauswitch = False

        setkeyboardoption = Array("!!", "1", "2", "3")

        setreducecontrast = False
        sethideunnecessaryinfo = False
        setanimationswitch = True
        setspinningsakuraswitch = True

        setfontswitch = False
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    'CMD Menu...
    Public Sub MenuHelp_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_DialogOpen.wav"
        FormHelp.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormHelp.windowanimationtargetleft = (Screen.Width / 2) - (11040 / 2)
        FormHelp.windowanimationtargettop = (Screen.Height / 2) - (7995 / 2)
        FormHelp.windowanimationtargetwidth = 11040
        FormHelp.windowanimationtargetheight = 7995
        FormHelp.Show
    End Sub
    Public Sub CmdHelp_Click()
        Call MenuHelp_Click
    End Sub
    Public Sub MenuSoundSwitch_Click()
        Select Case soundswitch
            Case True
                soundswitch = False
                WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_SwitchOff.wav"
                MenuSoundSwitch.Checked = False
                CmdSoundSwitch.Caption = "音效: 关"
                ShapeLightSoundSwitch.BorderStyle = 0
                ShapeLightSoundSwitch.FillStyle = 1
            Case False
                soundswitch = True
                WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_SwitchOn.wav"
                MenuSoundSwitch.Checked = True
                CmdSoundSwitch.Caption = "音效: 开"
                ShapeLightSoundSwitch.BorderStyle = 1
                ShapeLightSoundSwitch.FillStyle = 0
        End Select
    End Sub
    Public Sub CmdSoundSwitch_Click()
        Call MenuSoundSwitch_Click
    End Sub
    Public Sub MenuEXIT_Click()
        End
    End Sub
    Public Sub CmdEXIT_Click()
        Call MenuEXIT_Click
    End Sub
    Public Sub MenuStart_Click()
        'DISABLED LINE: If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_KanaMasterStart.wav"
        Select Case setgamemode
            Case 1
                Me.Hide: FormKanaMaster.Show
                FormKanaMaster.gamestatus = 0
                Call FormKanaMaster.MenuStartPauseResume_Click
            'Case 2
                'Me.Hide: FormRomajiMaster.Show
                'FormRomajiMaster.gamestatus = 0
                'Call FormRomajiMaster.MenuStartPauseResume_Click
            'Case 3
                'Me.Hide: FormKanasu.Show
                'FormKanasu.gamestatus = 0
                'Call FormKanasu.MenuStartPauseResume_Click
            'Case 4
                'Me.Hide: FormRomajisu.Show
                'FormRomajisu.gamestatus = 0
                'Call FormRomajisu.MenuStartPauseResume_Click
            Case Else
                MsgBox "错误：Variable setgamemode is out of range." & vbCrLf & "您可以通过 GitHub @SamToki 提供反馈以帮助解决问题。", vbCritical + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End Select
    End Sub
    Public Sub CmdStart_Click()
        Call MenuStart_Click
    End Sub

    'CMD Game Mode...
    Public Sub CmdGameModeKanaMaster_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        setgamemode = 1
        ShapeLightGameModeKanaMaster.BorderStyle = 1
        ShapeLightGameModeKanaMaster.FillStyle = 0
        ShapeLightGameModeRomajiMaster.BorderStyle = 0
        ShapeLightGameModeRomajiMaster.FillStyle = 1
        ShapeLightGameModeKanasu.BorderStyle = 0
        ShapeLightGameModeKanasu.FillStyle = 1
        ShapeLightGameModeRomajisu.BorderStyle = 0
        ShapeLightGameModeRomajisu.FillStyle = 1
    End Sub
    Public Sub CmdGameModeRomajiMaster_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        setgamemode = 2
        ShapeLightGameModeKanaMaster.BorderStyle = 0
        ShapeLightGameModeKanaMaster.FillStyle = 1
        ShapeLightGameModeRomajiMaster.BorderStyle = 1
        ShapeLightGameModeRomajiMaster.FillStyle = 0
        ShapeLightGameModeKanasu.BorderStyle = 0
        ShapeLightGameModeKanasu.FillStyle = 1
        ShapeLightGameModeRomajisu.BorderStyle = 0
        ShapeLightGameModeRomajisu.FillStyle = 1
    End Sub
    Public Sub CmdGameModeKanasu_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        setgamemode = 3
        ShapeLightGameModeKanaMaster.BorderStyle = 0
        ShapeLightGameModeKanaMaster.FillStyle = 1
        ShapeLightGameModeRomajiMaster.BorderStyle = 0
        ShapeLightGameModeRomajiMaster.FillStyle = 1
        ShapeLightGameModeKanasu.BorderStyle = 1
        ShapeLightGameModeKanasu.FillStyle = 0
        ShapeLightGameModeRomajisu.BorderStyle = 0
        ShapeLightGameModeRomajisu.FillStyle = 1
    End Sub
    Public Sub CmdGameModeRomajisu_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        setgamemode = 4
        ShapeLightGameModeKanaMaster.BorderStyle = 0
        ShapeLightGameModeKanaMaster.FillStyle = 1
        ShapeLightGameModeRomajiMaster.BorderStyle = 0
        ShapeLightGameModeRomajiMaster.FillStyle = 1
        ShapeLightGameModeKanasu.BorderStyle = 0
        ShapeLightGameModeKanasu.FillStyle = 1
        ShapeLightGameModeRomajisu.BorderStyle = 1
        ShapeLightGameModeRomajisu.FillStyle = 0
    End Sub

    'CMD Preset...
    Public Sub CmdPresetEasy_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        CheckboxKanaRange01.Value = 1: CheckboxKanaRange02.Value = 1: CheckboxKanaRange03.Value = 1: CheckboxKanaRange04.Value = 1: CheckboxKanaRange05.Value = 1: CheckboxKanaRange06.Value = 0: CheckboxKanaRange07.Value = 0: CheckboxKanaRange08.Value = 0: CheckboxKanaRange09.Value = 0: CheckboxKanaRange10.Value = 0: CheckboxKanaRange11.Value = 0
        RadioboxProgressModeKana.Value = True: HScrollProgressModeRepeatedTimes.Value = 1
        HScrollDifficultyInitialDifficulty.Value = 50: HScrollDifficultyNormalDifficulty.Max = HScrollDifficultyInitialDifficulty.Value: HScrollDifficultyNormalDifficulty.Value = 30: CheckboxDifficultyIncreaseDifficultyGradually.Value = 1: HScrollDifficultyReachNormalDifficultyAt.Value = 50: HScrollDifficultyCooldown.Value = 15: HScrollDifficultyMistakeHPDrain.Value = 10
        CheckboxModsSD.Value = 0: CheckboxModsPF.Value = 0: CheckboxModsNF.Value = 0: CheckboxModsAP.Value = 0: CheckboxModsAU.Value = 0
    End Sub
    Public Sub CmdPresetNormal_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        CheckboxKanaRange01.Value = 1: CheckboxKanaRange02.Value = 1: CheckboxKanaRange03.Value = 1: CheckboxKanaRange04.Value = 1: CheckboxKanaRange05.Value = 1: CheckboxKanaRange06.Value = 1: CheckboxKanaRange07.Value = 1: CheckboxKanaRange08.Value = 1: CheckboxKanaRange09.Value = 1: CheckboxKanaRange10.Value = 1: CheckboxKanaRange11.Value = 0
        RadioboxProgressModeKana.Value = True: HScrollProgressModeRepeatedTimes.Value = 1
        HScrollDifficultyInitialDifficulty.Value = 30: HScrollDifficultyNormalDifficulty.Max = HScrollDifficultyInitialDifficulty.Value: HScrollDifficultyNormalDifficulty.Value = 20: CheckboxDifficultyIncreaseDifficultyGradually.Value = 1: HScrollDifficultyReachNormalDifficultyAt.Value = 50: HScrollDifficultyCooldown.Value = 10: HScrollDifficultyMistakeHPDrain.Value = 10
        CheckboxModsSD.Value = 0: CheckboxModsPF.Value = 0: CheckboxModsNF.Value = 0: CheckboxModsAP.Value = 0: CheckboxModsAU.Value = 0
    End Sub
    Public Sub CmdPresetHard_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        CheckboxKanaRange01.Value = 1: CheckboxKanaRange02.Value = 1: CheckboxKanaRange03.Value = 1: CheckboxKanaRange04.Value = 1: CheckboxKanaRange05.Value = 1: CheckboxKanaRange06.Value = 1: CheckboxKanaRange07.Value = 1: CheckboxKanaRange08.Value = 1: CheckboxKanaRange09.Value = 1: CheckboxKanaRange10.Value = 1: CheckboxKanaRange11.Value = 0
        RadioboxProgressModeKana.Value = True: HScrollProgressModeRepeatedTimes.Value = 2
        HScrollDifficultyInitialDifficulty.Value = 30: HScrollDifficultyNormalDifficulty.Max = HScrollDifficultyInitialDifficulty.Value: HScrollDifficultyNormalDifficulty.Value = 20: CheckboxDifficultyIncreaseDifficultyGradually.Value = 1: HScrollDifficultyReachNormalDifficultyAt.Value = 20: HScrollDifficultyCooldown.Value = 10: HScrollDifficultyMistakeHPDrain.Value = 20
        CheckboxModsSD.Value = 0: CheckboxModsPF.Value = 0: CheckboxModsNF.Value = 0: CheckboxModsAP.Value = 0: CheckboxModsAU.Value = 0
    End Sub
    Public Sub CmdPresetMaster_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        CheckboxKanaRange01.Value = 1: CheckboxKanaRange02.Value = 1: CheckboxKanaRange03.Value = 1: CheckboxKanaRange04.Value = 1: CheckboxKanaRange05.Value = 1: CheckboxKanaRange06.Value = 1: CheckboxKanaRange07.Value = 1: CheckboxKanaRange08.Value = 1: CheckboxKanaRange09.Value = 1: CheckboxKanaRange10.Value = 1: CheckboxKanaRange11.Value = 1
        RadioboxProgressModeKana.Value = True: HScrollProgressModeRepeatedTimes.Value = 5
        HScrollDifficultyInitialDifficulty.Value = 30: HScrollDifficultyNormalDifficulty.Max = HScrollDifficultyInitialDifficulty.Value: HScrollDifficultyNormalDifficulty.Value = 15: CheckboxDifficultyIncreaseDifficultyGradually.Value = 1: HScrollDifficultyReachNormalDifficultyAt.Value = 10: HScrollDifficultyCooldown.Value = 5: HScrollDifficultyMistakeHPDrain.Value = 30
        CheckboxModsSD.Value = 0: CheckboxModsPF.Value = 0: CheckboxModsNF.Value = 0: CheckboxModsAP.Value = 0: CheckboxModsAU.Value = 0
    End Sub
    Public Sub CmdPresetDefaultKanaRange_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        CheckboxKanaRange01.Value = 1: CheckboxKanaRange02.Value = 1: CheckboxKanaRange03.Value = 1: CheckboxKanaRange04.Value = 1: CheckboxKanaRange05.Value = 1: CheckboxKanaRange06.Value = 1: CheckboxKanaRange07.Value = 1: CheckboxKanaRange08.Value = 1: CheckboxKanaRange09.Value = 1: CheckboxKanaRange10.Value = 1: CheckboxKanaRange11.Value = 0
    End Sub
    Public Sub CmdPresetOnlyHiragana_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        CheckboxKanaRange01.Value = 1: CheckboxKanaRange02.Value = 1: CheckboxKanaRange03.Value = 1: CheckboxKanaRange04.Value = 1: CheckboxKanaRange05.Value = 1: CheckboxKanaRange06.Value = 0: CheckboxKanaRange07.Value = 0: CheckboxKanaRange08.Value = 0: CheckboxKanaRange09.Value = 0: CheckboxKanaRange10.Value = 0: CheckboxKanaRange11.Value = 0
    End Sub
    Public Sub CmdPresetOnlyKatakana_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        CheckboxKanaRange01.Value = 0: CheckboxKanaRange02.Value = 0: CheckboxKanaRange03.Value = 0: CheckboxKanaRange04.Value = 0: CheckboxKanaRange05.Value = 0: CheckboxKanaRange06.Value = 1: CheckboxKanaRange07.Value = 1: CheckboxKanaRange08.Value = 1: CheckboxKanaRange09.Value = 1: CheckboxKanaRange10.Value = 1: CheckboxKanaRange11.Value = 0
    End Sub
    Public Sub CmdPresetAllKana_Click()
        If soundswitch = True Then WindowsMediaPlayer1.URL = App.Path & "\CZJSTappdata\CZJSTaudio\CZJSTaudio_Tab.wav"
        CheckboxKanaRange01.Value = 1: CheckboxKanaRange02.Value = 1: CheckboxKanaRange03.Value = 1: CheckboxKanaRange04.Value = 1: CheckboxKanaRange05.Value = 1: CheckboxKanaRange06.Value = 1: CheckboxKanaRange07.Value = 1: CheckboxKanaRange08.Value = 1: CheckboxKanaRange09.Value = 1: CheckboxKanaRange10.Value = 1: CheckboxKanaRange11.Value = 1
    End Sub

    'Settings...
    Public Sub TimerSettingsRefresher_Timer()
        'Kana Range...
        If CheckboxKanaRange01.Value = 1 Then setkanarangeswitch(1) = True Else setkanarangeswitch(1) = False
        If CheckboxKanaRange02.Value = 1 Then setkanarangeswitch(2) = True Else setkanarangeswitch(2) = False
        If CheckboxKanaRange03.Value = 1 Then setkanarangeswitch(3) = True Else setkanarangeswitch(3) = False
        If CheckboxKanaRange04.Value = 1 Then setkanarangeswitch(4) = True Else setkanarangeswitch(4) = False
        If CheckboxKanaRange05.Value = 1 Then setkanarangeswitch(5) = True Else setkanarangeswitch(5) = False
        If CheckboxKanaRange06.Value = 1 Then setkanarangeswitch(6) = True Else setkanarangeswitch(6) = False
        If CheckboxKanaRange07.Value = 1 Then setkanarangeswitch(7) = True Else setkanarangeswitch(7) = False
        If CheckboxKanaRange08.Value = 1 Then setkanarangeswitch(8) = True Else setkanarangeswitch(8) = False
        If CheckboxKanaRange09.Value = 1 Then setkanarangeswitch(9) = True Else setkanarangeswitch(9) = False
        If CheckboxKanaRange10.Value = 1 Then setkanarangeswitch(10) = True Else setkanarangeswitch(10) = False
        If CheckboxKanaRange11.Value = 1 Then setkanarangeswitch(11) = True Else setkanarangeswitch(11) = False

        settotalquestion = 0
        If setkanarangeswitch(1) = True Then settotalquestion = settotalquestion + 16
        If setkanarangeswitch(2) = True Then settotalquestion = settotalquestion + 13
        If setkanarangeswitch(3) = True Then settotalquestion = settotalquestion + 14
        If setkanarangeswitch(4) = True Then settotalquestion = settotalquestion + 13
        If setkanarangeswitch(5) = True Then settotalquestion = settotalquestion + 15
        If setkanarangeswitch(6) = True Then settotalquestion = settotalquestion + 16
        If setkanarangeswitch(7) = True Then settotalquestion = settotalquestion + 13
        If setkanarangeswitch(8) = True Then settotalquestion = settotalquestion + 15
        If setkanarangeswitch(9) = True Then settotalquestion = settotalquestion + 13
        If setkanarangeswitch(10) = True Then settotalquestion = settotalquestion + 15
        If setkanarangeswitch(11) = True Then settotalquestion = settotalquestion + 4

        'Prevent disabling all kanarangeswitch...
        If settotalquestion = 0 Then
            MsgBox "注意：不可以排除所有假名。将恢复默认范围。", vbExclamation + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
            CheckboxKanaRange01.Value = 1: CheckboxKanaRange02.Value = 1: CheckboxKanaRange03.Value = 1: CheckboxKanaRange04.Value = 1: CheckboxKanaRange05.Value = 1: CheckboxKanaRange06.Value = 1: CheckboxKanaRange07.Value = 1: CheckboxKanaRange08.Value = 1: CheckboxKanaRange09.Value = 1: CheckboxKanaRange10.Value = 1: CheckboxKanaRange11.Value = 0
        End If

        'Progress Mode...
        If RadioboxProgressModeKana.Value = True Then
            setprogressmode = 1
            setrepeatedtimes = HScrollProgressModeRepeatedTimes.Value
        End If
        If RadioboxProgressModeTime.Value = True Then
            setprogressmode = 2
            setrepeatedtimes = 9999
        End If
        setspecifiedtime = HScrollProgressModeSpecifiedTime.Value
        LabelProgressModeRepeatedTimesIndicator.Caption = setrepeatedtimes
        LabelProgressModeSpecifiedTimeIndicator.Caption = setspecifiedtime & " 分钟"

        'Difficulty...
        HScrollDifficultyNormalDifficulty.Max = HScrollDifficultyInitialDifficulty.Value

        setnormaldifficulty = HScrollDifficultyNormalDifficulty.Value / 10
        LabelDifficultyNormalDifficultyIndicator.Caption = Format(setnormaldifficulty, "0.0") & " 秒"

        If CheckboxDifficultyIncreaseDifficultyGradually.Value = 1 Then setincreasedifficultygraduallyswitch = True Else setincreasedifficultygraduallyswitch = False
        setinitialdifficulty = HScrollDifficultyInitialDifficulty.Value / 10
        LabelDifficultyInitialDifficultyIndicator.Caption = Format(setinitialdifficulty, "0.0") & " 秒"
        setreachnormaldifficultyat = HScrollDifficultyReachNormalDifficultyAt.Value
        LabelDifficultyReachNormalDifficultyAtIndicator.Caption = setreachnormaldifficultyat & "%"

        setcooldown = HScrollDifficultyCooldown.Value / 10
        LabelDifficultyCooldownIndicator.Caption = Format(setcooldown, "0.0") & " 秒"
        setmistakehpdrain = HScrollDifficultyMistakeHPDrain.Value
        LabelDifficultyMistakeHPDrainIndicator.Caption = Format(setmistakehpdrain, "0.0")

        'Mods...
        'Conflict resolver...
        If (CheckboxModsSD.Value = 1 And CheckboxModsPF.Value = 1) Or (CheckboxModsSD.Value = 1 And CheckboxModsNF.Value = 1) Or (CheckboxModsPF.Value = 1 And CheckboxModsNF.Value = 1) Then
            CheckboxModsSD.Value = 0
            CheckboxModsPF.Value = 0
            CheckboxModsNF.Value = 0
        End If
        If CheckboxModsSD.Value = 1 Then
            setmodsdswitch = True
            FormKanaMaster.LabelLightIndicatorModSD.BackColor = &HFF9000
            FormKanaMaster.LabelLightIndicatorModSD.ForeColor = &HFFFFFF
            '?????
        Else
            setmodsdswitch = False
            FormKanaMaster.LabelLightIndicatorModSD.BackColor = &H707070
            FormKanaMaster.LabelLightIndicatorModSD.ForeColor = &HB0B0B0
            '?????
        End If
        If CheckboxModsPF.Value = 1 Then
            setmodpfswitch = True
            FormKanaMaster.LabelLightIndicatorModPF.BackColor = &HFF9000
            FormKanaMaster.LabelLightIndicatorModPF.ForeColor = &HFFFFFF
            '?????
        Else
            setmodpfswitch = False
            FormKanaMaster.LabelLightIndicatorModPF.BackColor = &H707070
            FormKanaMaster.LabelLightIndicatorModPF.ForeColor = &HB0B0B0
            '?????
        End If
        If CheckboxModsNF.Value = 1 Then
            setmodnfswitch = True
            FormKanaMaster.LabelLightIndicatorModNF.BackColor = &HFF9000
            FormKanaMaster.LabelLightIndicatorModNF.ForeColor = &HFFFFFF
            '?????
        Else
            setmodnfswitch = False
            FormKanaMaster.LabelLightIndicatorModNF.BackColor = &H707070
            FormKanaMaster.LabelLightIndicatorModNF.ForeColor = &HB0B0B0
            '?????
        End If
        If CheckboxModsAP.Value = 1 Then
            setmodapswitch = True
            FormKanaMaster.LabelLightIndicatorModAP.BackColor = &HFF9000
            FormKanaMaster.LabelLightIndicatorModAP.ForeColor = &HFFFFFF
            '?????
        Else
            setmodapswitch = False
            FormKanaMaster.LabelLightIndicatorModAP.BackColor = &H707070
            FormKanaMaster.LabelLightIndicatorModAP.ForeColor = &HB0B0B0
            '?????
        End If
        If CheckboxModsAU.Value = 1 Then
            setmodauswitch = True
            FormKanaMaster.LabelLightIndicatorModAU.BackColor = &HFF9000
            FormKanaMaster.LabelLightIndicatorModAU.ForeColor = &HFFFFFF
            '?????
        Else
            setmodauswitch = False
            FormKanaMaster.LabelLightIndicatorModAU.BackColor = &H707070
            FormKanaMaster.LabelLightIndicatorModAU.ForeColor = &HB0B0B0
            '?????
        End If

        'Input...
        If TextboxKeyboardOption1.Text <> "" Then
            setkeyboardoption(1) = TextboxKeyboardOption1.Text: LabelKeyboardOption1.Caption = setkeyboardoption(1)
            FormKanaMaster.LabelOption1.Caption = setkeyboardoption(1)
            '?????
            TextboxKeyboardOption1.Text = ""
            If (setkeyboardoption(1) = setkeyboardoption(2)) Or (setkeyboardoption(1) = setkeyboardoption(3)) Or (setkeyboardoption(2) = setkeyboardoption(3)) Then MsgBox "注意：键位冲突。", vbExclamation + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End If
        If TextboxKeyboardOption2.Text <> "" Then
            setkeyboardoption(2) = TextboxKeyboardOption2.Text: LabelKeyboardOption2.Caption = setkeyboardoption(2)
            FormKanaMaster.LabelOption2.Caption = setkeyboardoption(2)
            '?????
            TextboxKeyboardOption2.Text = ""
            If (setkeyboardoption(1) = setkeyboardoption(2)) Or (setkeyboardoption(1) = setkeyboardoption(3)) Or (setkeyboardoption(2) = setkeyboardoption(3)) Then MsgBox "注意：键位冲突。", vbExclamation + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End If
        If TextboxKeyboardOption3.Text <> "" Then
            setkeyboardoption(3) = TextboxKeyboardOption3.Text: LabelKeyboardOption3.Caption = setkeyboardoption(3)
            FormKanaMaster.LabelOption3.Caption = setkeyboardoption(3)
            '?????
            TextboxKeyboardOption3.Text = ""
            If (setkeyboardoption(1) = setkeyboardoption(2)) Or (setkeyboardoption(1) = setkeyboardoption(3)) Or (setkeyboardoption(2) = setkeyboardoption(3)) Then MsgBox "注意：键位冲突。", vbExclamation + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End If

        'Display...
        If CheckboxDisplayReduceContrast.Value = 1 Then
            setreducecontrast = True
            FormKanaMaster.LabelDashboard.BackColor = &HB0B0B0
            'FormRomajiMaster.LabelDashboard.BackColor = &HB0B0B0
            'FormKanasu.LabelDashboard.ForeColor = &H707070
            'FormRomajisu.LabelDashboard.ForeColor = &H707070
        Else
            setreducecontrast = False
            FormKanaMaster.LabelDashboard.BackColor = &HFFFFFF
            'FormRomajiMaster.LabelDashboard.BackColor = &HFFFFFF
            'FormKanasu.LabelDashboard.ForeColor = &HB0B0B0
            'FormRomajisu.LabelDashboard.ForeColor = &HB0B0B0
        End If

        If CheckboxDisplayHideUnnecessaryInformation.Value = 1 Then
            sethideunnecessaryinfo = True
            FormKanaMaster.LabelTimeElapsed.Visible = False: FormKanaMaster.LabelTotalCount.Visible = False: FormKanaMaster.LabelProgress.Visible = False: FormKanaMaster.LabelScore.Visible = False
            FormKanaMaster.LabelHPTitle.Visible = False: FormKanaMaster.LabelHP.Visible = False
            FormKanaMaster.LabelComboCountTitle.Visible = False: FormKanaMaster.LabelComboCount.Visible = False: FormKanaMaster.LabelBestComboCount.Visible = False
            FormKanaMaster.LabelMissCountTitle.Visible = False: FormKanaMaster.LabelMissCount.Visible = False
            FormKanaMaster.LabelTimeLeftTitle.Visible = False: FormKanaMaster.LabelTimeLeft.Visible = False: FormKanaMaster.LabelCurrentDifficulty.Visible = False
            FormKanaMaster.LabelAverageReactionTimeTitle.Visible = False: FormKanaMaster.LabelAverageReactionTime.Visible = False
            FormKanaMaster.LabelAccuracyTitle.Visible = False: FormKanaMaster.LabelAccuracy.Visible = False
            FormKanaMaster.LabelOption1.Visible = False: FormKanaMaster.LabelOption2.Visible = False: FormKanaMaster.LabelOption3.Visible = False
            '?????
        Else
            sethideunnecessaryinfo = False
            FormKanaMaster.LabelTimeElapsed.Visible = True: FormKanaMaster.LabelTotalCount.Visible = True: FormKanaMaster.LabelProgress.Visible = True: FormKanaMaster.LabelScore.Visible = True
            FormKanaMaster.LabelHPTitle.Visible = True: FormKanaMaster.LabelHP.Visible = True
            FormKanaMaster.LabelComboCountTitle.Visible = True: FormKanaMaster.LabelComboCount.Visible = True: FormKanaMaster.LabelBestComboCount.Visible = True
            FormKanaMaster.LabelMissCountTitle.Visible = True: FormKanaMaster.LabelMissCount.Visible = True
            FormKanaMaster.LabelTimeLeftTitle.Visible = True: FormKanaMaster.LabelTimeLeft.Visible = True: FormKanaMaster.LabelCurrentDifficulty.Visible = True
            FormKanaMaster.LabelAverageReactionTimeTitle.Visible = True: FormKanaMaster.LabelAverageReactionTime.Visible = True
            FormKanaMaster.LabelAccuracyTitle.Visible = True: FormKanaMaster.LabelAccuracy.Visible = True
            FormKanaMaster.LabelOption1.Visible = True: FormKanaMaster.LabelOption2.Visible = True: FormKanaMaster.LabelOption3.Visible = True
            '?????
        End If

        If CheckboxDisplaySmoothAnimations.Value = 1 Then setanimationswitch = True Else setanimationswitch = False

        If CheckboxDisplaySpinningSakura.Value = 1 Then
            setspinningsakuraswitch = True
            FormKanaMaster.LineSpinningSakura1.Visible = True: FormKanaMaster.LineSpinningSakura2.Visible = True: FormKanaMaster.LineSpinningSakura3.Visible = True: FormKanaMaster.LineSpinningSakura4.Visible = True: FormKanaMaster.LineSpinningSakura5.Visible = True
            'FormRomajiMaster.LineSpinningSakura1.Visible = True: FormRomajiMaster.LineSpinningSakura2.Visible = True: FormRomajiMaster.LineSpinningSakura3.Visible = True: FormRomajiMaster.LineSpinningSakura4.Visible = True: FormRomajiMaster.LineSpinningSakura5.Visible = True
            'FormKanasu.LineSpinningSakura1.Visible = True: FormKanasu.LineSpinningSakura2.Visible = True: FormKanasu.LineSpinningSakura3.Visible = True: FormKanasu.LineSpinningSakura4.Visible = True: FormKanasu.LineSpinningSakura5.Visible = True
            'FormRomajisu.LineSpinningSakura1.Visible = True: FormRomajisu.LineSpinningSakura2.Visible = True: FormRomajisu.LineSpinningSakura3.Visible = True: FormRomajisu.LineSpinningSakura4.Visible = True: FormRomajisu.LineSpinningSakura5.Visible = True
        Else
            setspinningsakuraswitch = False
            FormKanaMaster.LineSpinningSakura1.Visible = False: FormKanaMaster.LineSpinningSakura2.Visible = False: FormKanaMaster.LineSpinningSakura3.Visible = False: FormKanaMaster.LineSpinningSakura4.Visible = False: FormKanaMaster.LineSpinningSakura5.Visible = False
            'FormRomajiMaster.LineSpinningSakura1.Visible = False: FormRomajiMaster.LineSpinningSakura2.Visible = False: FormRomajiMaster.LineSpinningSakura3.Visible = False: FormRomajiMaster.LineSpinningSakura4.Visible = False: FormRomajiMaster.LineSpinningSakura5.Visible = False
            'FormKanasu.LineSpinningSakura1.Visible = False: FormKanasu.LineSpinningSakura2.Visible = False: FormKanasu.LineSpinningSakura3.Visible = False: FormKanasu.LineSpinningSakura4.Visible = False: FormKanasu.LineSpinningSakura5.Visible = False
            'FormRomajisu.LineSpinningSakura1.Visible = False: FormRomajisu.LineSpinningSakura2.Visible = False: FormRomajisu.LineSpinningSakura3.Visible = False: FormRomajisu.LineSpinningSakura4.Visible = False: FormRomajisu.LineSpinningSakura5.Visible = False
        End If

        'Fonts...
        If CheckboxFontsSwitch.Value = 1 Then
            TextboxFontsKanaFont.Enabled = True: TextboxFontsRomajiFont.Enabled = True: CmdFontsApply.Enabled = True
        Else
            TextboxFontsKanaFont.Enabled = False: TextboxFontsRomajiFont.Enabled = False: CmdFontsApply.Enabled = False
            FormKanaMaster.LabelDashboard.Font = "MS PGothic": FormKanaMaster.CmdOption1.Font = "Microsoft Sans Serif": FormKanaMaster.CmdOption2.Font = "Microsoft Sans Serif": FormKanaMaster.CmdOption3.Font = "Microsoft Sans Serif"
            'FormRomajiMaster.LabelDashboard.Font = "Microsoft Sans Serif": FormRomajiMaster.CmdOption1.Font = "MS PGothic": FormRomajiMaster.CmdOption2.Font = "MS PGothic": FormRomajiMaster.CmdOption3.Font = "MS PGothic"
            'FormKanasu.LabelDashboard.Font = "MS PGothic": FormKanasu.CmdOption1.Font = "Microsoft Sans Serif": FormKanasu.CmdOption2.Font = "Microsoft Sans Serif": FormKanasu.CmdOption3.Font = "Microsoft Sans Serif"
            'FormRomajisu.LabelDashboard.Font = "Microsoft Sans Serif": FormRomajisu.CmdOption1.Font = "MS PGothic": FormRomajisu.CmdOption2.Font = "MS PGothic": FormRomajisu.CmdOption3.Font = "MS PGothic"
            'DISABLED LINE: MsgBox "已恢复默认字体！", vbInformation + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
        End If
    End Sub

    Public Sub CmdFontsApply_Click()
        FormKanaMaster.LabelDashboard.Font = TextboxFontsKanaFont.Text: FormKanaMaster.CmdOption1.Font = TextboxFontsRomajiFont.Text: FormKanaMaster.CmdOption2.Font = TextboxFontsRomajiFont.Text: FormKanaMaster.CmdOption3.Font = TextboxFontsRomajiFont.Text
        'FormRomajiMaster.LabelDashboard.Font = TextboxFontsRomajiFont.Text: FormRomajiMaster.CmdOption1.Font = TextboxFontsKanaFont.Text: FormRomajiMaster.CmdOption2.Font = TextboxFontsKanaFont.Text: FormRomajiMaster.CmdOption3.Font = TextboxFontsKanaFont.Text
        'FormKanasu.LabelDashboard.Font = TextboxFontsKanaFont.Text: FormKanasu.CmdOption1.Font = TextboxFontsRomajiFont.Text: FormKanasu.CmdOption2.Font = TextboxFontsRomajiFont.Text: FormKanasu.CmdOption3.Font = TextboxFontsRomajiFont.Text
        'FormRomajisu.LabelDashboard.Font = TextboxFontsRomajiFont.Text: FormRomajisu.CmdOption1.Font = TextboxFontsKanaFont.Text: FormRomajisu.CmdOption2.Font = TextboxFontsKanaFont.Text: FormRomajisu.CmdOption3.Font = TextboxFontsKanaFont.Text
        MsgBox "字体已更换！", vbInformation + vbOKOnly + vbDefaultButton1, "假名征服者(KanaMaster)"
    End Sub
