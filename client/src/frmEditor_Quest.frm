VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest System"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame7 
      Caption         =   "Quest Title"
      Height          =   975
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Rewards"
         Height          =   180
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Tasks"
         Height          =   180
         Index           =   3
         Left            =   4320
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optShowFrame 
         Caption         =   "Requirements"
         Height          =   180
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optShowFrame 
         Caption         =   "General"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   120
         MaxLength       =   30
         TabIndex        =   7
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quest List"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ListBox lstIndex 
         Height          =   7080
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraTasks 
      Caption         =   "Tasks"
      Height          =   6615
      Left            =   3600
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Frame Frame4 
         Caption         =   "Task Timer"
         Height          =   3375
         Left            =   3000
         TabIndex        =   86
         Top             =   3120
         Width           =   2175
         Begin VB.Frame Frame5 
            Caption         =   "Actions"
            Height          =   2175
            Left            =   0
            TabIndex        =   94
            Top             =   1200
            Width           =   2175
            Begin VB.TextBox txtMsg 
               Height          =   615
               Left            =   600
               MultiLine       =   -1  'True
               TabIndex        =   104
               Top             =   1440
               Width           =   1455
            End
            Begin VB.TextBox txtTaskY 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   1560
               TabIndex        =   102
               Text            =   "0"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtTaskX 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   1560
               TabIndex        =   101
               Text            =   "0"
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txtTaskTeleport 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   1560
               TabIndex        =   98
               Text            =   "0"
               Top             =   240
               Width           =   495
            End
            Begin VB.CheckBox chkTaskTeleport 
               Caption         =   "Telep?"
               Height          =   255
               Left            =   120
               TabIndex        =   97
               ToolTipText     =   "Deixe em 0 pra retornar pro mapa inicial da classe!"
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optReset 
               Caption         =   "Resetar Task?"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   96
               Top             =   960
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton optReset 
               Caption         =   "Resetar Quest?"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   95
               Top             =   1200
               Width           =   1575
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Msg?"
               Height          =   180
               Left            =   120
               TabIndex        =   105
               Top             =   1680
               Width           =   405
            End
            Begin VB.Label Label7 
               Caption         =   "Y:"
               Height          =   255
               Left            =   1320
               TabIndex        =   103
               ToolTipText     =   "Deixe em 0 pra retornar pro mapa inicial da classe!"
               Top             =   720
               Width           =   255
            End
            Begin VB.Label Label6 
               Caption         =   "X:"
               Height          =   255
               Left            =   1320
               TabIndex        =   100
               ToolTipText     =   "Deixe em 0 pra retornar pro mapa inicial da classe!"
               Top             =   480
               Width           =   255
            End
            Begin VB.Label Label5 
               Caption         =   "Map:"
               Height          =   255
               Left            =   1080
               TabIndex        =   99
               ToolTipText     =   "Deixe em 0 pra retornar pro mapa inicial da classe!"
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.TextBox txtTaskTimer 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   120
            TabIndex        =   92
            Text            =   "0"
            Top             =   840
            Width           =   735
         End
         Begin VB.OptionButton optTaskTimer 
            Caption         =   "Segs"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   91
            Top             =   960
            Width           =   735
         End
         Begin VB.OptionButton optTaskTimer 
            Caption         =   "Mins"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   90
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton optTaskTimer 
            Caption         =   "Horas"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   89
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optTaskTimer 
            Caption         =   "Dias"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   88
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CheckBox chkTaskTimer 
            Caption         =   "Tempo?"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Quant:"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5775
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   2775
         Begin VB.HScrollBar scrlNPC 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   43
            Top             =   1680
            Width           =   2535
         End
         Begin VB.HScrollBar scrlItem 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   42
            Top             =   2280
            Width           =   2535
         End
         Begin VB.HScrollBar scrlAmount 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   41
            Top             =   5040
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   40
            Top             =   2880
            Width           =   2535
         End
         Begin VB.TextBox txtTaskLog 
            Height          =   855
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   480
            Width           =   2535
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   38
            Top             =   3480
            Width           =   2535
         End
         Begin VB.CheckBox chkEnd 
            Caption         =   "End Quest Now?"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Left            =   120
            TabIndex        =   37
            Top             =   5400
            Width           =   1935
         End
         Begin VB.Label lblNPC 
            AutoSize        =   -1  'True
            Caption         =   "NPC: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "Item: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   48
            Top             =   2040
            Width           =   570
         End
         Begin VB.Label lblAmount 
            AutoSize        =   -1  'True
            Caption         =   "Amount: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   4800
            Width           =   795
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            X1              =   120
            X2              =   2640
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Label lblMap 
            AutoSize        =   -1  'True
            Caption         =   "Map: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   46
            Top             =   2640
            Width           =   525
         End
         Begin VB.Label lblLog 
            AutoSize        =   -1  'True
            Caption         =   "Task Log:"
            Height          =   180
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lblResource 
            AutoSize        =   -1  'True
            Caption         =   "Resource: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   44
            Top             =   3240
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   3000
         TabIndex        =   26
         Top             =   600
         Width           =   2175
         Begin VB.OptionButton optTask 
            Caption         =   "Nothing"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Slay NPC"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Gather Items"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Talk to NPC"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Reach Map"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Give Item to NPC"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   30
            Top             =   1560
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Kill Player"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Train with Resource"
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Top             =   2040
            Width           =   1815
         End
         Begin VB.OptionButton optTask 
            Caption         =   "Get from NPC"
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   27
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000000&
            X1              =   120
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.HScrollBar scrlTotalTasks 
         Height          =   255
         Left            =   1680
         Max             =   10
         Min             =   1
         TabIndex        =   24
         Top             =   240
         Value           =   1
         Width           =   3495
      End
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         Caption         =   "Selected Task: 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame fraRewards 
      Caption         =   "Rewards"
      Height          =   6495
      Left            =   3600
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtLevel 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   2760
         TabIndex        =   85
         Text            =   "0"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtExp 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   2760
         TabIndex        =   83
         Text            =   "0"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtlItemRewValue 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   720
         TabIndex        =   81
         Text            =   "0"
         Top             =   960
         Width           =   1815
      End
      Begin VB.HScrollBar scrlGiveSpell 
         Height          =   255
         LargeChange     =   50
         Left            =   2760
         TabIndex        =   79
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton cmdItemRewRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   1320
         TabIndex        =   66
         Top             =   3600
         Width           =   1215
      End
      Begin VB.ListBox lstItemRew 
         Height          =   2220
         ItemData        =   "frmEditor_Quest.frx":0000
         Left            =   120
         List            =   "frmEditor_Quest.frx":0007
         TabIndex        =   51
         Top             =   1320
         Width           =   2415
      End
      Begin VB.HScrollBar scrlItemRew 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   22
         Top             =   600
         Value           =   1
         Width           =   2415
      End
      Begin VB.CommandButton cmdItemRew 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   2760
         TabIndex        =   84
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quant:"
         Height          =   180
         Left            =   120
         TabIndex        =   82
         Top             =   960
         Width           =   525
      End
      Begin VB.Label lblGiveSpell 
         AutoSize        =   -1  'True
         Caption         =   "Spell: 0"
         Height          =   180
         Left            =   2760
         TabIndex        =   80
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         Caption         =   "Experience:"
         Height          =   180
         Left            =   2760
         TabIndex        =   50
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblItemRew 
         AutoSize        =   -1  'True
         Caption         =   "Item: 0 (1)"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame fraRequirements 
      Caption         =   "Requirements"
      Height          =   6495
      Left            =   3600
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.HScrollBar scrlReqClass 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   78
         Top             =   1680
         Value           =   1
         Width           =   2415
      End
      Begin VB.ListBox lstReqClass 
         Height          =   1140
         ItemData        =   "frmEditor_Quest.frx":0017
         Left            =   120
         List            =   "frmEditor_Quest.frx":0019
         TabIndex        =   77
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton cmdReqClassRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   1320
         TabIndex        =   75
         Top             =   3240
         Width           =   1215
      End
      Begin VB.HScrollBar scrlReqItemValue 
         Height          =   135
         Left            =   2760
         Max             =   10
         Min             =   1
         TabIndex        =   72
         Top             =   840
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar scrlReqItem 
         Height          =   255
         Left            =   2760
         Max             =   255
         TabIndex        =   71
         Top             =   480
         Value           =   1
         Width           =   2415
      End
      Begin VB.ListBox lstReqItem 
         Height          =   1860
         ItemData        =   "frmEditor_Quest.frx":001B
         Left            =   2760
         List            =   "frmEditor_Quest.frx":001D
         TabIndex        =   70
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdReqItemRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   3960
         TabIndex        =   68
         Top             =   3000
         Width           =   1215
      End
      Begin VB.HScrollBar scrlReqLevel 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin VB.HScrollBar scrlReqQuest 
         Height          =   255
         Left            =   120
         Max             =   70
         TabIndex        =   16
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdReqItem 
         Caption         =   "Update"
         Height          =   255
         Left            =   2760
         TabIndex        =   69
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdReqClass 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblReqClass 
         AutoSize        =   -1  'True
         Caption         =   "Class: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   74
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lblReqItem 
         AutoSize        =   -1  'True
         Caption         =   "Item Needed: 0 (1)"
         Height          =   180
         Left            =   2760
         TabIndex        =   73
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblReqLevel 
         AutoSize        =   -1  'True
         Caption         =   "Level: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblReqQuest 
         AutoSize        =   -1  'True
         Caption         =   "Quest: None"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   960
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      Height          =   6615
      Left            =   3600
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         Height          =   255
         Left            =   4800
         TabIndex        =   119
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         Height          =   255
         Left            =   4560
         TabIndex        =   118
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-"
         Height          =   255
         Left            =   4320
         TabIndex        =   116
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         Height          =   255
         Left            =   4080
         TabIndex        =   115
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "-"
         Height          =   255
         Left            =   3840
         TabIndex        =   113
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   255
         Left            =   3600
         TabIndex        =   112
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton optRepeat 
         Caption         =   "Normal?"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   111
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtSegs 
         Enabled         =   0   'False
         Height          =   270
         Left            =   2880
         TabIndex        =   109
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optRepeat 
         Caption         =   "Tempo?"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   108
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optRepeat 
         Caption         =   "Diaria?"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   107
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optRepeat 
         Caption         =   "Repeatitive?"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   106
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdTakeItemRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   3960
         TabIndex        =   65
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTakeItem 
         Caption         =   "Update"
         Height          =   255
         Left            =   2760
         TabIndex        =   64
         Top             =   6240
         Width           =   1215
      End
      Begin VB.ListBox lstTakeItem 
         Height          =   2040
         ItemData        =   "frmEditor_Quest.frx":001F
         Left            =   2760
         List            =   "frmEditor_Quest.frx":0021
         TabIndex        =   62
         Top             =   4080
         Width           =   2415
      End
      Begin VB.ListBox lstGiveItem 
         Height          =   2040
         ItemData        =   "frmEditor_Quest.frx":0023
         Left            =   120
         List            =   "frmEditor_Quest.frx":0025
         TabIndex        =   60
         Top             =   4080
         Width           =   2415
      End
      Begin VB.TextBox txtQuestLog 
         Height          =   270
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   58
         Top             =   240
         Width           =   3495
      End
      Begin VB.HScrollBar scrlTakeItem 
         Height          =   255
         Left            =   2760
         Max             =   255
         TabIndex        =   55
         Top             =   3480
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar scrlTakeItemValue 
         Height          =   135
         Left            =   2760
         Max             =   10
         Min             =   1
         TabIndex        =   54
         Top             =   3840
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar scrlGiveItemValue 
         Height          =   135
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   53
         Top             =   3840
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar scrlGiveItem 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   52
         Top             =   3480
         Value           =   1
         Width           =   2415
      End
      Begin VB.TextBox txtSpeech 
         Height          =   1455
         Left            =   120
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1440
         Width           =   5055
      End
      Begin VB.CommandButton cmdGiveItemRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   1320
         TabIndex        =   63
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGiveItem 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label lblRealTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo:"
         Height          =   180
         Left            =   2400
         TabIndex        =   121
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label12 
         Caption         =   "Secs"
         Height          =   255
         Left            =   4680
         TabIndex        =   120
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Minuts"
         Height          =   255
         Left            =   4080
         TabIndex        =   117
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Horas"
         Height          =   255
         Left            =   3600
         TabIndex        =   114
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Segs:"
         Height          =   255
         Left            =   2400
         TabIndex        =   110
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Starting Quest Log:"
         Height          =   180
         Left            =   120
         TabIndex        =   59
         Top             =   250
         Width           =   1485
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   5160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblTakeItem 
         AutoSize        =   -1  'True
         Caption         =   "Take Item on the End: 0 (1)"
         Height          =   180
         Left            =   2760
         TabIndex        =   57
         Top             =   3240
         Width           =   2100
      End
      Begin VB.Label lblGiveItem 
         AutoSize        =   -1  'True
         Caption         =   "Give Item on Start: 0 (1)"
         Height          =   180
         Left            =   120
         TabIndex        =   56
         Top             =   3240
         Width           =   1875
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   5160
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label lblQ1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////

Option Explicit
Private TempTask As Long

Private Sub chkTaskTeleport_Click()
    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.Teleport = chkTaskTeleport
End Sub

Private Sub chkTaskTimer_Click()
    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.Active = chkTaskTimer
End Sub

Private Sub Command1_Click()
    If Not txtSegs.enabled Then Exit Sub
    txtSegs = Int(txtSegs) + Int(3600)    ' 1 Hora tem 3600 segundos!
End Sub

Private Sub Command2_Click()
    If Not txtSegs.enabled Then Exit Sub
    If Int(txtSegs) >= Int(3600) Then
        txtSegs = Int(txtSegs) - Int(3600)    ' Retira 1 Hora tem 3600 segundos!
    Else
        txtSegs = 0
    End If
End Sub

Private Sub Command3_Click()
    If Not txtSegs.enabled Then Exit Sub
    txtSegs = Int(txtSegs) + Int(60)    ' 1 Minuto tem 60 segundos!
End Sub

Private Sub Command4_Click()
    If Not txtSegs.enabled Then Exit Sub
    If Int(txtSegs) >= 60 Then
        txtSegs = Int(txtSegs) - Int(60)    ' 1 Minuto tem 60 segundos!
    Else
        txtSegs = 0
    End If
End Sub

Private Sub Command5_Click()
    If Not txtSegs.enabled Then Exit Sub
    txtSegs = Int(txtSegs) + Int(1)
End Sub

Private Sub Command6_Click()
    If Not txtSegs.enabled Then Exit Sub
    If Int(txtSegs) >= Int(1) Then
        txtSegs = Int(txtSegs) - Int(1)
    Else
        txtSegs = 0
    End If
End Sub

Private Sub Form_Load()
    scrlTotalTasks.max = MAX_TASKS
    scrlNPC.max = MAX_NPCS
    scrlItem.max = MAX_ITEMS
    scrlMap.max = MAX_MAPS
    scrlResource.max = MAX_RESOURCES
    scrlAmount.max = MAX_BYTE
    scrlReqLevel.max = MAX_LEVELS
    scrlReqQuest.max = MAX_QUESTS
    scrlReqItem.max = MAX_ITEMS
    scrlReqItemValue.max = MAX_BYTE
    scrlGiveItem.max = MAX_ITEMS
    scrlGiveItemValue.max = MAX_BYTE
    scrlTakeItem.max = MAX_ITEMS
    scrlTakeItemValue.max = MAX_BYTE
    scrlItemRew.max = MAX_ITEMS
    scrlGiveSpell.max = MAX_SPELLS
End Sub

Private Sub cmdSave_Click()
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        QuestEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    QuestEditorCancel
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    ClearQuest EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    QuestEditorInit
End Sub

Private Sub lstIndex_Click()
    QuestEditorInit
End Sub

Private Sub optRepeat_Click(Index As Integer)

    If Index <> 3 Then
        txtSegs = 0
        txtSegs.enabled = False
    Else
        txtSegs.enabled = True
    End If

    Quest(EditorIndex).Repeat = Index
End Sub

Private Sub optReset_Click(Index As Integer)
    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.ResetType = Index
End Sub

Private Sub optTaskTimer_Click(Index As Integer)
    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.TimerType = Index
End Sub

Private Sub scrlGiveSpell_Change()
    If scrlGiveSpell.Value > 0 Then lblGiveSpell.caption = "Spell Reward: (" & scrlGiveSpell.Value & ")" & Trim$(Spell(scrlGiveSpell).Name) Else: lblGiveSpell.caption = "Spell Reward: (" & scrlGiveSpell.Value & ")"
    Quest(EditorIndex).RewardSpell = scrlGiveSpell.Value
End Sub

Private Sub scrlTotalTasks_Change()
    Dim i As Long

    lblSelected = "Selected Task: " & scrlTotalTasks.Value

    LoadTask EditorIndex, scrlTotalTasks.Value
End Sub

Private Sub optTask_Click(Index As Integer)
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Order = Index
    LoadTask EditorIndex, scrlTotalTasks.Value
End Sub

Private Sub txtEXP_Change()
    If Not IsNumeric(txtExp.text) Then txtExp.text = 0
    If txtExp.text > MAX_LONG Then txtExp.text = 0

    Quest(EditorIndex).RewardExp = txtExp.text
End Sub

Private Sub txtLevel_Change()
    If Not IsNumeric(txtLevel.text) Then txtLevel.text = 0
    If txtLevel.text > MAX_LONG Then txtLevel.text = 0

    Quest(EditorIndex).RewardLevel = txtLevel.text
End Sub

Private Sub txtlItemRewValue_Change()
    If Not IsNumeric(txtlItemRewValue.text) Then txtlItemRewValue.text = 0
    If txtlItemRewValue.text > MAX_LONG Then txtlItemRewValue.text = 0

    If scrlItemRew.Value > 0 Then
        lblItemRew.caption = "Item: " & scrlItemRew.Value & "-" & Trim$(Item(scrlItemRew.Value).Name) & "(" & txtlItemRewValue.text & ")"
    Else
        lblItemRew.caption = "Item: " & scrlItemRew.Value & "(" & txtlItemRewValue.text & ")"
    End If
End Sub

Private Sub txtMsg_Change()
    If txtMsg = vbNullString Then
        Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.Msg = vbNullString
    Else
        Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.Msg = txtMsg
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long
    tmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtQuestLog_Change()
    Quest(EditorIndex).QuestLog = Trim$(txtQuestLog.text)
End Sub

Private Sub txtSegs_Change()

    If Not IsNumeric(txtSegs) Then
        txtSegs = 0
    End If

    lblRealTime = "Tempo: " & SecondsToHMS(txtSegs)

    Quest(EditorIndex).Time = txtSegs
End Sub

Private Sub txtSpeech_Change()
    Quest(EditorIndex).Speech = Trim$(txtSpeech.text)
End Sub

Private Sub txtTaskLog_Change()
    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskLog = Trim$(txtTaskLog.text)
End Sub

Private Sub scrlReqLevel_Change()
    lblReqLevel.caption = "Level: " & scrlReqLevel.Value
    Quest(EditorIndex).RequiredLevel = scrlReqLevel.Value
End Sub

Private Sub scrlReqQuest_Change()
    If Not scrlReqQuest.Value = 0 Then
        If Not Trim$(Quest(scrlReqQuest.Value).Name) = "" Then
            lblReqQuest.caption = "Quest: " & Trim$(Quest(scrlReqQuest.Value).Name)
        Else
            lblReqQuest.caption = "Quest: None"
        End If
    Else
        lblReqQuest.caption = "Quest: None"
    End If
    Quest(EditorIndex).RequiredQuest = scrlReqQuest.Value
End Sub

'Alatar v1.2

Private Sub scrlReqItem_Change()
    lblReqItem.caption = "Item Needed: " & scrlReqItem.Value & " (" & scrlReqItemValue.Value & ")"
End Sub

Private Sub scrlReqItemValue_Change()
    lblReqItem.caption = "Item Needed: " & scrlReqItem.Value & " (" & scrlReqItemValue.Value & ")"
End Sub

Private Sub cmdReqItem_Click()
    Dim Index As Long

    Index = lstReqItem.ListIndex + 1    'the selected item
    If Index = 0 Then Exit Sub
    If scrlReqItem.Value < 1 Or scrlReqItem.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlReqItem.Value).Name) = "" Then Exit Sub

    Quest(EditorIndex).RequiredItem(Index).Item = scrlReqItem.Value
    Quest(EditorIndex).RequiredItem(Index).Value = scrlReqItemValue.Value
    UpdateQuestRequirementItems
End Sub

Private Sub cmdReqItemRemove_Click()
    Dim Index As Long

    Index = lstReqItem.ListIndex + 1
    If Index = 0 Then Exit Sub

    Quest(EditorIndex).RequiredItem(Index).Item = 0
    Quest(EditorIndex).RequiredItem(Index).Value = 1
    UpdateQuestRequirementItems
End Sub

Private Sub scrlReqClass_Change()
    If scrlReqClass.Value < 1 Or scrlReqClass.Value > Max_Classes Then
        lblReqClass.caption = "Class: 0"
    Else
        lblReqClass.caption = "Class: " & scrlReqClass.Value & " (" & Trim$(Class(scrlReqClass.Value).Name) & ")"
    End If
End Sub

Private Sub cmdReqClass_Click()
    Dim Index As Long

    Index = lstReqClass.ListIndex + 1    'the selected class
    If Index = 0 Then Exit Sub
    If scrlReqClass.Value < 1 Or scrlReqClass.Value > Max_Classes Then Exit Sub
    If Trim$(Class(scrlReqClass.Value).Name) = "" Then Exit Sub

    Quest(EditorIndex).RequiredClass(Index) = scrlReqClass.Value
    UpdateQuestClass
End Sub

Private Sub cmdReqClassRemove_Click()
    Dim Index As Long

    Index = lstReqClass.ListIndex + 1
    If Index = 0 Then Exit Sub

    Quest(EditorIndex).RequiredClass(Index) = 0
    UpdateQuestClass
End Sub

Private Sub scrlItemRew_Change()
    If scrlItemRew > 0 Then
        lblItemRew.caption = "Item: " & scrlItemRew.Value & "-" & Trim$(Item(scrlItemRew.Value).Name) & "(" & txtlItemRewValue.text & ")"
    Else
        lblItemRew.caption = "Item: " & scrlItemRew.Value & "(" & txtlItemRewValue.text & ")"
    End If
End Sub

'Alatar v1.2
Private Sub cmdItemRew_Click()
    Dim Index As Long

    Index = lstItemRew.ListIndex + 1    'the selected item
    If Index = 0 Then Exit Sub
    If scrlItemRew.Value < 1 Or scrlItemRew.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlItemRew.Value).Name) = "" Then Exit Sub

    Quest(EditorIndex).RewardItem(Index).Item = scrlItemRew.Value
    Quest(EditorIndex).RewardItem(Index).Value = txtlItemRewValue.text
    UpdateQuestRewardItems
End Sub

Private Sub cmdItemRewRemove_Click()
    Dim Index As Long

    Index = lstItemRew.ListIndex + 1
    If Index = 0 Then Exit Sub

    Quest(EditorIndex).RewardItem(Index).Item = 0
    Quest(EditorIndex).RewardItem(Index).Value = 1
    UpdateQuestRewardItems
End Sub
'/Alatar v1.2

'Alatar v1.2
Private Sub scrlGiveItem_Change()
    lblGiveItem = "Give Item on Start: " & scrlGiveItem.Value & " (" & scrlGiveItemValue.Value & ")"
End Sub

Private Sub scrlGiveItemValue_Change()
    lblGiveItem = "Give Item on Start: " & scrlGiveItem.Value & " (" & scrlGiveItemValue.Value & ")"
End Sub

Private Sub cmdGiveItem_Click()
    Dim Index As Long

    Index = lstGiveItem.ListIndex + 1    'the selected item
    If Index = 0 Then Exit Sub
    If scrlGiveItem.Value < 1 Or scrlGiveItem.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlGiveItem.Value).Name) = "" Then Exit Sub

    Quest(EditorIndex).GiveItem(Index).Item = scrlGiveItem.Value
    Quest(EditorIndex).GiveItem(Index).Value = scrlGiveItemValue.Value
    UpdateQuestGiveItems
End Sub

Private Sub cmdGiveItemRemove_Click()
    Dim Index As Long

    Index = lstGiveItem.ListIndex + 1
    If Index = 0 Then Exit Sub

    Quest(EditorIndex).GiveItem(Index).Item = 0
    Quest(EditorIndex).GiveItem(Index).Value = 1
    UpdateQuestGiveItems
End Sub

Private Sub scrlTakeItem_Change()
    lblTakeItem = "Take Item on the End: " & scrlTakeItem.Value & " (" & scrlTakeItemValue.Value & ")"
End Sub

Private Sub scrlTakeItemValue_Change()
    lblTakeItem = "Take Item on the End: " & scrlTakeItem.Value & " (" & scrlTakeItemValue.Value & ")"
End Sub

Private Sub cmdTakeItem_Click()
    Dim Index As Long

    Index = lstTakeItem.ListIndex + 1    'the selected item
    If Index = 0 Then Exit Sub
    If scrlTakeItem.Value < 1 Or scrlTakeItem.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlTakeItem.Value).Name) = "" Then Exit Sub

    Quest(EditorIndex).TakeItem(Index).Item = scrlTakeItem.Value
    Quest(EditorIndex).TakeItem(Index).Value = scrlTakeItemValue.Value
    UpdateQuestTakeItems
End Sub

Private Sub cmdTakeItemRemove_Click()
    Dim Index As Long

    Index = lstTakeItem.ListIndex + 1
    If Index = 0 Then Exit Sub

    Quest(EditorIndex).TakeItem(Index).Item = 0
    Quest(EditorIndex).TakeItem(Index).Value = 1
    UpdateQuestTakeItems
End Sub
'/Alatar v1.2

Private Sub scrlAmount_Change()
    lblAmount.caption = "Amount: " & scrlAmount.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Amount = scrlAmount.Value
End Sub

Private Sub scrlNPC_Change()
    If scrlNPC.Value > 0 Then
        lblNPC.caption = "NPC: " & scrlNPC.Value & "-" & Trim$(NPC(scrlNPC.Value).Name)
    Else
        lblNPC.caption = "NPC: " & scrlNPC.Value
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.Value).NPC = scrlNPC.Value
End Sub

Private Sub scrlItem_Change()
    If scrlItem.Value > 0 Then
        lblItem.caption = "Item: " & scrlItem.Value & "-" & Trim$(Item(scrlItem.Value).Name)
    Else
        lblItem.caption = "Item: " & scrlItem.Value
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Item = scrlItem.Value
End Sub

Private Sub scrlMap_Change()
    lblMap.caption = "Map: " & scrlMap.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Map = scrlMap.Value
End Sub

Private Sub scrlResource_Change()
    If scrlResource.Value > 0 Then
        lblResource.caption = "Res: " & scrlResource.Value & "-" & Trim$(Resource(scrlResource.Value).Name)
    Else
        lblResource.caption = "NPC: " & scrlResource.Value
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Resource = scrlResource.Value
End Sub

Private Sub chkEnd_Click()
    If chkEnd.Value = 1 Then
        Quest(EditorIndex).Task(scrlTotalTasks.Value).QuestEnd = True
    Else
        Quest(EditorIndex).Task(scrlTotalTasks.Value).QuestEnd = False
    End If
End Sub

Private Sub optShowFrame_Click(Index As Integer)
    fraGeneral.visible = False
    fraRequirements.visible = False
    fraRewards.visible = False
    fraTasks.visible = False

    If optShowFrame(Index).Value = True Then
        Select Case Index
        Case 0
            fraGeneral.visible = True
        Case 1
            fraRequirements.visible = True
        Case 2
            fraRewards.visible = True
        Case 3
            fraTasks.visible = True
        End Select
    End If
End Sub

Private Sub txtTaskTeleport_Change()
    If Not IsNumeric(txtTaskTeleport) Then
        txtTaskTeleport = Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.MapNum
        Exit Sub
    End If

    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.MapNum = txtTaskTeleport
End Sub

Private Sub txtTaskTimer_Change()
    If Not IsNumeric(txtTaskTimer) Then
        txtTaskTimer = Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.Timer
        Exit Sub
    End If

    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.Timer = txtTaskTimer.text
End Sub

Private Sub txtTaskX_Change()
    If Not IsNumeric(txtTaskX) Then
        txtTaskX = Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.X
        Exit Sub
    End If

    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.X = txtTaskX
End Sub

Private Sub txtTaskY_Change()
    If Not IsNumeric(txtTaskY) Then
        txtTaskY = Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.Y
        Exit Sub
    End If

    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskTimer.Y = txtTaskY
End Sub
