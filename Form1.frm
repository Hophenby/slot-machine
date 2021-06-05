VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "摇号机"
   ClientHeight    =   10710
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15615
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Height          =   1815
      Left            =   120
      TabIndex        =   87
      Top             =   8760
      Visible         =   0   'False
      Width           =   15375
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   80
      Text            =   "42"
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "调整"
      Height          =   1095
      Left            =   840
      TabIndex        =   72
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   600
      TabIndex        =   70
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command10 
         Caption         =   "重置"
         Height          =   495
         Left            =   480
         TabIndex        =   71
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "按下面的按钮以重置调整"
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label16 
         Caption         =   "支持批量操作"
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "调整模式下点击相应座位即可调整"
         Height          =   375
         Left            =   240
         TabIndex        =   84
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "!表示参与,x表示不参与"
         Height          =   375
         Left            =   240
         TabIndex        =   83
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "说明:"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "每竖排抽一人"
      Height          =   495
      Left            =   14040
      TabIndex        =   54
      Top             =   9600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "每横排抽一人"
      Height          =   495
      Left            =   14040
      TabIndex        =   53
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "5连抽"
      Height          =   495
      Left            =   12720
      TabIndex        =   52
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "10连抽"
      Height          =   495
      Left            =   12720
      TabIndex        =   51
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "摇号机"
      Height          =   6255
      Left            =   3720
      TabIndex        =   14
      Top             =   2520
      Width           =   6855
      Begin VB.CommandButton Command6 
         Height          =   375
         Index           =   6
         Left            =   6120
         TabIndex        =   79
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Index           =   5
         Left            =   6120
         TabIndex        =   78
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Index           =   4
         Left            =   6120
         TabIndex        =   77
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Index           =   3
         Left            =   6120
         TabIndex        =   76
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Index           =   2
         Left            =   6120
         TabIndex        =   75
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   74
         Top             =   4200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   69
         Top             =   4800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Index           =   5
         Left            =   1560
         TabIndex        =   68
         Top             =   5520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Index           =   4
         Left            =   2520
         TabIndex        =   67
         Top             =   5520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   66
         Top             =   5520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   65
         Top             =   5520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   64
         Top             =   5520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   63
         Top             =   5520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "后门"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   5760
         TabIndex        =   88
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "前门"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   5640
         TabIndex        =   62
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "讲台"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   61
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   60
         Left            =   600
         TabIndex        =   60
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   50
         Left            =   1560
         TabIndex        =   59
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   40
         Left            =   2520
         TabIndex        =   58
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   3480
         TabIndex        =   57
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   4440
         TabIndex        =   56
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   5400
         TabIndex        =   55
         Top             =   4800
         Width           =   375
      End
      Begin VB.Line Line3 
         Index           =   10
         X1              =   2400
         X2              =   3000
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line3 
         Index           =   9
         X1              =   1440
         X2              =   2040
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line3 
         Index           =   8
         X1              =   3360
         X2              =   3960
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line3 
         Index           =   7
         X1              =   4320
         X2              =   4920
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line3 
         Index           =   6
         X1              =   5280
         X2              =   5880
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line2 
         X1              =   480
         X2              =   1080
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line13 
         Index           =   4
         X1              =   5280
         X2              =   5280
         Y1              =   1080
         Y2              =   5280
      End
      Begin VB.Line Line12 
         Index           =   5
         X1              =   5280
         X2              =   5880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line11 
         Index           =   5
         X1              =   5280
         X2              =   5880
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line10 
         Index           =   1
         X1              =   5880
         X2              =   5880
         Y1              =   5280
         Y2              =   1080
      End
      Begin VB.Line Line9 
         Index           =   1
         X1              =   5880
         X2              =   5280
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   5280
         X2              =   5880
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   5280
         X2              =   5880
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   5280
         X2              =   5880
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   5280
         X2              =   5880
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line13 
         Index           =   3
         X1              =   4320
         X2              =   4320
         Y1              =   1080
         Y2              =   5280
      End
      Begin VB.Line Line12 
         Index           =   0
         X1              =   4320
         X2              =   4920
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line11 
         Index           =   0
         X1              =   4320
         X2              =   4920
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line10 
         Index           =   0
         X1              =   4920
         X2              =   4920
         Y1              =   5280
         Y2              =   1080
      End
      Begin VB.Line Line9 
         Index           =   0
         X1              =   4920
         X2              =   4320
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   4320
         X2              =   4920
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   4320
         X2              =   4920
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   4320
         X2              =   4920
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   4320
         X2              =   4920
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line13 
         Index           =   2
         X1              =   3360
         X2              =   3360
         Y1              =   1080
         Y2              =   5280
      End
      Begin VB.Line Line12 
         Index           =   4
         X1              =   3360
         X2              =   3960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line11 
         Index           =   4
         X1              =   3360
         X2              =   3960
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line10 
         Index           =   5
         X1              =   3960
         X2              =   3960
         Y1              =   5280
         Y2              =   1080
      End
      Begin VB.Line Line9 
         Index           =   6
         X1              =   3960
         X2              =   3360
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line8 
         Index           =   6
         X1              =   3360
         X2              =   3960
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line5 
         Index           =   5
         X1              =   3360
         X2              =   3960
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line4 
         Index           =   5
         X1              =   3360
         X2              =   3960
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line3 
         Index           =   5
         X1              =   3360
         X2              =   3960
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line13 
         Index           =   1
         X1              =   1440
         X2              =   1440
         Y1              =   1080
         Y2              =   5280
      End
      Begin VB.Line Line12 
         Index           =   3
         X1              =   1440
         X2              =   2040
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line11 
         Index           =   3
         X1              =   1440
         X2              =   2040
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line10 
         Index           =   4
         X1              =   2040
         X2              =   2040
         Y1              =   5280
         Y2              =   1080
      End
      Begin VB.Line Line9 
         Index           =   5
         X1              =   2040
         X2              =   1440
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line8 
         Index           =   5
         X1              =   1440
         X2              =   2040
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line5 
         Index           =   4
         X1              =   1440
         X2              =   2040
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line4 
         Index           =   4
         X1              =   1440
         X2              =   2040
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line3 
         Index           =   4
         X1              =   1440
         X2              =   2040
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   480
         X2              =   1080
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line1 
         X1              =   480
         X2              =   480
         Y1              =   1080
         Y2              =   5280
      End
      Begin VB.Line Line12 
         Index           =   2
         X1              =   480
         X2              =   1080
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line11 
         Index           =   2
         X1              =   480
         X2              =   1080
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line10 
         Index           =   3
         X1              =   1080
         X2              =   1080
         Y1              =   5280
         Y2              =   1080
      End
      Begin VB.Line Line9 
         Index           =   3
         X1              =   1080
         X2              =   480
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line8 
         Index           =   3
         X1              =   480
         X2              =   1080
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line5 
         Index           =   3
         X1              =   480
         X2              =   1080
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   480
         X2              =   1080
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   2400
         X2              =   3000
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   2400
         X2              =   3000
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line5 
         Index           =   2
         X1              =   2400
         X2              =   3000
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line8 
         Index           =   2
         X1              =   2400
         X2              =   3000
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line9 
         Index           =   2
         X1              =   3000
         X2              =   2400
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line10 
         Index           =   2
         X1              =   3000
         X2              =   3000
         Y1              =   5280
         Y2              =   1080
      End
      Begin VB.Line Line11 
         Index           =   1
         X1              =   2400
         X2              =   3000
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line12 
         Index           =   1
         X1              =   2400
         X2              =   3000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line13 
         Index           =   0
         X1              =   2400
         X2              =   2400
         Y1              =   1080
         Y2              =   5280
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   5400
         TabIndex        =   50
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   5400
         TabIndex        =   49
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   5400
         TabIndex        =   48
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   4440
         TabIndex        =   47
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   5400
         TabIndex        =   46
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   4440
         TabIndex        =   45
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   4440
         TabIndex        =   44
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   4440
         TabIndex        =   43
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   3480
         TabIndex        =   42
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   32
         Left            =   3480
         TabIndex        =   41
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   33
         Left            =   3480
         TabIndex        =   40
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   34
         Left            =   3480
         TabIndex        =   39
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   3480
         TabIndex        =   38
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   3480
         TabIndex        =   37
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   41
         Left            =   2520
         TabIndex        =   36
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   42
         Left            =   2520
         TabIndex        =   35
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   2520
         TabIndex        =   34
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   44
         Left            =   2520
         TabIndex        =   33
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   45
         Left            =   2520
         TabIndex        =   32
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   46
         Left            =   2520
         TabIndex        =   31
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   51
         Left            =   1560
         TabIndex        =   30
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   52
         Left            =   1560
         TabIndex        =   29
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   53
         Left            =   1560
         TabIndex        =   28
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   54
         Left            =   1560
         TabIndex        =   27
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   55
         Left            =   1560
         TabIndex        =   26
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   56
         Left            =   1560
         TabIndex        =   25
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   61
         Left            =   600
         TabIndex        =   24
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   62
         Left            =   600
         TabIndex        =   23
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   63
         Left            =   600
         TabIndex        =   22
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   64
         Left            =   600
         TabIndex        =   21
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   65
         Left            =   600
         TabIndex        =   20
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   66
         Left            =   600
         TabIndex        =   19
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   4440
         TabIndex        =   18
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   4440
         TabIndex        =   17
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   5400
         TabIndex        =   16
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label ex 
         Caption         =   "！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   5400
         TabIndex        =   15
         Top             =   1800
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   3720
      TabIndex        =   9
      Top             =   9000
      Width           =   2175
      Begin VB.CommandButton Command8 
         Caption         =   "<"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "5"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   ">"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "摇出人数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "摇号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8040
      TabIndex        =   8
      Top             =   8880
      Width           =   4095
   End
   Begin VB.Timer Timer2 
      Left            =   10440
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10080
      Top             =   1920
   End
   Begin VB.CommandButton Command12 
      Caption         =   "完成"
      Height          =   1095
      Left            =   840
      TabIndex        =   73
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "人数"
      Height          =   735
      Left            =   480
      TabIndex        =   81
      Top             =   7440
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   1470
      Left            =   6600
      Picture         =   "Form1.frx":324A
      Top             =   8880
      Width           =   1485
   End
   Begin VB.Label Label12 
      Caption         =   "---习近平"
      Height          =   255
      Left            =   13920
      TabIndex        =   7
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "基础不牢, 地动山摇。"
      Height          =   375
      Left            =   11040
      TabIndex        =   6
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "---President Xi"
      Height          =   255
      Left            =   13320
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Without a solid foundation,it would lead to serious vulnerabilities."
      Height          =   615
      Left            =   11040
      TabIndex        =   4
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label Label8 
      Caption         =   "Knowledge makes humble, ignorance makes proud."
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label7 
      Caption         =   "Re-Edited"
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "By 霍芬比 Hophenby "
      Height          =   615
      Left            =   10080
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "摇号机"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4440
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
For i = 0 To 6
        For j = 1 To 6
            
               ex(j & i).Visible = False
            
        Next
    Next
    For f = 1 To 10
    Do
       
            p = Int(Rnd * 7)
            q = Int(Rnd * 6) + 1
          If ex(q & p).Caption = "！" And ex(q & p).Visible = 0 Then
       Exit Do
       End If
    Loop While (1)
        ex(q & p).Visible = True
    Next
End Sub

Private Sub Command10_Click()
For i = 0 To 6
        For j = 1 To 6
            
               ex(j & i).Caption = "！"
            
        Next
    Next

End Sub

Private Sub Command11_Click()
Frame2.Visible = 1
Command11.Visible = 0
Frame4.Visible = 1

For i = 0 To 6
        For j = 1 To 6
            Command5(j).Visible = 1
            Command6(i).Visible = 1
            
               ex(j & i).Visible = 1
            
        Next
    Next

End Sub

Private Sub Command12_Click()
ct = 0
For i = 0 To 6
        For j = 1 To 6
            
             If ex(j & i).Caption = "！" Then
             ct = ct + 1
            End If
        Next
    Next
If ct >= 10 Then
Command11.Visible = 1
Frame2.Visible = 0
Frame4.Visible = 0
For i = 0 To 6
        For j = 1 To 6
            Command5(j).Visible = 0
            Command6(i).Visible = 0
            
               ex(j & i).Visible = 1
            
        Next
    Next
    Text1.Text = ct
    Else
    MsgBox "参加人数太少!"
    End If
End Sub

Private Sub Command2_Click()
For i = 0 To 6
        For j = 1 To 6
            
               ex(j & i).Visible = False
            
        Next
    Next
    For f = 1 To 5
    Do
       
            p = Int(Rnd * 7)
            q = Int(Rnd * 6) + 1
          If ex(q & p).Caption = "！" And ex(q & p).Visible = 0 Then
       Exit Do
       End If
    Loop While (1)
        ex(q & p).Visible = True
    Next
End Sub

Private Sub Command3_Click()
For i = 0 To 6
        For j = 1 To 6
            
               ex(j & i).Visible = False
            
        Next
    Next
 For f = 0 To 6
  ct = 0
 For w = 1 To 6
 If ex(w & f).Caption = "！" Then
 ct = ct + 1
 End If
 Next
 If ct <> 0 Then
    Do
       
            p = f
            q = Int(Rnd * 6) + 1
       If ex(q & p).Caption = "！" And ex(q & p).Visible = 0 Then
       Exit Do
       End If
    Loop While (1)
        ex(q & p).Visible = True
    
    End If
    Next
End Sub

Private Sub Command4_Click()
For i = 0 To 6
        For j = 1 To 6
            
               ex(j & i).Visible = False
            
        Next
    Next
 For f = 1 To 6
 ct = 0
 For w = 0 To 6
 If ex(f & w).Caption = "！" Then
 ct = ct + 1
 End If
 Next
 If ct <> 0 Then
    Do
       
            p = Int(Rnd * 7)
            q = f
          If ex(q & p).Caption = "！" And ex(q & p).Visible = 0 Then
       Exit Do
       End If
    Loop While (1)
        ex(q & p).Visible = True
    
    End If
    Next
End Sub




Private Sub Command5_Click(Index As Integer)
For i = 0 To 6
       
          If ex(Index & i).Caption <> "x" Then
               ex(Index & i).Caption = "x"
            Else
            ex(Index & i).Caption = "！"
       End If
    Next
End Sub

Private Sub Command6_Click(Index As Integer)
For j = 1 To 6
       
          If ex(j & Index).Caption <> "x" Then
               ex(j & Index).Caption = "x"
            Else
            ex(j & Index).Caption = "！"
 End If
    Next
End Sub

Private Sub Command7_Click()
If Text4.Text <> 0 And Timer1.Enabled = 0 Then
  Timer1.Enabled = 1
  Command7.Caption = "停止"
ElseIf Text4.Text = 0 Then
MsgBox "人数不能为零!"
Else

Timer1.Enabled = 0
  Command7.Caption = "摇号"
  
End If
End Sub

Private Sub Command8_Click()
If Text4.Text <> 0 Then
Text4.Text = Text4.Text - 1
End If
End Sub

Private Sub Command9_Click()
If Text4.Text <> Text1.Text Then
Text4.Text = Text4.Text + 1
End If
End Sub

Private Sub ex_Click(Index As Integer)
If Index <> 0 And ex(Index).Caption = "！" And Command11.Visible = 0 Then
ex(Index).Caption = "x"
ElseIf Index <> 0 And ex(Index).Caption = "x" And Command11.Visible = 0 Then
 ex(Index).Caption = "！"
End If
End Sub

Private Sub Form_Load()
    Randomize
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()
For i = 0 To 6
        For j = 1 To 6
            
               ex(j & i).Visible = False
            
        Next
    Next
    For f = 1 To Text4.Text
    Do
       
            p = Int(Rnd * 7)
            q = Int(Rnd * 6) + 1
          If ex(q & p).Caption = "！" And ex(q & p).Visible = 0 Then
       Exit Do
       End If
    Loop While (1)
        ex(q & p).Visible = True
    Next
End Sub






Private Sub 座位模式_Click()

End Sub
