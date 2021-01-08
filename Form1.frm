VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "摇号机"
   ClientHeight    =   10710
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15615
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   3720
      TabIndex        =   84
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton Command8 
         Caption         =   "<"
         Height          =   375
         Left            =   360
         TabIndex        =   87
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   ">"
         Height          =   375
         Left            =   1440
         TabIndex        =   85
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
         TabIndex        =   88
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
      TabIndex        =   83
      Top             =   8880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "摇号结果"
      Height          =   4935
      Left            =   2880
      TabIndex        =   38
      Top             =   2880
      Visible         =   0   'False
      Width           =   7815
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
         Index           =   86
         Left            =   480
         TabIndex        =   82
         Top             =   1320
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
         Index           =   85
         Left            =   480
         TabIndex        =   81
         Top             =   1920
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
         Index           =   84
         Left            =   480
         TabIndex        =   80
         Top             =   2520
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
         Index           =   83
         Left            =   480
         TabIndex        =   79
         Top             =   3120
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
         Index           =   82
         Left            =   480
         TabIndex        =   78
         Top             =   3720
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
         Index           =   81
         Left            =   480
         TabIndex        =   77
         Top             =   4320
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
         Index           =   76
         Left            =   1080
         TabIndex        =   76
         Top             =   1320
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
         Index           =   75
         Left            =   1080
         TabIndex        =   75
         Top             =   1920
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
         Index           =   74
         Left            =   1080
         TabIndex        =   74
         Top             =   2520
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
         Index           =   73
         Left            =   1080
         TabIndex        =   73
         Top             =   3120
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
         Index           =   72
         Left            =   1080
         TabIndex        =   72
         Top             =   3720
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
         Index           =   71
         Left            =   1080
         TabIndex        =   71
         Top             =   4320
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
         Left            =   2400
         TabIndex        =   70
         Top             =   1320
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
         Left            =   2400
         TabIndex        =   69
         Top             =   1920
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
         Left            =   2400
         TabIndex        =   68
         Top             =   2520
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
         Left            =   2400
         TabIndex        =   67
         Top             =   3120
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
         Left            =   2400
         TabIndex        =   66
         Top             =   3720
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
         Left            =   2400
         TabIndex        =   65
         Top             =   4320
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
         Left            =   3000
         TabIndex        =   64
         Top             =   1320
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
         Left            =   3000
         TabIndex        =   63
         Top             =   1920
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
         Left            =   3000
         TabIndex        =   62
         Top             =   2520
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
         Left            =   3000
         TabIndex        =   61
         Top             =   3120
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
         Left            =   3000
         TabIndex        =   60
         Top             =   3720
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
         Left            =   3000
         TabIndex        =   59
         Top             =   4320
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
         Left            =   4320
         TabIndex        =   58
         Top             =   1320
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
         Left            =   4320
         TabIndex        =   57
         Top             =   1920
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
         Left            =   4320
         TabIndex        =   56
         Top             =   2520
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
         Left            =   4320
         TabIndex        =   55
         Top             =   3120
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
         Left            =   4320
         TabIndex        =   54
         Top             =   3720
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
         Left            =   4320
         TabIndex        =   53
         Top             =   4320
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
         Left            =   4920
         TabIndex        =   52
         Top             =   1320
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
         Left            =   4920
         TabIndex        =   51
         Top             =   1920
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
         Left            =   4920
         TabIndex        =   50
         Top             =   2520
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
         Left            =   4920
         TabIndex        =   49
         Top             =   3120
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
         Left            =   4920
         TabIndex        =   48
         Top             =   3720
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
         Left            =   4920
         TabIndex        =   47
         Top             =   4320
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
         Left            =   6240
         TabIndex        =   46
         Top             =   2520
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
         Left            =   6240
         TabIndex        =   45
         Top             =   3120
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
         Left            =   6240
         TabIndex        =   44
         Top             =   3720
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
         Left            =   6840
         TabIndex        =   43
         Top             =   2520
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
         Left            =   6240
         TabIndex        =   42
         Top             =   4320
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
         Left            =   6840
         TabIndex        =   41
         Top             =   3120
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
         Left            =   6840
         TabIndex        =   40
         Top             =   3720
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
         Index           =   11
         Left            =   6840
         TabIndex        =   39
         Top             =   4320
         Width           =   375
      End
      Begin VB.Line Line14 
         X1              =   360
         X2              =   360
         Y1              =   1200
         Y2              =   4800
      End
      Begin VB.Line Line13 
         X1              =   2280
         X2              =   2280
         Y1              =   1200
         Y2              =   4800
      End
      Begin VB.Line Line12 
         Index           =   2
         X1              =   360
         X2              =   1560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line11 
         Index           =   2
         X1              =   360
         X2              =   1560
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line10 
         Index           =   3
         X1              =   960
         X2              =   960
         Y1              =   4800
         Y2              =   1200
      End
      Begin VB.Line Line9 
         Index           =   3
         X1              =   1560
         X2              =   360
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line8 
         Index           =   3
         X1              =   360
         X2              =   1560
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line7 
         Index           =   3
         X1              =   1560
         X2              =   1560
         Y1              =   4800
         Y2              =   1200
      End
      Begin VB.Line Line5 
         Index           =   3
         X1              =   360
         X2              =   1560
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   360
         X2              =   1560
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   360
         X2              =   1560
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line12 
         Index           =   1
         X1              =   2280
         X2              =   3480
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line11 
         Index           =   1
         X1              =   2280
         X2              =   3480
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line10 
         Index           =   2
         X1              =   2880
         X2              =   2880
         Y1              =   4800
         Y2              =   1200
      End
      Begin VB.Line Line9 
         Index           =   2
         X1              =   3480
         X2              =   2280
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line8 
         Index           =   2
         X1              =   2280
         X2              =   3480
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line7 
         Index           =   2
         X1              =   3480
         X2              =   3480
         Y1              =   4800
         Y2              =   1200
      End
      Begin VB.Line Line5 
         Index           =   2
         X1              =   2280
         X2              =   3480
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   2280
         X2              =   3480
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   2280
         X2              =   3480
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line12 
         Index           =   0
         X1              =   4200
         X2              =   5400
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line11 
         Index           =   0
         X1              =   4200
         X2              =   5400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line10 
         Index           =   1
         X1              =   4800
         X2              =   4800
         Y1              =   4800
         Y2              =   1200
      End
      Begin VB.Line Line9 
         Index           =   1
         X1              =   5400
         X2              =   4200
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   4200
         X2              =   5400
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line7 
         Index           =   1
         X1              =   5400
         X2              =   5400
         Y1              =   4800
         Y2              =   1200
      End
      Begin VB.Line Line6 
         Index           =   1
         X1              =   4200
         X2              =   4200
         Y1              =   4800
         Y2              =   1200
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   4200
         X2              =   5400
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   4200
         X2              =   5400
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   4200
         X2              =   5400
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line10 
         Index           =   0
         X1              =   6720
         X2              =   6720
         Y1              =   4800
         Y2              =   2400
      End
      Begin VB.Line Line9 
         Index           =   0
         X1              =   7320
         X2              =   6120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   6120
         X2              =   7320
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line7 
         Index           =   0
         X1              =   7320
         X2              =   7320
         Y1              =   4800
         Y2              =   2400
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   6120
         X2              =   6120
         Y1              =   4800
         Y2              =   2400
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   6120
         X2              =   7320
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   6120
         X2              =   7320
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   6120
         X2              =   7320
         Y1              =   4800
         Y2              =   4800
      End
   End
   Begin VB.Timer Timer2 
      Left            =   10440
      Top             =   1920
   End
   Begin VB.CommandButton Command6 
      Caption         =   "单次摇号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Picture         =   "Form1.frx":10CA
      TabIndex        =   37
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "清零"
      Height          =   495
      Left            =   12480
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0"
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   3
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   2
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   1
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   0
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "隐藏测试"
      Height          =   495
      Left            =   8760
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "显示测试"
      Height          =   495
      Left            =   8760
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10080
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1680
      TabIndex        =   0
      Top             =   7200
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "停止"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1680
      TabIndex        =   5
      Top             =   7200
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Caption         =   "摇号结果"
      Height          =   3975
      Left            =   2400
      TabIndex        =   1
      Top             =   2760
      Width           =   4815
      Begin VB.Line Line2 
         X1              =   2400
         X2              =   2400
         Y1              =   480
         Y2              =   3600
      End
      Begin VB.Line Line1 
         X1              =   840
         X2              =   4080
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   72
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   3
         Left            =   3000
         TabIndex        =   19
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   72
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   2
         Left            =   960
         TabIndex        =   18
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   72
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   1
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   72
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   0
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.Image Image3 
      Height          =   1470
      Left            =   6600
      Picture         =   "Form1.frx":4632
      Top             =   8880
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   6240
      Picture         =   "Form1.frx":5F21
      Top             =   7680
      Width           =   660
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   240
      Picture         =   "Form1.frx":93C3
      Top             =   7200
      Width           =   1485
   End
   Begin VB.Label Label12 
      Caption         =   "---习近平"
      Height          =   255
      Left            =   13920
      TabIndex        =   36
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "基础不牢, 地动山摇。"
      Height          =   375
      Left            =   11040
      TabIndex        =   35
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "---President Xi"
      Height          =   255
      Left            =   13320
      TabIndex        =   34
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Without a solid foundation,it would lead to serious vulnerabilities."
      Height          =   615
      Left            =   11040
      TabIndex        =   33
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label Label8 
      Caption         =   "Knowledge makes humble, ignorance makes proud."
      Height          =   495
      Left            =   4560
      TabIndex        =   32
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label7 
      Caption         =   "Re-Edited"
      Height          =   375
      Left            =   9120
      TabIndex        =   31
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "总次数"
      Height          =   255
      Index           =   4
      Left            =   10800
      TabIndex        =   29
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "频率     %"
      Height          =   255
      Index           =   3
      Left            =   12480
      TabIndex        =   27
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "频率     %"
      Height          =   255
      Index           =   2
      Left            =   12480
      TabIndex        =   26
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "频率     %"
      Height          =   255
      Index           =   1
      Left            =   12480
      TabIndex        =   25
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "频率     %"
      Height          =   255
      Index           =   0
      Left            =   12480
      TabIndex        =   24
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "序号4:   次数"
      Height          =   255
      Index           =   3
      Left            =   10200
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "序号3:   次数"
      Height          =   255
      Index           =   2
      Left            =   10200
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "序号2:   次数"
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "序号1:   次数"
      Height          =   255
      Index           =   0
      Left            =   10200
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "By 霍芬比 Hophenby "
      Height          =   615
      Left            =   10080
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   480
      Width           =   5415
   End
   Begin VB.Menu 组号模式 
      Caption         =   "组号模式"
   End
   Begin VB.Menu 座位模式 
      Caption         =   "座位模式"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s


Private Sub Command1_Click()
    Timer1.Enabled = True
    Command1.Visible = False
End Sub

Private Sub Command2_Click()
    For i = 0 To 3
        Label2(i).Visible = True
        Label6(i).Visible = True
        Text1(i).Visible = True
        Text2(i).Visible = True
    Next i
    Text3.Visible = True
    Command3.Visible = True
    Command5.Visible = True
    Label6(4).Visible = True
End Sub

Private Sub Command3_Click()
    For i = 0 To 3
        Label2(i).Visible = False
        Label6(i).Visible = False
        Text1(i).Visible = False
        Text2(i).Visible = False
    Next i
    Text3.Visible = False
    Command3.Visible = False
    Command5.Visible = False
    Label6(4).Visible = False


End Sub

Private Sub Command4_Click()

Timer1.Enabled = False
Timer2.Interval = 1
Timer2.Enabled = True
Command1.Enabled = 0
    Command1.Visible = True
End Sub

Private Sub Command5_Click()

    For i = 0 To 3
        Text1(i).Text = 0
        Text2(i).Text = 0
    Next i
        Text3.Text = 0
End Sub

Private Sub Command6_Click()
    r = Int(Rnd * 4)
    Text3.Text = Text3.Text + 1
    Text1(r).Text = Text1(r).Text + 1
    For i = 0 To 3
        Label1(i).Visible = False
        Text2(i).Text = Int(Text1(i).Text / Text3.Text * 100)
    Next i
        Label1(r).Visible = True
End Sub

Private Sub Command7_Click()
If Text4.Text <> 0 Then
    For i = 1 To 6
        For j = 1 To 8
            If i < 5 Or j > 2 Then
               ex(j & i).Visible = False
            End If
        Next
    Next
    For f = 1 To Text4.Text
    Do
        Do
            p = Int(Rnd * 6) + 1
            q = Int(Rnd * 8) + 1
        Loop While (p >= 5 And q <= 2)
    Loop While ex(q & p).Visible = True
        ex(q & p).Visible = True
    Next
Else
    MsgBox "人数不能为零!"
End If
End Sub

Private Sub Command8_Click()
If Text4.Text <> 0 Then
Text4.Text = Text4.Text - 1
End If
End Sub

Private Sub Command9_Click()
If Text4.Text <> 44 Then
Text4.Text = Text4.Text + 1
End If
End Sub

Private Sub Form_Load()
    Randomize
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()
    
    Do
    r = Int(Rnd * 4)
    Loop While (r = s)
    s = r
    Text3.Text = Text3.Text + 1
    Text1(r).Text = Text1(r).Text + 1
    For i = 0 To 3
        Label1(i).Visible = False
        Text2(i).Text = Int(Text1(i).Text / Text3.Text * 100)
    Next i
        Label1(r).Visible = True
End Sub

Private Sub Timer2_Timer()
    
    Do
    r = Int(Rnd * 4)
    Loop While (r = s)
    s = r
    Text3.Text = Text3.Text + 1
    Text1(r).Text = Text1(r).Text + 1
    For i = 0 To 3
        Label1(i).Visible = False
        Text2(i).Text = Int(Text1(i).Text / Text3.Text * 100)
    Next i
        Label1(r).Visible = True
    Timer2.Interval = (Timer2.Interval + 1) * 2
    If Timer2.Interval >= 1200 Then
    Timer2.Enabled = False
    Command1.Enabled = True
    End If
End Sub

Private Sub 组号模式_Click()
If Timer1.Enabled = 0 And Timer2.Enabled = 0 Then
Command2.Visible = 1
    Frame2.Visible = 1
Command6.Visible = 1
Image2.Visible = 1
Command1.Visible = 1
Image1.Visible = 1
Command4.Visible = 1
Command7.Visible = 0
Image3.Visible = 0
Frame1.Visible = 0
Frame3.Visible = 0
End If
End Sub

Private Sub 座位模式_Click()
If Timer1.Enabled = 0 And Timer2.Enabled = 0 Then
Command2.Visible = 0
    For i = 0 To 3
        Label1(i).Visible = 0
        Label2(i).Visible = 0
        Label6(i).Visible = 0
        Text1(i).Visible = 0
        Text2(i).Visible = 0
    Next i
    Text3.Visible = 0
    Command3.Visible = 0
    Command5.Visible = 0
    Label6(4).Visible = 0
    Frame2.Visible = 0
Command6.Visible = 0
Image2.Visible = 0
Command1.Visible = 0
Image1.Visible = 0
Command4.Visible = 0
Command7.Visible = 1
Image3.Visible = 1
Frame1.Visible = 1
Frame3.Visible = 1
End If
End Sub
