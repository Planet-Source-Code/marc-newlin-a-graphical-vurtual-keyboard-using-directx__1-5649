VERSION 5.00
Begin VB.Form vkb 
   BackColor       =   &H80000007&
   Caption         =   "Virtual Keyboard"
   ClientHeight    =   3930
   ClientLeft      =   585
   ClientTop       =   3705
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   13875
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   9240
      Top             =   2400
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   211
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   156
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   78
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   73
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   82
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton KEY 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   83
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   75
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   76
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   77
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   79
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   80
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   81
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   71
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Num Lock"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   69
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   181
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   72
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   74
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   55
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   200
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   205
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   208
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   203
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Page Down"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   209
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   207
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Page Up"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   201
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   199
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   210
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Space Bar"
      Height          =   495
      Index           =   57
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Ctrl"
      Height          =   495
      Index           =   157
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton KEY 
      Height          =   495
      Index           =   221
      Left            =   7440
      Picture         =   "VKB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton KEY 
      Height          =   495
      Index           =   220
      Left            =   6600
      Picture         =   "VKB.frx":0556
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Alt"
      Height          =   495
      Index           =   184
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Alt"
      Height          =   495
      Index           =   56
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton KEY 
      Height          =   495
      Index           =   219
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "VKB.frx":0AA4
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Ctrl"
      Height          =   495
      Index           =   29
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton KEY 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   53
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   45
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   46
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   47
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   48
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   49
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   50
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   51
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   44
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Shift"
      Height          =   495
      Index           =   54
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Shift"
      Height          =   495
      Index           =   42
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton KEY 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   37
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   36
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   38
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   35
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   34
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   33
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      BackColor       =   &H8000000A&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   39
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   40
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Caps Lock"
      Height          =   495
      Index           =   58
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Enter"
      Height          =   1095
      Index           =   28
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton KEY 
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   43
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "]"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "["
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Tab"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton KEY 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "`"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   41
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Pause "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   197
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Print Scrn"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   183
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "Scroll Lock"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   70
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   59
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F12"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   88
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F11"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   87
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   67
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   60
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   61
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   62
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F10"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   68
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   63
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F6"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   64
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F7"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   65
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "F8"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   66
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton KEY 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   52
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   11400
      Picture         =   "VKB.frx":0FF2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "vkb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pressed(255) As Boolean, Numb As Long, Numbstr As String, directistate As DIKEYBOARDSTATE, directidev As DirectInputDevice, directi As DirectInput, directx As New DirectX7
Private Sub Form_Load()
Set directi = directx.DirectInputCreate()
     Set directidev = directi.CreateDevice("GUID_SysKeyboard")
     directidev.SetCommonDataFormat DIFORMAT_KEYBOARD
     directidev.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
     directidev.Acquire
     Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
    directidev.GetDeviceStateKeyboard directistate
     Open "list.txt" For Input As #1
     Do While Not EOF(1)
        Line Input #1, Numbstr
        Numb = Int(Numbstr)
        CheckKey (Numb)
     Loop
     Close #1
     DoEvents
End Sub
Private Sub CheckKey(keynum As Long)
If directistate.KEY(keynum) <> 0 Then KEY(keynum).BackColor = &HFFFF& Else KEY(keynum).BackColor = &H8000000A
End Sub
