VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Scientific Calculator"
   ClientHeight    =   8850
   ClientLeft      =   4965
   ClientTop       =   2940
   ClientWidth     =   13725
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   13725
   Begin VB.CommandButton Command63 
      BackColor       =   &H008080FF&
      Caption         =   "w"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   36
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command62 
      BackColor       =   &H00FFFF00&
      Caption         =   "log a (x)"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command61 
      BackColor       =   &H0080FFFF&
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton Command60 
      BackColor       =   &H00FF8080&
      Caption         =   "ncr"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command59 
      BackColor       =   &H00FFFF00&
      Caption         =   "npr"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command58 
      BackColor       =   &H00FF8080&
      Caption         =   "coth-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command57 
      BackColor       =   &H00FFFF00&
      Caption         =   "sech-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command56 
      BackColor       =   &H00FF8080&
      Caption         =   "cosech-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command55 
      BackColor       =   &H00FFFF00&
      Caption         =   "tanh-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command54 
      BackColor       =   &H00FF8080&
      Caption         =   "cosh-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command53 
      BackColor       =   &H00FFFF00&
      Caption         =   "sinh-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command52 
      BackColor       =   &H00FF8080&
      Caption         =   "coth"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command51 
      BackColor       =   &H00FFFF00&
      Caption         =   "sech"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command50 
      BackColor       =   &H00FF8080&
      Caption         =   "cosech"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command49 
      BackColor       =   &H00FFFF00&
      Caption         =   "tanh"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command48 
      BackColor       =   &H00FF8080&
      Caption         =   "cosh"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command47 
      BackColor       =   &H00FFFF00&
      Caption         =   "sinh"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command46 
      BackColor       =   &H00FF8080&
      Caption         =   "Eng"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command45 
      BackColor       =   &H0080C0FF&
      Caption         =   "[ ]"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00FFFF00&
      Caption         =   "cot-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command43 
      BackColor       =   &H00FF8080&
      Caption         =   "sec-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command42 
      BackColor       =   &H0080C0FF&
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H00FFFF00&
      Caption         =   "cosec-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command40 
      BackColor       =   &H008080FF&
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   20.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command39 
      BackColor       =   &H00FFFF00&
      Caption         =   "10^x"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command38 
      BackColor       =   &H00404040&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H00FFFF00&
      Caption         =   "cos-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H00FF8080&
      Caption         =   "sin-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H00FF8080&
      Caption         =   "tan-1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H008080FF&
      Caption         =   " ! "
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00FF8080&
      Caption         =   "Mod"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00FF8080&
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00FFFF00&
      Caption         =   "cot"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00FF8080&
      Caption         =   "sec"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00FFFF00&
      Caption         =   "cosec"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H008080FF&
      Caption         =   "e^x"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H0080C0FF&
      Caption         =   "Neg"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H008080FF&
      Caption         =   "sgn"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00FF8080&
      Caption         =   "log"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H0080C0FF&
      Caption         =   "ln"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00FFFF00&
      Caption         =   "x^(1/2)"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00FF8080&
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00FF8080&
      Caption         =   "x^n"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FFFF00&
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FFFF00&
      Caption         =   "x^2"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FF8080&
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00404080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6840
      Width           =   4215
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00404040&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00404040&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00404040&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00404040&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Width           =   4215
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00808080&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00808080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00808080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00808080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   13335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   13335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Sayantan Saha"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   3015
   End
   Begin VB.Menu Calculator 
      Caption         =   "Calculator"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Dim i As Integer
Dim j As Integer
Dim b As Double
Dim op As Integer
Dim fn As Double
Dim r As Double
Private Sub Command1_Click()
Text2.Text = Text2.Text & 7
End Sub

Private Sub Command10_Click()
Text2.Text = Text2.Text & 0
End Sub

Private Sub Command11_Click()
Text2.Text = Text2.Text & "."
End Sub

Private Sub Command12_Click()
Text2.Text = " "
Text1.Text = " "
End Sub

Private Sub Command13_Click()
op = 1
Text1.Text = Text2.Text & "+"
Text2.Text = " "
End Sub

Private Sub Command14_Click()
op = 2
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command15_Click()
op = 3
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command16_Click()
op = 4
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command17_Click()
If Text2.Text = " " Then
Text1.Text = "Invalid Input"
ElseIf op = 1 Then
a = Text2.Text
b = Text1.Text
Text1.Text = a + b
ElseIf op = 2 Then
a = Text2.Text
b = Text1.Text
Text1.Text = b - a
ElseIf op = 3 Then
a = Text2.Text
b = Text1.Text
Text1.Text = a * b
ElseIf op = 4 Then
a = Text2.Text
b = Text1.Text
If Text2.Text = 0 Then
Text1.Text = "Infinity"
Else
Text1.Text = b / a
End If
ElseIf op = 5 Then
If Text2.Text = " " Then
Text1.Text = "Invalid Input"
Else
Text1.Text = Math.Sin(Text2.Text * 3.14159265358979 / 180)
Text1.Text = Format(Text1.Text, "0.000")
End If
ElseIf op = 6 Then
If Text2.Text = " " Then
Text1.Text = "Invalid Input"
Else
Text1.Text = Math.Cos(Text2.Text * 3.14159265358979 / 180)
Text1.Text = Format(Text1.Text, "0.000")
End If
ElseIf op = 7 Then
If Text2.Text = " " Then
Text1.Text = "Invalid Input"
Else
r = (Text2.Text / 90)
If (r - Int(r)) = 0 And r Mod 2 <> 0 Then
Text1.Text = "Infinity"
Else
Text1.Text = Math.Tan(Text2.Text * 3.14159265358979 / 180)
Text1.Text = Format(Text1.Text, "0.000")
End If
End If
ElseIf op = 8 Then
If Text2.Text = " " Then
Text1.Text = "Invalid Input"
ElseIf Text2.Text <= 0 Then
Text1.Text = "Math Error"
Else
Text1.Text = Math.Log(Text2.Text) / 2.30258509299404
End If
ElseIf op = 9 Then
If Text2.Text <= 0 Then
Text1.Text = "Math Error"
Else
Text1.Text = Math.Log(Text2.Text)
End If
ElseIf op = 10 Then
Text1.Text = (Text2.Text * Text2.Text)
ElseIf op = 11 Then
a = Text2.Text
b = Text1.Text
Text1.Text = b ^ a
ElseIf op = 12 Then
Text1.Text = Math.Sqr(Text2.Text)
ElseIf op = 13 Then
Text1.Text = Math.Sgn(Text2.Text)
ElseIf op = 14 Then
Text1.Text = (0 - Text2.Text)
ElseIf op = 15 Then
Text1.Text = Math.Exp(Text2.Text)
ElseIf op = 16 Then
r = Text2.Text / 180
If Text2.Text = 0 Then
Text1.Text = "Infinity"
ElseIf (r - Int(r)) = 0 Then
Text1.Text = "Infinity"
Else
Text1.Text = 1 / Math.Sin(Text2.Text * 3.14159265358979 / 180)
End If
ElseIf op = 17 Then
r = Text2.Text / 90
If Text2.Text = 90 Then
Text1.Text = "Infinity"
ElseIf (r - Int(r)) = 0 And r Mod 2 <> 0 Then
Text1.Text = "Infinity"
Else
Text1.Text = 1 / Math.Cos(Text2.Text * 3.14159265358979 / 180)
End If
ElseIf op = 18 Then
r = Text2.Text / 180
If (r - Int(r)) = 0 Then
Text1.Text = "Infinity"
ElseIf (r * 2 - Int(r * 2)) = 0 And r * 2 Mod 2 <> 0 Then
Text1.Text = 0
Else: Text1.Text = 1 / Math.Tan(Text2.Text * 3.14159265358979 / 180)
End If
ElseIf op = 19 Then
If Text2.Text = 0 Then
Text1.Text = "Infinity"
Else
Text1.Text = 1 / Text2.Text
End If
ElseIf op = 20 Then
If Text2.Text = 0 Then
Text1.Text = 1
Else
For i = 1 To (Text2.Text - 1)
Text2.Text = Text2.Text * i
Text1.Text = Text2.Text
Next
End If
ElseIf op = 22 Then
Text1.Text = Math.Abs(Text2.Text)
ElseIf op = 21 Then
Text1.Text = Math.Atn(Text2.Text) / 3.14159265358979 * 180
ElseIf op = 23 Then
a = Text2.Text
If Text2.Text = 1 Then
Text1.Text = 90
ElseIf Text2.Text > 1 Then
Text1.Text = "Math Error"
Else
Text1.Text = Math.Atn(a / (Math.Sqr(1 - (a * a)))) / 3.14159265358979 * 180
End If
ElseIf op = 24 Then
a = Text2.Text
If Text2.Text = 0 Then
Text1.Text = 90
ElseIf Text2.Text > 1 Then
Text1.Text = "Math Error"
Else
Text1.Text = Math.Atn((Math.Sqr(1 - (a * a))) / a) / 3.14159265358979 * 180
End If
ElseIf op = 25 Then
a = Text2.Text
b = Text1.Text
Text1.Text = b / a * 100
ElseIf op = 28 Then
If Text2.Text = 0 Then
Text1.Text = "Math Error"
Else
a = 1 / Text2.Text
If Text2.Text = 1 Then
Text1.Text = 90
ElseIf Text2.Text < 1 Then
Text1.Text = "Math Error"
Else
Text1.Text = Math.Atn(a / (Math.Sqr(1 - (a * a)))) / 3.14159265358979 * 180
End If
End If
ElseIf op = 29 Then
If Text2.Text = 0 Then
Text1.Text = "Math Error"
Else
a = 1 / Text2.Text
If Text2.Text = 1 Then
Text1.Text = 0
ElseIf Text2.Text < 1 Then
Text1.Text = "Math Error"
Else
Text1.Text = Math.Atn((Math.Sqr(1 - (a * a))) / a) / 3.14159265358979 * 180
End If
End If
ElseIf op = 30 Then
If Text2.Text = 0 Then
Text1.Text = 90
Else
a = 1 / Text2.Text
Text1.Text = Math.Atn(a) / 3.14159265358979 * 180
End If
ElseIf op = 31 Then
Text1.Text = 10 ^ Text2.Text
ElseIf op = 32 Then
Text1.Text = Int(Text2.Text)
ElseIf op = 33 Then
Text1.Text = Text2.Text / 1000 & " * 10^3"
ElseIf op = 34 Then
Text1.Text = (2.71828182845905 ^ (Text2.Text) - 2.71828182845905 ^ (-Text2.Text)) / 2
ElseIf op = 35 Then
Text1.Text = (2.71828182845905 ^ (Text2.Text) + 2.71828182845905 ^ (-Text2.Text)) / 2
ElseIf op = 36 Then
a = (2.71828182845905 ^ (Text2.Text) - 2.71828182845905 ^ (-Text2.Text))
b = (2.71828182845905 ^ (Text2.Text) + 2.71828182845905 ^ (-Text2.Text))
Text1.Text = a / b
ElseIf op = 37 Then
Text1.Text = 1 / ((2.71828182845905 ^ (Text2.Text) - 2.71828182845905 ^ (-Text2.Text)) / 2)
ElseIf op = 38 Then
Text1.Text = 1 / ((2.71828182845905 ^ (Text2.Text) + 2.71828182845905 ^ (-Text2.Text)) / 2)
ElseIf op = 39 Then
a = (2.71828182845905 ^ (Text2.Text) - 2.71828182845905 ^ (-Text2.Text))
b = (2.71828182845905 ^ (Text2.Text) + 2.71828182845905 ^ (-Text2.Text))
Text1.Text = b / a
ElseIf op = 40 Then
Text1.Text = Math.Log(Text2.Text + Math.Sqr(Text2.Text ^ 2 + 1))
ElseIf op = 41 Then
Text1.Text = Math.Log(Text2.Text + Math.Sqr(Text2.Text ^ 2 - 1))
ElseIf op = 42 Then
If Text2.Text = 1 Then
Text1.Text = "Infinity"
ElseIf Text2.Text > 1 Then
Text1.Text = "Math Error"
Else
Text1.Text = Math.Log((1 + Text2.Text) / (1 - Text2.Text)) / 2
End If
ElseIf op = 43 Then
Text1.Text = 1 / Math.Log(Text2.Text + Math.Sqr(Text2.Text ^ 2 + 1))
ElseIf op = 44 Then
If Text2.Text <= 1 Then
Text1.Text = "Math Error"
Else
Text1.Text = 1 / Math.Log(Text2.Text + Math.Sqr(Text2.Text ^ 2 - 1))
End If
ElseIf op = 45 Then
If Text2.Text >= 1 Then
Text1.Text = "Math Error"
Else
Text1.Text = Math.Log((1 - Text2.Text) / (1 + Text2.Text)) / 2
End If
ElseIf op = 46 Then
j = (Text1.Text - Text2.Text)
If j < 0 Then
Text1.Text = "Math Error"
ElseIf j = 0 Then
For i = 1 To Text1.Text - 1
Text1.Text = Text1.Text * i
Next
ElseIf Text1.Text < 1 Then
Text1.Text = "Math Error"
ElseIf Text2.Text < 1 And Text2.Text <> 0 Then
Text1.Text = "Math Error"
ElseIf Text1.Text >= 1 And Text2.Text = 0 Then
Text1.Text = 1
ElseIf Text1.Text = 0 And Text2.Text = 0 Then
Text1.Text = 1
ElseIf Text1.Text >= 1 And Text2.Text >= 1 Then
For i = 1 To (Text1.Text - 1)
Text1.Text = Text1.Text * i
Next
For i = 1 To (j - 1)
j = j * i
Next
Text1.Text = Text1.Text / j
End If
ElseIf op = 47 Then
If Text2.Text < 1 And Text2.Text <> 0 Then
Text1.Text = "Math Error"
ElseIf Text1.Text < 1 Then
Text1.Text = "Math Error"
ElseIf Text2.Text = 0 Then
Text1.Text = 1
Else
j = (Text1.Text - Text2.Text)
If j < 0 Then
Text1.Text = "Math Error"
ElseIf j = 0 Then
Text1.Text = 1
Else
For i = 1 To (Text1.Text - 1)
Text1.Text = Text1.Text * i
Next
For i = 1 To (Text2.Text - 1)
Text2.Text = Text2.Text * i
Next
For i = 1 To (j - 1)
j = j * i
Next
Text1.Text = Text1.Text / Text2.Text / j
End If
End If
ElseIf op = 48 Then
Text1.Text = Text2.Text
ElseIf op = 49 Then
If Text2.Text = 1 Or Text2.Text = 0 Then
Text1.Text = "Math Error"
ElseIf Text1.Text = 0 Then
Text1.Text = "Math Error"
Else
Text1.Text = Math.Log(Text1.Text) / Math.Log(Text2.Text)
End If
End If
End Sub

Private Sub Command18_Click()
op = 5
Text2.Text = Text2.Text & "sin"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command19_Click()
op = 10
End Sub

Private Sub Command2_Click()
Text2.Text = Text2.Text & 8
End Sub

Private Sub Command20_Click()
op = 6
Text2.Text = Text2.Text & "cos"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command21_Click()
op = 11
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command22_Click()
op = 7
Text2.Text = Text2.Text & "tan"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command23_Click()
op = 12
End Sub

Private Sub Command24_Click()
op = 9
Text2.Text = Text2.Text & "ln"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command25_Click()
op = 8
Text2.Text = Text2.Text & "log"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command26_Click()
op = 13
Text2.Text = Text2.Text & "sgn"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command27_Click()
op = 14
End Sub

Private Sub Command28_Click()
op = 15
Text1.Text = Text2.Text & "e^"
Text2.Text = " "
End Sub

Private Sub Command29_Click()
op = 16
Text2.Text = Text2.Text & "cosec"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command3_Click()
Text2.Text = Text2.Text & 9
End Sub

Private Sub Command30_Click()
op = 17
Text2.Text = Text2.Text & "sec"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command31_Click()
op = 18
Text2.Text = Text2.Text & "cot"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command32_Click()
op = 19
End Sub

Private Sub Command33_Click()
op = 22
Text1.Text = "|" & Text2.Text & "|"
End Sub

Private Sub Command34_Click()
op = 20
End Sub

Private Sub Command35_Click()
op = 21
Text2.Text = Text2.Text & "tan-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command36_Click()
op = 23
Text2.Text = Text2.Text & "sin-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub
Private Sub Command37_Click()
op = 24
Text2.Text = Text2.Text & "cos-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command38_Click()
op = 25
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command39_Click()
op = 31
Text1.Text = Text2.Text & "10^"
Text2.Text = " "
End Sub

Private Sub Command4_Click()
Text2.Text = Text2.Text & 4
End Sub

Private Sub Command40_Click()
Text2.Text = Text2.Text & 3.14159265358979
End Sub

Private Sub Command41_Click()
op = 28
Text2.Text = Text2.Text & "cosec-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command42_Click()
Text2.Text = Text2.Text & 2.71828182845905
End Sub

Private Sub Command43_Click()
op = 29
Text2.Text = Text2.Text & "sec-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command44_Click()
op = 30
Text2.Text = Text2.Text & "cot-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command45_Click()
op = 32
End Sub

Private Sub Command46_Click()
op = 33
End Sub

Private Sub Command47_Click()
op = 34
Text2.Text = Text2.Text & "sinh"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command48_Click()
op = 35
Text2.Text = Text2.Text & "cosh"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command49_Click()
op = 36
Text2.Text = Text2.Text & "tanh"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command5_Click()
Text2.Text = Text2.Text & 5
End Sub

Private Sub Command50_Click()
op = 37
Text2.Text = Text2.Text & "cosech"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command51_Click()
op = 38
Text2.Text = Text2.Text & "sech"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command52_Click()
op = 39
Text2.Text = Text2.Text & "coth"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command53_Click()
op = 40
Text2.Text = Text2.Text & "sinh-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command54_Click()
op = 41
Text2.Text = Text2.Text & "cosh-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command55_Click()
op = 42
Text2.Text = Text2.Text & "tanh-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command56_Click()
op = 43
Text2.Text = Text2.Text & "cosech-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command57_Click()
op = 44
Text2.Text = Text2.Text & "sech-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command58_Click()
op = 45
Text2.Text = Text2.Text & "coth-1"
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command59_Click()
op = 46
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command6_Click()
Text2.Text = Text2.Text & 6
End Sub

Private Sub Command60_Click()
op = 47
Text2.Text = Text2.Text
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command61_Click()
op = 48
If Text1.Text = " " Then
Text2.Text = fn
Else
fn = Text1.Text
End If
End Sub

Private Sub Command62_Click()
op = 49
Text1.Text = Text2.Text
Text2.Text = " "
End Sub

Private Sub Command63_Click()
Text1.Text = "(-1/2)+[i/2*(3)^0.5] " & ", (-1/2)-[i/2*(3)^0.5]"
End Sub

Private Sub Command7_Click()
Text2.Text = Text2.Text & 1
End Sub

Private Sub Command8_Click()
Text2.Text = Text2.Text & 2
End Sub

Private Sub Command9_Click()
Text2.Text = Text2.Text & 3
End Sub
