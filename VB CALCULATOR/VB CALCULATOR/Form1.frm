VERSION 5.00
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin Project1.PictureG PictureG1 
      Height          =   9000
      Left            =   0
      Top             =   0
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   15875
      GIF             =   "Form1.frx":0000
      Stretch         =   2
      DelayLoad       =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
Form2.Show
End Sub
