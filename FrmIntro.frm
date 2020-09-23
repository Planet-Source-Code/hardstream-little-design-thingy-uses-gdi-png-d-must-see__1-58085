VERSION 5.00
Begin VB.Form FrmIntro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intro"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNo 
      Caption         =   "Cancel, I'm in a hurry!"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes, I have time, so continue"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "USES A LOT OF CPU POWER WHEN U DRAG THE FORM!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   4440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "©Copyright HardStream Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   3180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Double click to close PNG form!"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2265
   End
   Begin VB.Label lblIntro 
      AutoSize        =   -1  'True
      Caption         =   "To load a PNG form can take a while on some computers!"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label lblIntro 
      Caption         =   $"FrmIntro.frx":0000
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNo_Click()
End
End Sub

Private Sub cmdYes_Click()
Unload Me
FrmSplash.Show
End Sub
