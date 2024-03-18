VERSION 5.00
Begin VB.Form frmShowProcess 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmShowProcess"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstCaptions 
      Height          =   5325
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.ListBox lstProcess 
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblProceso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmShowProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

