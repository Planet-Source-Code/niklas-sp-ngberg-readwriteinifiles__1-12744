VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Read/Write INI-files"
   ClientHeight    =   2235
   ClientLeft      =   3225
   ClientTop       =   4530
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2235
   ScaleWidth      =   2955
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContactName 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtContactEmail 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   2100
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "nickokick@spray.se"
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   60
      TabIndex        =   7
      Top             =   1980
      Width           =   2835
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Niklas Spångberg"
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   60
      TabIndex        =   6
      Top             =   1740
      Width           =   2835
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Email :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'  Terms of Agreement:

'  By using this code, you agree to the following terms:

' 1) You may use this code in your own programs and may
'    compile it into an .exe/.dll/.ocx and distribute it
'    in binary format freely and with no charge.

' 2) You MAY NOT redistribute this code (for example to a
'    web site) without written permission from the
'    original author. Failure to do so is a violation of
'    copyright laws.

' 3) You may link to this code from another website, but
'    ONLY if it is not wrapped in a frame.

'  This code is copywrited by Niklas Spångberg 2000
'  Email: nickokick@spray.se

Private Sub cmdLoad_Click()

    'Load ContactEmail
    txtContactEmail.Text = ReadIni("Email", "ContactEmail")

    'Load ContactName
    txtContactName.Text = ReadIni("Name", "ContactName")
    
End Sub

Private Sub cmdSave_Click()
    Dim x As Boolean
    
    ' Save ContactEmail
    x = WriteIni("Email", "ContactEmail", txtContactEmail.Text)

    ' Save ContactName
    x = WriteIni("Name", "ContactName", txtContactName.Text)
    
End Sub
