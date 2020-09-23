VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00DEDAC5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clone Vb"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00DEDAC5&
      Caption         =   "Make sure of this please "
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5895
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDAC5&
         Height          =   1395
         ItemData        =   "Form1.frx":030A
         Left            =   3720
         List            =   "Form1.frx":0329
         TabIndex        =   7
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":038D
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000040C0&
         X1              =   120
         X2              =   5760
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":0480
         ForeColor       =   &H00404040&
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDAC5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   3840
      Pattern         =   "*.frm;*.vbp"
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDAC5&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DEDAC5&
      Caption         =   "Make  project.exe"
      Height          =   465
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File will be compiled:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDAC5&
      BackStyle       =   0  'Transparent
      Caption         =   "My vb clone(Project to Executable compiler)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'We declare our main vb and our main project file

Dim vb
Dim prj

vb = Dir("VB6.EXE")                                 'vb6.exe  location
If vb <> "VB6.EXE" Then
MsgBox "VB not found", vbInformation, "There is no vb here"
Exit Sub

End If

If vb = "VB6.EXE" Then                            'If vb6.exe found in our app.path

MsgBox "Vb6 found  press ok to compile to exe", vbInformation, "Great !"

End If

prj = Text1.Text                                      'Default project path
If Len(Text1.Text) < 2 Then
MsgBox "You didnt define your default project file", vbInformation, "ERROR"
Exit Sub
Else

Shell vb & " /make " & prj, vbHide      'Compile to exe

MsgBox "Finished", vbInformation, "Job done"
End
End If

End Sub

Private Sub File1_Click()
Text1.Text = File1.FileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Do not forget to vote me if you like this project", vbInformation, "An honesty side"

End
End Sub
