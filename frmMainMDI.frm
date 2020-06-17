VERSION 5.00
Begin VB.MDIForm frmMainMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Polyhedron explorer"
   ClientHeight    =   4755
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5985
   Icon            =   "frmMainMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   5925
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   5985
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Ready."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMRUFile 
         Caption         =   "mnuMRUFile"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuLanguage 
         Caption         =   "&Language"
         Begin VB.Menu mnuLanEnglish 
            Caption         =   "&English"
         End
         Begin VB.Menu mnuLanRussian 
            Caption         =   "&Russian"
         End
      End
   End
End
Attribute VB_Name = "frmMainMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'======================================================
' This module is responsible for the parent window behavior
'======================================================

Private Sub MDIForm_Load()
'======================================================
' This is executed when the application starts
'======================================================
Init
FillStrings
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'======================================================
' When the application is closed, all settings are saved and
' the instance of the Common Dialog is terminated.
'======================================================
SaveSettings
Set CD = Nothing
End
End Sub

Private Sub mnuExit_Click()
'======================================================
' Exit program request from menu
'======================================================
FileExit
End Sub

Private Sub mnuLanEnglish_Click()
'======================================================
ChangeLanguage lanEnglish
End Sub

Private Sub mnuLanRussian_Click()
'======================================================
ChangeLanguage lanRussian
End Sub

Private Sub mnuMRUFile_Click(Index As Integer)
'======================================================
' One of the six recent files is selected to be opened
'======================================================
FileMRUClick Index
End Sub

Private Sub mnuNew_Click()
'======================================================
' A new polyhedron is going to be created
'======================================================
FileNew
End Sub

Private Sub mnuOpen_Click()
'======================================================
' This command is executed when a file open command is issued from menu
'======================================================
FileOpen
End Sub

Private Sub mnuOptions_Click()
'======================================================
' Currently, all the options available are just the language settings
'======================================================
mnuLanEnglish.Checked = False
mnuLanRussian.Checked = False
Select Case setLanguage
    Case lanEnglish
        mnuLanEnglish.Checked = True
    Case lanRussian
        mnuLanRussian.Checked = True
End Select
End Sub

'=====================================================

Public Sub FillStrings()
'======================================================
' Initialializes all strings and captions of the parent window
' takes current language into account
'======================================================
Caption = GetString(Res_Caption)
mnuExit.Caption = GetString(ResMnu_Exit)
mnuFile.Caption = GetString(ResMnu_File)
mnuNew.Caption = GetString(ResMnu_New)
mnuOpen.Caption = GetString(ResMnu_Open)
mnuOptions.Caption = GetString(ResMnu_Options)
End Sub
