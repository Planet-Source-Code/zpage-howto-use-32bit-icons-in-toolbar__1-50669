VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "HowTo: Use 32Bit Icons in ToolBar"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Oben ausrichten
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   1111
      ButtonWidth     =   1217
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Load"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Save"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Settings"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Help"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright © 2oo3, Zpage[Myst]"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":57E2
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin ComctlLib.ImageList imlToolbarIcons 
      Left            =   120
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5956
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":634C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":669E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":69F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6D42
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Initialize()
  InitCommonControls
End Sub
'The things above are for the XP Style

'That are the commands for the Toolbar
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
On Error Resume Next
    Select Case Button.Index
        Case 1
            MsgBox "Please add code for Button 'New'"
        Case 2
            MsgBox "Please add code for Button 'Load'"
        Case 3
            MsgBox "Please add code for Button 'Save'"
        Case 4
            MsgBox "Please add code for Button 'Settings'"
        Case 5
            MsgBox "Please add code for Button 'Help'"
        Case 6
            MsgBox "HowTo: Use 32Bit Icons in ToolBar" & vbNewLine & "Copyright © 2oo3, Zpage[Myst]", vbInformation, "About"
        Case 7
            End
    End Select
End Sub
