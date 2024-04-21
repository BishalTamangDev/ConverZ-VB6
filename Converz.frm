VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ConverZ | Noobgrammer"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15240
   FillStyle       =   2  'Horizontal Line
   Icon            =   "Converz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   2  'Point
   ScaleWidth      =   762
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame EditFrame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "EditFrame"
      Height          =   7935
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   15255
      Begin VB.OptionButton EditNarratorZira 
         BackColor       =   &H8000000E&
         Caption         =   "Zira"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   31
         Top             =   5400
         Width           =   1215
      End
      Begin VB.OptionButton EditNarratorDavid 
         BackColor       =   &H8000000E&
         Caption         =   "David"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   30
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton EditAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   29
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton EditUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   28
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox EditSpecificView 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   27
         Text            =   "Converz.frx":34947
         Top             =   6000
         Width           =   8415
      End
      Begin VB.CommandButton EditDeleteFile 
         Caption         =   "Delete File"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   22
         Top             =   7200
         Width           =   4335
      End
      Begin VB.TextBox EditConversation 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   840
         Width           =   10575
      End
      Begin VB.CommandButton EditDeleteDialogue 
         Caption         =   "Delete Dialogue"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   17
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CommandButton EditExchangeCharacter 
         Caption         =   "Exchange Character"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   16
         Top             =   5520
         Width           =   2055
      End
      Begin VB.FileListBox EditFileList 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6210
         Left            =   10800
         Pattern         =   "*.converz*"
         TabIndex        =   15
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label DialogueSerialLablel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "sn"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   8640
         TabIndex        =   35
         Top             =   6840
         Width           =   225
      End
      Begin VB.Label EditClearScreen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clear Screen"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   7200
         TabIndex        =   34
         Top             =   5520
         Width           =   1275
      End
      Begin VB.Label EditFilenameLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxx"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1680
         TabIndex        =   32
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narator"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   5520
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a file"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   12240
         TabIndex        =   20
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename : "
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1395
      End
      Begin VB.Image Image13 
         Height          =   615
         Left            =   120
         Picture         =   "Converz.frx":3494D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   10575
      End
      Begin VB.Image Image12 
         Height          =   615
         Left            =   10800
         Picture         =   "Converz.frx":34AD5
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4335
      End
      Begin VB.Image EditForward 
         Height          =   360
         Left            =   10200
         Picture         =   "Converz.frx":34C4B
         Top             =   7440
         Width           =   360
      End
      Begin VB.Image EditPlay 
         Height          =   375
         Left            =   9600
         Picture         =   "Converz.frx":34FBF
         Stretch         =   -1  'True
         Top             =   7440
         Width           =   255
      End
      Begin VB.Image EditBackward 
         Height          =   360
         Left            =   8760
         Picture         =   "Converz.frx":3523D
         Top             =   7440
         Width           =   375
      End
   End
   Begin VB.Frame ViewFrame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "View Frame"
      Height          =   7935
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   15255
      Begin VB.CommandButton ViewDialog 
         Caption         =   "Dialogue"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   14
         Top             =   7320
         Width           =   1215
      End
      Begin VB.TextBox ViewConversation 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Text            =   "Converz.frx":355CF
         Top             =   840
         Width           =   9255
      End
      Begin VB.FileListBox ViewFileList 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   10800
         Pattern         =   "*.converz*"
         TabIndex        =   10
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label ViewFilenameLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxx"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   3000
         TabIndex        =   33
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename : "
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   1395
      End
      Begin VB.Image Image8 
         Height          =   615
         Left            =   1440
         Picture         =   "Converz.frx":355D5
         Stretch         =   -1  'True
         Top             =   120
         Width           =   9255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a file"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   12240
         TabIndex        =   11
         Top             =   240
         Width           =   1470
      End
      Begin VB.Image Image7 
         Height          =   615
         Left            =   10800
         Picture         =   "Converz.frx":3575D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4335
      End
      Begin VB.Image Image6 
         Height          =   1200
         Left            =   120
         Picture         =   "Converz.frx":358D3
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Image Image4 
         Height          =   1200
         Left            =   120
         Picture         =   "Converz.frx":3BC50
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame AddFrame 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Add Frame"
      Height          =   7935
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   15255
      Begin VB.TextBox AddPreview 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   9360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Text            =   "Converz.frx":406BD
         Top             =   840
         Width           =   5775
      End
      Begin VB.OptionButton OptionZira 
         BackColor       =   &H8000000E&
         Caption         =   "Zira"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   23
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton ResetButton 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   26
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton ClearScreenButton 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   25
         Top             =   5280
         Width           =   1215
      End
      Begin VB.OptionButton OptionDavid 
         BackColor       =   &H8000000E&
         Caption         =   "David"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   24
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   9
         Top             =   6960
         Width           =   1215
      End
      Begin VB.TextBox AddText1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "Converz.frx":406C3
         Top             =   120
         Width           =   7815
      End
      Begin VB.CommandButton StartConversation 
         Caption         =   "Dialogue"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   4
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Image AddPlay 
         Height          =   270
         Left            =   600
         Picture         =   "Converz.frx":406C9
         Stretch         =   -1  'True
         Top             =   6480
         Width           =   285
      End
      Begin VB.Image AddNext 
         Height          =   360
         Left            =   960
         Picture         =   "Converz.frx":40947
         Top             =   6480
         Width           =   360
      End
      Begin VB.Image AddPrevious 
         Height          =   360
         Left            =   120
         Picture         =   "Converz.frx":40CBB
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a character"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   7440
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   11760
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   9360
         Picture         =   "Converz.frx":4104D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   5775
      End
      Begin VB.Image AddSpeakerZira 
         Height          =   1200
         Left            =   120
         Picture         =   "Converz.frx":411C3
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Image AddSpeakerDavid 
         Height          =   1200
         Left            =   120
         Picture         =   "Converz.frx":47540
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ConverZ"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   21
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image EditImage 
      Height          =   675
      Left            =   13680
      Picture         =   "Converz.frx":4BFAD
      Top             =   0
      Width           =   1350
   End
   Begin VB.Image ViewImage 
      Height          =   675
      Left            =   12480
      Picture         =   "Converz.frx":4FC79
      Top             =   0
      Width           =   1350
   End
   Begin VB.Image AddImage 
      Height          =   675
      Left            =   11280
      Picture         =   "Converz.frx":53A52
      Top             =   0
      Width           =   1350
   End
   Begin VB.Image Image5 
      Height          =   735
      Left            =   0
      Picture         =   "Converz.frx":57CC3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim response As Boolean
Dim AllowDelete As Boolean

Dim i As Integer
Dim j As Integer
Dim res As Integer
Dim form_count As Integer
Dim temp_count As Integer
Dim serial As Integer

Dim filename As String
Dim character As String * 6
Dim speaker_1 As String * 6
Dim speaker_2 As String * 6

Dim temp_info As info
Dim last_info As info
Dim information As info


Private Sub EditClearScreen_Click()
    EditSpecificView.Text = ""
End Sub


Private Sub EditConversation_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


'clear
Private Sub ClearScreenButton_Click()
    AddText1.Text = ""
End Sub


Private Sub AddPreview_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub ViewConversation_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub ResetButton_Click()
    On Error GoTo errorhandler
        Kill "temp.txt"
        Open "temp.txt" For Random As #1
        Close #1
        AddPreview.Text = ""
errorhandler:
    Close #1
    Exit Sub
End Sub


Private Sub Form_Terminate()
    On Error GoTo errorhandler
    Kill "temp.txt"
    Exit Sub
    
errorhandler:
    Exit Sub
End Sub


Private Sub EditPlay_Click()
    If EditSpecificView.Text <> "" Then
        If EditNarratorDavid.Value = True Then
            Call TrySpeech(EditSpecificView.Text, "David")
        ElseIf EditNarratorZira.Value = True Then
            Call TrySpeech(EditSpecificView.Text, "Zira")
        Else
            res = MsgBox("          Please choose the narrator.", vbOKOnly + vbInformation, "Narrator")
        End If
    Else
        res = MsgBox("          Please enter something first.", vbOKOnly + vbInformation, "Empty textbox")
    End If
End Sub


Private Sub AddPlay_Click()
    If AddText1.Text <> "" Then
        If OptionDavid.Value = True Then
            Call TrySpeech(AddText1, "David")
        Else
            Call TrySpeech(AddText1, "Zira")
        End If
    Else
        res = MsgBox("          Please enter something first.", vbOKOnly + vbInformation, "Empty textbox")
    End If
End Sub


'forward button press
Private Sub AddNext_Click()
    If AddText1.Text <> "" Then
        If OptionDavid.Value = True Then
            information.character = "David"
        Else
            information.character = "Zira"
        End If
    
        information.speech = AddText1.Text
        
        temp_info.speech = information.speech
        temp_info.character = information.character
        
        Open "temp.txt" For Random As #3 Len = 106
            form_count = speech_count("temp.txt", temp_info)
            Put #3, form_count + 1, information
        Close #3
        
        'update preview section
        AddPreview.Text = ""
        
        For i = 1 To form_count + 1 Step 1
            Open "temp.txt" For Random As #4 Len = 106
                Get #4, i, information
                AddPreview.Text = AddPreview.Text + information.character + " : " + information.speech + vbNewLine
            Close #4
        Next i
        
        AddText1.Text = ""
        
        'swap characters
        If OptionDavid.Value = True Then
            OptionDavid.Value = False
            OptionZira.Value = True
        Else
            OptionDavid.Value = True
            OptionZira.Value = False
        End If
    End If
End Sub


Private Sub AddPrevious_Click()
    'count number of speech present in temporary file
    form_count = speech_count("temp.txt", temp_info)
    
    If form_count = 0 Then
        AddText1.Text = ""
    Else
        'copy all the contents in new file except the last info
        Open "temp.txt" For Random As #3 Len = 106
        Open "dup.txt" For Random As #4 Len = 106
            Get #3, form_count, last_info
            For i = 1 To form_count - 1 Step 1
                Get #3, i, temp_info
                Put #4, i, temp_info
            Next i
        Close #3
        Close #4
        
        'rename "dup.txt" as "temp.txt"
        Kill "temp.txt"
        Name "dup.txt" As "temp.txt"
        
        'update preview screen
        AddPreview.Text = ""
        
        form_count = speech_count("temp.txt", information)
        For i = 1 To form_count + 1 Step 1
            Open "temp.txt" For Random As #5 Len = 106
                Get #5, i, information
                AddPreview.Text = AddPreview.Text + information.character + " : " + information.speech + vbNewLine
            Close #5
        Next i
        
        'update add screen
        AddText1.Text = last_info.speech
        
        'swap characters based on previously deleted character
        If last_info.character = speaker_1 Then
            OptionDavid.Value = True
            OptionZira.Value = False
        Else
            OptionDavid.Value = False
            OptionZira.Value = True
        End If
    End If
End Sub


Private Sub AddText1_KeyPress(KeyAscii As Integer)
    If AddText1.Text <> "" And KeyAscii = 13 Then
        'AddText1.Text = " "
        KeyAscii = 0
        If OptionDavid.Value = True Then
            information.character = "David"
        Else
            information.character = "Zira"
        End If
    
        information.speech = AddText1.Text
        
        temp_info.speech = information.speech
        temp_info.character = information.character
        
        Open "temp.txt" For Random As #3 Len = 106
            form_count = speech_count("temp.txt", temp_info)
            Put #3, form_count + 1, information
        Close #3
        
        'update preview section
        AddPreview.Text = ""
        For i = 1 To form_count + 1 Step 1
            Open "temp.txt" For Random As #4 Len = 106
                Get #4, i, information
                AddPreview.Text = AddPreview.Text + information.character + " : " + information.speech + vbNewLine
            Close #4
        Next i
        
        AddText1.Text = ""
        
        'swap characters
        If OptionDavid.Value = True Then
            OptionDavid.Value = False
            OptionZira.Value = True
        Else
            OptionDavid.Value = True
            OptionZira.Value = False
        End If
    End If
End Sub


Private Sub EditAdd_Click()
    If EditSpecificView.Text <> "" And AllowDelete = True Then
        If EditNarratorDavid.Value = False And EditNarratorZira.Value = False Then
            response = MsgBox("          Please choose the narrator first.", vbOKOnly + vbInformation, "Unundentified Narrator")
            Exit Sub
        ElseIf EditNarratorDavid.Value = True Then
            information.character = speaker_1
        Else
            information.character = speaker_2
        End If
        
        information.speech = EditSpecificView.Text
        
        Open filename For Random As #3 Len = 106
            form_count = speech_count(filename, temp_info)
            Put #3, form_count + 1, information
        Close #3
        
        'update textbox values
        Open filename For Random As #4 Len = 106
            form_count = speech_count(filename, information)
            EditConversation.Text = ""
            For i = 1 To form_count Step 1
                Get #4, i, information
                EditConversation.Text = EditConversation.Text & i & " " + information.character + " : " + information.speech + vbNewLine
            Next i
        Close #4
        
        Open filename For Random As #5 Len = 106
            If temp_count > 0 Then
                Get #5, temp_count, information
            Else
                Get #5, 1, information
            End If
        Close #5
        
        'check option
        If information.character = speaker_1 Then
            EditNarratorDavid.Value = True
        Else
            EditNarratorZira.Value = True
        End If
            
        EditSpecificView.Text = information.speech
    ElseIf AllowDelete = False Then
        res = MsgBox("          Please choose the file first.", vbOKOnly + vbInformation, "No file chosen!")
    End If
End Sub


'edit frame -> backward move
Private Sub EditBackward_Click()
    temp_count = temp_count - 1
    
    If temp_count > 0 And temp_count < form_count And AllowDelete = True Then
        Open filename For Random As #3 Len = 106
            Get #3, temp_count, information
        Close #3
        
        EditSpecificView.Text = information.speech
        
        If information.character = speaker_1 Then
            EditNarratorDavid.Value = True
        Else
            EditNarratorZira.Value = True
        End If
    Else
        temp_count = temp_count + 1
    End If
    
    DialogueSerialLablel.Caption = temp_count
End Sub


Private Sub EditDeletedialogue_Click()
    'copy only the data that doesnot match with the file location i.e. i <> temp_count
    If EditSpecificView.Text <> "" And AllowDelete = True Then
        DialogueSerialLablel.Caption = temp_count
        
        Open filename For Random As #3 Len = 106
        Open "dup.txt" For Random As #4 Len = 106
            form_count = speech_count(filename, information)
            j = 1
            For i = 1 To form_count Step 1
                Get #3, i, information
                If i <> temp_count Then
                    Put #4, j, information
                    j = j + 1
                End If
            Next i
        Close #4
        Close #3
        
        Kill filename
        Name "dup.txt" As filename
        
        'update textboxes
        'load file content
        EditConversation.Text = ""
        
        Open filename For Random As #4 Len = 106
            form_count = speech_count(filename, information)
            For i = 1 To form_count Step 1
                Get #4, i, information
                EditConversation.Text = EditConversation.Text & i & " " + information.character + " : " + information.speech + vbNewLine
            Next i
        Close #4
        
        If temp_count >= form_count Then temp_count = form_count
        
        DialogueSerialLablel.Caption = temp_count
        
        Open filename For Random As #4 Len = 106
            If temp_count > 0 Then
                Get #4, temp_count, information
            Else
                Get #4, 1, information
            End If
            EditSpecificView.Text = information.speech
        Close #4
            
        If information.character = speaker_1 Then
            EditNarratorDavid.Value = True
        Else
            EditNarratorZira.Value = True
        End If
    ElseIf AllowDelete = False Then
        res = MsgBox("          Please choose the file first.", vbOKOnly + vbInformation, "No file chosen!")
    End If
End Sub


Private Sub EditDeleteFile_Click()
    If AllowDelete = True Then
        res = MsgBox("Are you sure you want to delete this file?", vbYesNo + vbInformation, "File Deletion")
        If res = 6 Then
            AllowDelete = False
            On Error GoTo errorhandler
            Kill filename
        
            filename = "temp.txt"
        
            EditFilenameLabel.Caption = ""
        
            EditFileList.Refresh
            EditConversation.Text = ""
            EditSpecificView.Text = ""
            
            DialogueSerialLablel.Caption = ""
            
            Exit Sub
errorhandler:
            Exit Sub
        End If
    ElseIf AllowDelete = False Then
        res = MsgBox("          Please choose the file first.", vbOKOnly + vbInformation, "No file chosen!")
    End If
End Sub


Private Sub EditExchangeCharacter_Click()
    If EditSpecificView.Text <> "" And AllowDelete = True And EditNarratorZira <> False Or EditNarratorDavid <> False Then
        Open filename For Random As #3 Len = 106
            Get #3, temp_count, temp_info
            
            If EditNarratorDavid = True Then
                temp_info.character = speaker_2
            Else
                temp_info.character = speaker_1
            End If
            
            Put #3, temp_count, temp_info
        Close #3
        
        EditConversation.Text = ""
        
        'update conversation preview textbox
        'load file content
        Open filename For Random As #4 Len = 106
            form_count = speech_count(filename, information)
            
            For i = 1 To form_count Step 1
                Get #4, i, information
                EditConversation.Text = EditConversation.Text & i & " " + information.character + " : " + information.speech + vbNewLine
            Next i
        Close #4
        
        Open filename For Random As #4 Len = 106
            Get #4, temp_count, information
            EditSpecificView.Text = information.speech
        Close #4
            
        If information.character = speaker_1 Then
            EditNarratorDavid.Value = True
        Else
            EditNarratorZira.Value = True
        End If
    ElseIf AllowDelete = False Then
        res = MsgBox("          Please choose the file first.", vbOKOnly + vbInformation, "No file chosen!")
    End If
End Sub


'edit frame -> file list
Private Sub EditFileList_dblClick()
    EditConversation.Text = ""
    filename = EditFileList.filename
    temp_count = 1
    DialogueSerialLablel.Caption = temp_count
    AllowDelete = True
        
    EditFilenameLabel.Caption = filename
    
    'load file content
    Open filename For Random As #3 Len = 106
        form_count = speech_count(filename, information)
        
        For i = 1 To form_count Step 1
            Get #3, i, information
            EditConversation.Text = EditConversation.Text & i & " " + information.character + " : " + information.speech + vbNewLine
        Next i
    Close #3
    
    form_count = speech_count(filename, information)
    
    EditSpecificView.Text = ""
    
    If form_count > 0 Then
        temp_count = 1
        Open filename For Random As #4 Len = 106
            Get #4, 1, information
        Close #4
        
        If information.character = speaker_1 Then
            EditNarratorDavid.Value = True
        Else
            EditNarratorZira.Value = True
        End If
        
        EditSpecificView.Text = information.speech
    End If
End Sub


Private Sub EditForward_Click()
    temp_count = temp_count + 1
    
    If temp_count > 0 And temp_count <= form_count And AllowDelete = True Then
        Open filename For Random As #3 Len = 106
            Get #3, temp_count, information
        Close #3
        
        EditSpecificView.Text = information.speech
        
        If information.character = speaker_1 Then
            EditNarratorDavid.Value = True
        Else
            EditNarratorZira.Value = True
        End If
    Else
        temp_count = temp_count - 1
    End If
    
    DialogueSerialLablel.Caption = temp_count
End Sub


'edit frame -> update
Private Sub EditUpdate_Click()
    If EditSpecificView.Text <> "" And AllowDelete = True Then
        Open filename For Random As #3 Len = 106
            Get #3, temp_count, temp_info
            temp_info.speech = EditSpecificView.Text
            Put #3, temp_count, temp_info
        Close #3
        
        EditConversation.Text = ""
        
        'update conversation preview textbox
        'load file content
        Open filename For Random As #4 Len = 106
            form_count = speech_count(filename, information)
            
            For i = 1 To form_count Step 1
                Get #4, i, information
                EditConversation.Text = EditConversation.Text & i & " " + information.character + " : " + information.speech + vbNewLine
            Next i
        Close #4
        
        Open filename For Random As #5 Len = 106
            Get #5, temp_count, information
        Close #5
        
        'check option
        If information.character = speaker_1 Then
            EditNarratorDavid.Value = True
        Else
            EditNarratorZira.Value = True
        End If
        
        EditSpecificView.Text = information.speech
    ElseIf AllowDelete = False Then
        res = MsgBox("          Please choose the file first.", vbOKOnly + vbInformation, "No file chosen!")
    End If
End Sub


Private Sub Viewdialogue_Click()
    If ViewConversation.Text <> "" Then
        Open filename For Random As #3 Len = 106
            i = 1
            While EOF(3) = False
                Get #3, i, information
                If information.character = speaker_1 Then
                    Call Conversation(information.speech, speaker_1)
                Else
                    Call Conversation(information.speech, speaker_2)
                End If
                i = i + 1
            Wend
        Close #3
    Else
        res = MsgBox("          Please select the file first.", vbOKOnly + vbInformation, "No file chosen!")
    End If
End Sub


Private Sub StartConversation_Click()
    If AddPreview.Text <> "" Then
        i = 1
        Open "temp.txt" For Random As #5 Len = 106
            While EOF(5) = False
                Get #5, i, information
                If information.character = speaker_1 Then
                    Call Conversation(information.speech, speaker_1)
                Else
                    Call Conversation(information.speech, speaker_2)
                End If
                i = i + 1
            Wend
        Close #5
    Else
        res = MsgBox("          No conversation has been created.", vbOKOnly + vbInformation, "Empty textbox")
    End If
End Sub


'save command button
Private Sub Command9_Click()
    'ask for filename
    filename = InputBox("Enter filename for saving the conversation : ")
    If filename <> "" Then
        filename = filename & ".converz"
        
        'check for filename redundency
        Status = filename_redundency(filename)
        
        If Status = True Then
            response = MsgBox("     Sorry this filename is already taken.", vbOKOnly + vbInformation, "Filename Redundency")
        Else
            Name "temp.txt" As filename
            Open "temp.txt" For Random As #3
            Close #3
            'reset preview textbox
            AddPreview.Text = ""
        End If
    End If
End Sub


'add navigation
Private Sub AddImage_Click()
    AddImage.Picture = LoadPicture("Image/Add active.jpg")
    ViewImage.Picture = LoadPicture("Image/View passive.jpg")
    EditImage.Picture = LoadPicture("Image/Edit passive.jpg")
    
    AllowDelete = False
    
    'reset textbox values
    ViewConversation.Text = ""
    EditConversation.Text = ""
    EditSpecificView.Text = ""
    
    'reset textbox values of edit frame
    If EditFrame.Visible = True Then
        AllowDelete = False
        filename = "temp.txt"
        EditConversation.Text = ""
        EditSpecificView.Text = ""
        EditFilenameLabel.Caption = ""
        DialogueSerialLablel.Caption = ""
    End If
    
    'reset textbox values of view frame
    If ViewFrame.Visible = True Then
        ViewConversation.Text = ""
        ViewFilenameLabel.Caption = ""
    End If
    
    'frame visibility
    AddFrame.Visible = True
    ViewFrame.Visible = False
    EditFrame.Visible = False
End Sub


Private Sub ViewDialog_Click()
    If ViewConversation.Text <> "" Then
        i = 1
        Open filename For Random As #5 Len = 106
            While EOF(5) = False
                Get #5, i, information
                If information.character = speaker_1 Then
                    Call Conversation(information.speech, speaker_1)
                Else
                    Call Conversation(information.speech, speaker_2)
                End If
                i = i + 1
            Wend
        Close #5
    Else
        res = MsgBox("          Please choose the file first.", vbOKOnly + vbInformation, "Empty textbox")
    End If
End Sub

'view frame -> file list
Private Sub ViewFileList_DblClick()
    filename = ViewFileList.filename
    ViewFilenameLabel.Caption = filename
    ViewConversation.Text = ""
    
    Open filename For Random As #3 Len = 106
        form_count = speech_count(filename, information)
        For i = 1 To form_count Step 1
            Get #3, i, information
            ViewConversation.Text = ViewConversation.Text + information.character + " : " + information.speech + vbNewLine
        Next i
    Close #3
End Sub


'view navigation
Private Sub ViewImage_Click()
    AddImage.Picture = LoadPicture("Image/Add passive.jpg")
    ViewImage.Picture = LoadPicture("Image/View active.jpg")
    EditImage.Picture = LoadPicture("Image/Edit passive.jpg")
    
    'refresh file list
    ViewFileList.Refresh
    
    AllowDelete = False
    
    'reset textbox values of edit frame
    If EditFrame.Visible = True Then
        AllowDelete = False
        EditConversation.Text = ""
        EditSpecificView.Text = ""
        EditFilenameLabel.Caption = ""
        filename = "temp.txt"
        DialogueSerialLablel.Caption = ""
    End If
    
    'reset textbox values of add frame
    If AddFrame.Visible = True Then
        AddText1.Text = ""
        AddPreview.Text = ""
        On Error GoTo errorhandler
        Kill "temp.txt"
    End If
    
errorhandler:
    'frame visibility
    AddFrame.Visible = False
    ViewFrame.Visible = True
    EditFrame.Visible = False
End Sub


'edit navigation
Private Sub EditImage_Click()
    AddImage.Picture = LoadPicture("Image/Add passive.jpg")
    ViewImage.Picture = LoadPicture("Image/View passive.jpg")
    EditImage.Picture = LoadPicture("Image/Edit active.jpg")
    
    'refresh file list of edit frame
    EditFileList.Refresh
    AllowDelete = False
    
    'reset textbox values of add frame
    If AddFrame.Visible = True Then
        AddText1.Text = ""
        AddPreview.Text = ""
       
        On Error GoTo errorhandler
            Kill "temp.txt"
    End If
errorhandler:
    
    'reset textbox values of edit frame
    If ViewFrame.Visible = True Then
        ViewFilenameLabel.Caption = ""
        ViewConversation.Text = ""
    End If
    
    'frame visibility
    AddFrame.Visible = False
    ViewFrame.Visible = False
    EditFrame.Visible = True
End Sub


Private Sub Form_Load()
    AddImage.Picture = LoadPicture("Image/Add active.jpg")
    ViewImage.Picture = LoadPicture("Image/View passive.jpg")
    EditImage.Picture = LoadPicture("Image/Edit passive.jpg")
    
    'frame visibility
    AddFrame.Visible = True
    ViewFrame.Visible = False
    EditFrame.Visible = False
    
    AddText1.Text = ""
    AddPreview.Text = ""
    
    Open "temp.txt" For Output As #1
    Close #1
    
    character = "David"
    speaker_1 = "David"
    speaker_2 = "Zira"
    active_speaker = "David"
    
    AllowDelete = False
    OptionDavid.Value = True
    OptionZira.Value = False
    
    temp_count = 0
    
    'view frame default value setting
    ViewConversation.Text = ""
    EditSpecificView.Text = ""
    
    EditFilenameLabel.Caption = ""
    ViewFilenameLabel.Caption = ""
    DialogueSerialLablel.Caption = ""
End Sub
