VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Easy Icon Extraction"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstIcons 
      Height          =   2790
      Left            =   6120
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox picIcon 
      Height          =   840
      Left            =   3105
      ScaleHeight     =   780
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.FileListBox FilDLLandEXE 
      Height          =   2820
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.DirListBox Dir 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2370
   End
   Begin VB.DriveListBox drv 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   225
      Width           =   2355
   End
   Begin VB.Label Label2 
      Caption         =   "Filename = "
      Height          =   525
      Left            =   315
      TabIndex        =   6
      Top             =   4710
      Width           =   5850
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Icons in File = "
      Height          =   300
      Left            =   2775
      TabIndex        =   5
      Top             =   225
      Width           =   2010
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ExtractIcon Lib "shell32.dll" _
    Alias "ExtractIconA" (ByVal hInst As Long, _
    ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long) As Long
    
Private Declare Function DestroyIcon Lib "user32" _
    (ByVal hIcon As Long) As Long
    
Private Declare Function DrawIconEx Lib "user32" _
    (ByVal hDC As Long, _
    ByVal xLeft As Long, ByVal yTop As Long, _
    ByVal hIcon As Long, _
    ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, _
    ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long
    
Private Const DI_NORMAL = &H3

Private MyPath As String

Private Sub Form_Load()
    'default drive
    drv.Drive = "C:"
    'only show files with these extentions
    FilDLLandEXE.Pattern = "*.dll;*.exe"
End Sub

Private Sub drv_Change()
    Dir.Path = drv.Drive
End Sub

Private Sub dir_Change()
    FilDLLandEXE.Path = Dir.Path
End Sub

Private Sub FilDLLandEXE_Click()
    Dim i As Long
    Dim Count As Long
    
    MyPath = Dir.Path & "\" & FilDLLandEXE.FileName
    
    Label2.Caption = "Filename = " & MyPath

    'clear LstIcons
    LstIcons.Clear
    
    'Get the number of icons in the file
    Count = ExtractIcon(0, MyPath, -1)
    
    If Count < 1 Then
        Exit Sub
    End If
    'List them for availability
    For i = 0 To Count - 1
        LstIcons.AddItem i
    Next
    
    'sets label1's caption
    Label1.Caption = "Number of Icons in File = " & LstIcons.ListCount
End Sub

Private Sub LstIcons_Click()
    Dim Ico As Long
    
    'if there are no icons don't go on
    If LstIcons.ListCount < 1 Then
        Exit Sub
    End If

    'Get the icon from the index currently selected in
    'lstIcons
    Ico = ExtractIcon(0, MyPath, LstIcons.ListIndex)
    
    'Clear picIcon and put the new icon in it
    picIcon.Cls
    DrawIconEx picIcon.hDC, 0, 0, Ico, _
        0, 0, 0, 0, DI_NORMAL
        
    'Destroy the old icon to free resources
    DestroyIcon Ico
End Sub

