VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Little Windows SPY"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4755
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtHwnd 
      Height          =   330
      Left            =   1035
      TabIndex        =   5
      Top             =   960
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4035
      Top             =   660
   End
   Begin VB.TextBox wndClassname 
      Height          =   330
      Left            =   1035
      TabIndex        =   3
      Top             =   555
      Width           =   2415
   End
   Begin VB.TextBox txtWinCaption 
      Height          =   330
      Left            =   1035
      TabIndex        =   1
      Top             =   150
      Width           =   4905
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "HWND"
      Height          =   195
      Left            =   30
      TabIndex        =   4
      Top             =   1035
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Class Name"
      Height          =   195
      Left            =   30
      TabIndex        =   2
      Top             =   615
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Window Text"
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   210
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Sub Command1_Click()
        Unload Form1: End
End Sub

Private Sub Timer1_Timer()
Dim Wnd, Length As Long
Dim Wnd_Caption As String
Dim M_Mouse As POINTAPI
Dim Wnd_Class As Long
Dim Wnd_ClassCaption As String


    GetCursorPos M_Mouse ' Get the mouse x and y positions
    
    Wnd = WindowFromPoint(M_Mouse.x, M_Mouse.y) ' Get the windows handle
    Wnd_Caption = Space(256) ' Sets wnd_caption to a space of 256
    Wnd_ClassCaption = Space(256) ' ' Sets wnd_caption to a space of 256
    Length = GetWindowText(Wnd, Wnd_Caption, Len(Wnd_Caption)) ' Get the caption text
    Wnd_Caption = Left(Wnd_Caption, Length) ' trim of any rubbish
    
    
    ' Get the class name
    Wnd_Class = GetClassName(Wnd, Wnd_ClassCaption, Len(Wnd_ClassCaption))
    Wnd_ClassCaption = Left(Wnd_ClassCaption, Wnd_Class)
    
    txtWinCaption.Text = Wnd_Caption ' Display the windows text
    wndClassname.Text = Wnd_ClassCaption ' Displays the windows class name
    txtHwnd.Text = Wnd
    
    
End Sub
