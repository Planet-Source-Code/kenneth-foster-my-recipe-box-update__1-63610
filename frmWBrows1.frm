VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmWebBrows 
   Caption         =   "Get Recipes"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   330
      Left            =   6450
      TabIndex        =   9
      Top             =   30
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Go"
      Height          =   330
      Left            =   8235
      TabIndex        =   8
      Top             =   375
      Width           =   465
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      TabIndex        =   7
      Top             =   375
      Width           =   8025
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Home"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4290
      TabIndex        =   6
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   330
      Left            =   3240
      TabIndex        =   5
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ReFresh"
      Height          =   330
      Left            =   2205
      TabIndex        =   4
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Forward"
      Height          =   330
      Left            =   1170
      TabIndex        =   3
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      Height          =   330
      Left            =   135
      TabIndex        =   2
      Top             =   30
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   150
      TabIndex        =   1
      Text            =   "      Make a  Selection"
      Top             =   750
      Width           =   8535
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5925
      Left            =   135
      TabIndex        =   0
      Top             =   1140
      Width           =   9225
      ExtentX         =   16272
      ExtentY         =   10451
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmWebBrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()

    Combo1.AddItem "http://www.hungrymonster.com/recipe/recipe-search.cfm"
    Combo1.AddItem "http://www.foodtv.com"
    Combo1.AddItem "http://www.cdkitchen.com/search/allsearch.shtml"
    Combo1.AddItem "http://www.copykat.com/asp/recipes.asp"
End Sub

Private Sub Combo1_Click()
    WebBrowser1.Navigate (Combo1.Text)
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    WebBrowser1.GoBack
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    WebBrowser1.GoForward
End Sub

Private Sub Command3_Click()
    WebBrowser1.Refresh
End Sub

Private Sub Command4_Click()
    WebBrowser1.Stop
End Sub

Private Sub Command5_Click()
  If Combo1.Text = "" Then Exit Sub
  WebBrowser1.Navigate (Combo1.Text)
End Sub

Private Sub Command6_Click()
    If Text1.Text = "" Then Exit Sub
    WebBrowser1.Navigate (Text1.Text)
End Sub

Private Sub Form_Resize()
   WebBrowser1.Width = frmWebBrows.Width - 500
   WebBrowser1.Height = frmWebBrows.Height - 1800
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub
