VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "My Recipe Box          by Ken Foster"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13620
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   13620
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ThreeDText ThreeDText1 
      Height          =   900
      Left            =   4020
      TabIndex        =   17
      Top             =   1695
      Width           =   3975
      _extentx        =   7011
      _extenty        =   1588
      caption         =   "My Recipe Box"
      colors          =   16711935
      colore          =   8454016
      colorf          =   16744703
      font            =   "Form1.frx":0562
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "  Recipe  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5700
      Left            =   3930
      TabIndex        =   15
      Top             =   2520
      Width           =   8325
      Begin RichTextLib.RichTextBox rtfMain 
         Height          =   5355
         Left            =   210
         TabIndex        =   16
         Top             =   255
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   9446
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Form1.frx":0586
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5430
         Left            =   105
         Top             =   210
         Width           =   8145
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   2265
      Left            =   8970
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "Form1.frx":05FF
      Top             =   255
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Control Panel "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   4020
      TabIndex        =   1
      Top             =   75
      Width           =   4875
      Begin VB.CommandButton cmdWebBrows 
         Caption         =   "WebBrows"
         Height          =   285
         Left            =   2745
         TabIndex        =   19
         Top             =   1350
         Width           =   1080
      End
      Begin VB.CommandButton cmdPrintPreview 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Preview"
         Height          =   285
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   885
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Edit"
         Height          =   285
         Left            =   1035
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   885
         Width           =   765
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Cancel"
         Height          =   285
         Left            =   1905
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1350
         Width           =   765
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Print"
         Height          =   285
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   885
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add"
         Height          =   285
         Left            =   1035
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1350
         Width           =   765
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Delete"
         Height          =   285
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1350
         Width           =   765
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   885
         Width           =   765
      End
      Begin VB.TextBox txtRecipeName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   165
         TabIndex        =   2
         Top             =   450
         Width           =   4545
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Number    of Recipes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3705
         TabIndex        =   8
         Top             =   885
         Width           =   1140
      End
      Begin VB.Label lblTotRecords 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3930
         TabIndex        =   7
         Top             =   1335
         Width           =   765
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.ListBox lstRecipe 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7830
      Left            =   75
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   375
      Width           =   3825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Recipes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   18
      Top             =   90
      Width           =   1965
   End
   Begin VB.Image Image2 
      Height          =   1410
      Left            =   12285
      Picture         =   "Form1.frx":1435
      Stretch         =   -1  'True
      Top             =   5055
      Width           =   1275
   End
   Begin VB.Image Image5 
      Height          =   675
      Left            =   8085
      Picture         =   "Form1.frx":247B
      Top             =   1845
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   1080
      Left            =   12285
      Picture         =   "Form1.frx":2EFA
      Top             =   3975
      Width           =   405
   End
   Begin VB.Image Image4 
      Height          =   1080
      Left            =   12630
      Picture         =   "Form1.frx":3AAC
      Stretch         =   -1  'True
      Top             =   3975
      Width           =   930
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   5565
      Left            =   12270
      Top             =   2640
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   1305
      Left            =   12300
      Picture         =   "Form1.frx":4A5B
      Stretch         =   -1  'True
      Top             =   2670
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cooking Measurements"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9075
      TabIndex        =   14
      Top             =   45
      Width           =   2205
   End
   Begin VB.Image Image6 
      Height          =   1710
      Left            =   12270
      Picture         =   "Form1.frx":5589
      Stretch         =   -1  'True
      Top             =   6465
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'****************************************************
'*
'*         Project Name : My Recipe Box
'*        Version Number: 1.4.4
'*           Author Name: Ken Foster
'*                 Date : December 11, 2005
'*        Freeware - Use anyway you want.
'*
'****************************************************
'   Print Preview is not by me. See frm and/or modules of credits
'   Updated  December 19, 2005
'***************** Table of Procedures *************
'   Private Sub Form_Load
'   Private Sub cmdAdd_Click
'   Private Sub cmdCancel_Click
'   Private Sub cmdDelete_Click
'   Private Sub cmdEdit_Click
'   Public  Sub cmdPrint_Click
'   Private Sub cmdPrintPreview_Click
'   Private Sub cmdSave_Click
'   Private Sub cmdWebBrows_Click
'   Private Sub lstRecipe_Click
'   Private Sub List_Load
'   Private Sub List_Save
'   Private Sub List_Remove
'   Private Sub RTB_Save
'   Private Sub RTB_Load
'   Private Sub File_Delete
'   Private Sub File_Exists
'   Private Sub Text1_DblClick
'***************** End of Table ********************
'

Private Sub Form_Load()
Dim strPath As String
Dim strMapName As String
Dim fStg As String
Dim fSLen As Integer

'load recipes into listbox
lstRecipe.Clear
strPath = Dir(App.Path & "\RecipeFolder" & "\*.rtf")

If Not strPath = "" Then                                  'yes, there are files here so
   Do                                                     'go get them
      strMapName = strPath
      fSLen = Len(strMapName) - 4                         'filename length minus extension
      fStg = Mid$(strMapName, 1, fSLen)                   'filename without extension
      lstRecipe.AddItem fStg                              'put filename into listbox

      strPath = Dir$
   Loop Until strPath = ""
Else
   MsgBox "No files found!", vbCritical + vbOKOnly, "File - Error"
End If
  
   lblTotRecords.Caption = lstRecipe.ListCount            'Show how many recipes there are in list
   txtRecipeName.Locked = True                            'disable textbox until needed
End Sub
   
Private Sub cmdAdd_Click()                                'clear and enable all needed controls
   'On Error Resume Next
    
   rtfMain.Text = ""
   rtfMain.Locked = False
   txtRecipeName.Locked = False
   cmdSave.Enabled = True
   txtRecipeName.SetFocus
End Sub

Private Sub cmdCancel_Click()                             'clear and disable controls
   txtRecipeName.Locked = True
   txtRecipeName.Text = ""
   rtfMain.Text = ""
   rtfMain.Locked = True
   cmdSave.Enabled = False
End Sub

Private Sub cmdDelete_Click()
   Dim iResponse As Integer
   
   If rtfMain.Text = "" Then
      MsgBox "Nothing to Delete", vbInformation + vbOKOnly, "Delete Error"
      Exit Sub
   End If
   
   iResponse = MsgBox("Are you sure ?", vbInformation + vbYesNo, "Delete this file.")
   If iResponse = 7 Then Exit Sub  'no was selected
   
   Call File_Delete(lstRecipe)                                'delete the recipe from file
   Call List_Remove(lstRecipe)                                'delete the recipe from listbox
   lblTotRecords.Caption = lstRecipe.ListCount
   rtfMain.Text = ""
   txtRecipeName.Locked = True
End Sub

Private Sub cmdEdit_Click()
   If rtfMain.Text = "" Then
      MsgBox "Nothing to edit", vbInformation + vbOKOnly, "Edit Error"
      Exit Sub
   End If
   
   rtfMain.Locked = False
   txtRecipeName.Text = lstRecipe.Text
   cmdSave.Enabled = True
End Sub

Public Sub cmdPrint_Click()  'this needs to be Public ,don't change or print button does'nt work on preview page
     
    ' Print the contents of the RichTextBox with a one inch margin
      On Error GoTo err1
      
      If rtfMain.Text = "" Then
         MsgBox "Nothing to Print", vbInformation + vbOKOnly, "No Recipe to Print"
         Exit Sub
      End If
      
      PrintRTF rtfMain, 1440, 1440, 1440, 1440                   '1440 Twips = 1 Inch
      Exit Sub
err1:
    Select Case Err.Number
        Case 482
            MsgBox "Make sure that you have a printer installed.  If a " & _
                "printer is installed, go into your printer properties " & _
                "look under the Setup tab, and make sure the ICM checkbox " & _
                "is checked and try printing again.", , "Printer Error"
            Exit Sub
        Case Else
            MsgBox Err.Number & " " & Err.Description
    End Select
End Sub

Private Sub cmdPrintPreview_Click()

    If rtfMain.Text = "" Then
       MsgBox "Nothing to preview", vbInformation + vbOKOnly, "No Recipe to Preview"
       Exit Sub
    End If
    
    PrintPreview rtfMain, 1400, 1400, 1400, 1400, Printer.Orientation
End Sub

Private Sub cmdSave_Click()
   Dim Fname As String
   Dim iResponse As String
   
   If txtRecipeName.Text = "" Then
      MsgBox "Please enter a Name."
      Exit Sub
   End If
   
  ' File_Exists
   Fname = App.Path & "\RecipeFolder\" & txtRecipeName.Text & ".rtf"
   FileExists (Fname)

   If FileExists(Fname) = True Then
      iResponse = MsgBox("File Exists!! Do you want to overwrite file?", vbYesNo, "File Exists")
      If iResponse = vbNo Then Exit Sub
      Call RTB_Save                                             'save updated recipe
   Else
      Call RTB_Save                                             'save recipe
      lblTotRecords.Caption = lstRecipe.ListCount               'update file count
      lstRecipe.AddItem txtRecipeName.Text                      'add to listbox
   End If
   
   'control logic
   txtRecipeName.Text = ""
   rtfMain.Text = ""
   rtfMain.Locked = True
   cmdSave.Enabled = False
   txtRecipeName.Locked = True
End Sub

Private Sub cmdWebBrows_Click()
   frmWebBrows.Show
End Sub

Private Sub lstRecipe_Click()
   rtfMain.Text = ""                                            'clears window before loading next recipe
   Call RTB_Load
End Sub

Private Sub List_Remove(TheList As ListBox)
   On Error Resume Next
   If TheList.ListCount < 0 Then Exit Sub
   TheList.RemoveItem TheList.ListIndex
End Sub

Private Sub RTB_Save()
   Dim fFile As Integer
   
   fFile = FreeFile
   Open App.Path & "\RecipeFolder\" & txtRecipeName & ".rtf" For Output As fFile
   Print #fFile, rtfMain.Text                                          ' String location you want To save
   Close fFile
End Sub

Private Sub RTB_Load()
   
   Dim FileLength As Integer
   Dim var1 As String
   Dim fFile As Integer
   
   fFile = FreeFile
   If lstRecipe.ListIndex = -1 Then Exit Sub                           'No item selected
   
   rtfMain.Text = ""
   Open App.Path & "\RecipeFolder\" & lstRecipe & ".rtf" For Input As #fFile
   FileLength = LOF(fFile)
   var1 = Input(FileLength, #fFile)
   rtfMain.Text = var1
   rtfMain.SelStart = 0                                                'Puts Beginning of code at top
   Close #fFile
End Sub

Private Sub File_Delete(TList As ListBox)
   If TList = "" Then Exit Sub
   Kill App.Path & "\RecipeFolder\" & TList & ".rtf"
End Sub

Private Function FileExists(strPath As String) As Integer

    FileExists = Not (Dir(strPath) = "")

End Function

Private Sub Text1_DblClick()   'copies Cooking Measurements to rtfMain for printing out a copy
 rtfMain.Text = ""
 rtfMain.Text = Text1.Text
 cmdPrint.SetFocus
End Sub
