VERSION 5.00
Begin VB.Form frmSave 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   5700
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdbCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5700
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtboxFilename 
      Appearance      =   0  'Flat
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Text            =   "Filename"
      Top             =   5040
      Width           =   5475
   End
   Begin VB.DriveListBox Drive1 
      ForeColor       =   &H00004000&
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   4080
      Width           =   5475
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00004000&
      Height          =   3930
      Left            =   2880
      TabIndex        =   1
      Top             =   60
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00004000&
      Height          =   3915
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2835
   End
   Begin VB.Label labelFilenameCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename to save to:-"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   4800
      Width           =   3435
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbCancel_Click()
Me.Hide
frmBASSMK2.Enabled = True
frmBASSMK2.Hide
frmBASSMK2.Show
End Sub

Private Sub cmdbSave_Click()

If LCase$(cmdbSave.Caption) = "save" Then
'stops the error where if your in the root dir (i.e. 'C:\')
'then it puts 'C:\filename' instead of 'C:\\filename'
    If Len(File1.Path) = 3 Then
        frmBASSMK2.subSaveToTxt (File1.Path + txtboxFileName.Text)
    Else
        frmBASSMK2.subSaveToTxt (File1.Path + "\" + txtboxFileName.Text)
    End If
    'show the main form
    Me.Hide
    frmBASSMK2.Enabled = True
    frmBASSMK2.Hide
    frmBASSMK2.Show
Else
    If LCase$(cmdbSave.Caption) = "save wave" Then
        If Len(File1.Path) = 3 Then
            frmMakeWave.txtboxFileName.Text = (File1.Path + txtboxFileName.Text) + ".wav"
        Else
            frmMakeWave.txtboxFileName.Text = (File1.Path + "\" + txtboxFileName.Text) + ".wav"
        End If
        'show the save wave form
        Me.Hide
        frmBASSMK2.Hide
        frmMakeWave.Hide
        frmMakeWave.Show
    Else
        If Len(File1.Path) = 3 Then
            frmBASSMK2.subLoadFromTxt (File1.Path + File1.FileName)
        Else
            frmBASSMK2.subLoadFromTxt (File1.Path + "\" + File1.FileName)
        End If
        'show the main form
        Me.Hide
        frmBASSMK2.Enabled = True
        frmBASSMK2.Hide
        frmBASSMK2.Show
    End If
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
'file there is a file selected then display it in the text
'box
If File1.FileName <> "" Then txtboxFileName.Text = File1.FileName
End Sub

Private Sub File1_DblClick()
cmdbSave_Click
End Sub
