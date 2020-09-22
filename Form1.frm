VERSION 5.00
Begin VB.Form frmBassMK2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Bass Maker 2«"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7185
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pboxWavePreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   675
      Left            =   60
      ScaleHeight     =   675
      ScaleWidth      =   1815
      TabIndex        =   11
      Top             =   60
      Width           =   1815
      Begin VB.Timer timerWavePreview 
         Interval        =   500
         Left            =   1140
         Top             =   120
      End
   End
   Begin VB.PictureBox pboxLargeGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   1440
      MouseIcon       =   "Form1.frx":1272
      MousePointer    =   99  'Custom
      ScaleHeight     =   2415
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   1560
      Width           =   5295
      Begin VB.Line lineH2CrossFollowingMouse 
         BorderColor     =   &H00008000&
         X1              =   9100
         X2              =   8640
         Y1              =   9340
         Y2              =   9340
      End
      Begin VB.Line lineV2CrossFollowingMouse 
         BorderColor     =   &H00008000&
         X1              =   9100
         X2              =   9100
         Y1              =   2340
         Y2              =   3480
      End
      Begin VB.Line lineV1CrossFollowingMouse 
         BorderColor     =   &H00008000&
         X1              =   7260
         X2              =   7260
         Y1              =   9940
         Y2              =   9080
      End
      Begin VB.Line lineH1CrossFollowingMouse 
         BorderColor     =   &H00008000&
         X1              =   5580
         X2              =   9120
         Y1              =   3600
         Y2              =   3600
      End
   End
   Begin VB.HScrollBar hscrolGraphWaveNO 
      Height          =   195
      LargeChange     =   1000
      Left            =   1440
      Max             =   9700
      SmallChange     =   100
      TabIndex        =   6
      Top             =   4260
      Width           =   5295
   End
   Begin VB.PictureBox pboxGraphWaveNOScale 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   5295
      TabIndex        =   3
      ToolTipText     =   "This is the Wave Number scale, move the scroll bar below to change this scale"
      Top             =   3960
      Width           =   5295
   End
   Begin VB.PictureBox pboxMiniGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1440
      ScaleHeight     =   735
      ScaleWidth      =   5295
      TabIndex        =   2
      ToolTipText     =   "This is a Mini Graph of your whole project"
      Top             =   840
      Width           =   5295
   End
   Begin VB.PictureBox pboxGraphScaleHertz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   180
      ScaleHeight     =   2415
      ScaleWidth      =   1275
      TabIndex        =   1
      ToolTipText     =   "This is the Hertz Scale, double click to change this scale"
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label label2BASSMk2 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   4395
      TabIndex        =   10
      Top             =   255
      Width           =   315
   End
   Begin VB.Label labelBASSMk2 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4380
      TabIndex        =   9
      Top             =   240
      Width           =   315
   End
   Begin VB.Label labelCurrentValue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Value = 0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4860
      TabIndex        =   8
      Top             =   300
      Width           =   2235
   End
   Begin VB.Line lineDotFixerLine 
      BorderColor     =   &H0000FF00&
      Tag             =   "fixes a little spot on the axes where the X-Y meet"
      X1              =   1425
      X2              =   3075
      Y1              =   3975
      Y2              =   2805
   End
   Begin VB.Label labelGraphHertz 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hertz"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label labelWaveNOScaleLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wave Number >"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4020
      Width           =   1275
   End
   Begin VB.Image imageBassMK2 
      Height          =   600
      Left            =   2340
      Picture         =   "Form1.frx":13C4
      Stretch         =   -1  'True
      ToolTipText     =   "Welcome to BASS MK2"
      Top             =   60
      Width           =   2340
   End
   Begin VB.Label labelCurrentCords 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "( 0 , 0 )"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4860
      TabIndex        =   7
      Top             =   60
      Width           =   2235
   End
   Begin VB.Menu menuFileMenu 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu menuNewProj 
         Caption         =   "New Project"
         Index           =   8
         Shortcut        =   ^N
      End
      Begin VB.Menu menuSaveMenu 
         Caption         =   "Save Project"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu menuLoadMenu 
         Caption         =   "Load Project"
         Index           =   3
         Shortcut        =   ^L
      End
      Begin VB.Menu menuMakeWave 
         Caption         =   "Make Wave File"
         Index           =   4
         Shortcut        =   ^W
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
         Index           =   5
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menugraph 
      Caption         =   "Graph"
      Index           =   6
      Begin VB.Menu menucngHertzScale 
         Caption         =   "Change Hertz Scale"
         Index           =   7
         Shortcut        =   ^Y
      End
   End
End
Attribute VB_Name = "frmBASSMK2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
'Bass Maker 2
'----------------------------------------------------------
'
'Updates
'----------------------------------------------------------
'» Easier to use interface
'» Better terminology
'» Users enter data through a graph rather than a textbox
'» Use of Dir and File boxes instead of Windows
'  CommonDialog
'» Added option for silence and Rectification for wave
'  length differences
'» User can now use up to 22050Hz, (the old bass maker
'  displayed Hertz too high)
'» Fixed Code, Algorithms run slightly faster
'» Use of Arrays instead of reading from Sequential Files
'» Better explanations for Code
'----------------------------------------------------------

Dim arrOldFreqVAlue&(1)
'stores the current X,Y mouse cords in terms of hertz and
'wave no.
Dim intCurrentYCordS%
Dim intCurrentXCordS%
'used to store length of the Divisions on the X-axis of
'the graph
Dim longintDivision#
'temp integer (it does come in handy for returning values
'over two different subs)
Dim intTempHolder%
Dim inttempHolder2%

'set up the form
Private Sub Form_Load()
'set the wave No scale
hscrolGraphWaveNO_Change
'set the hertz scale
intMaxHertzScale% = 500
subFixHertzAxes
End Sub

'---------------------------------------------------------
'Exit Bass Maker
'---------------------------------------------------------
'User Selects Exit from the drop down menu
Private Sub menuExit_Click(Index As Integer)
'initiate unload form (jumps to Form_Unload sub)
Unload Me
End Sub
'When unloading the form, show a confirmation message
Private Sub Form_Unload(Cancel As Integer)
'get confirmation from the user
Response$ = MsgBox("Are You Sure?", vbYesNo, "Confirm")
'if user clicks yes then exit bass maker
If Response$ = vbYes Then End
'stop the form unloading
Cancel = 1
End Sub
'----------------------------------------------------------

'----------------------------------------------------------
'Change the Hertz Scale
'----------------------------------------------------------
Private Sub menucngHertzScale_Click(Index As Integer)
pboxGraphScaleHertz_DblClick
End Sub

Private Sub menuLoadMenu_Click(Index As Integer)
'show the save project form
Me.Enabled = False
frmSave.Show
frmSave.cmdbSave.Caption = "Load"
frmSave.txtboxFilename.Enabled = False
frmSave.File1.Pattern = "*.bs2"
End Sub

Public Sub subLoadFromTxt(Filename1 As String)
On Error Resume Next
'close all files
Close
'open the selected file for reading
Open Filename1 For Input As #1
'read all values from the file
Input #1, A$
intMaxHertzScale% = Val(A$)
For I = 0 To 10000
    Input #1, A$
    arrPointsForFreq&(I) = Val(A$)
Next I
'close all files
Close
'refresh form
subFixHertzAxes
hscrolGraphWaveNO_Change
End Sub


Private Sub menuMakeWave_Click(Index As Integer)
Me.Enabled = False
frmSave.Show
frmSave.cmdbSave.Caption = "Save Wave"
frmSave.txtboxFilename.Enabled = True
frmSave.File1.Pattern = "*.*"
End Sub

Private Sub menuNewProj_Click(Index As Integer)
'set all points to zero
For I = 0 To 10000
    arrPointsForFreq&(I) = 0
Next I
'redraw the form
Form_Load
End Sub


Private Sub menuSaveMenu_Click(Index As Integer)
'show the save project form
Me.Enabled = False
frmSave.Show
frmSave.cmdbSave.Caption = "Save"
frmSave.txtboxFilename.Enabled = True
frmSave.File1.Pattern = "*.bs2"
End Sub

Public Sub subSaveToTxt(Filename1 As String)
On Error Resume Next
'close all files
Close
'open the selected file for writing to
Open Filename1 + ".bs2" For Output As #1
'save all data to a file, the ';' stops vb from putting
'an enter for each print statement
Print #1, Trim$(Str$(intMaxHertzScale%)) + ",";
For I = 0 To 10000
    Print #1, Trim$(Str$(arrPointsForFreq&(I))) + ",";
Next I
'close all files
Close
'notify the user
Call MsgBox("Project Saved to " + Filename1 + ".bs2", vbInformation, "Save Complete")
End Sub

Private Sub pboxGraphScaleHertz_DblClick()
120 'show a inputbox to get a value for the new scale
TempString$ = InputBox("What is the new Hertz scale you wish set? (100-minimum, 22050-maximum)", "New Hertz Scale", "500Hz")
'if the value is out of bounds then ask for it again
If Val(TempString$) = 0 Then Exit Sub
If Val(TempString$) < 100 Then GoTo 120
If Val(TempString$) > 22050 Then GoTo 120
'set new value
intMaxHertzScale% = Val(TempString$)
subFixHertzAxes
'redraw graph
subDrawGraph
End Sub
'draw the new hertz scale
Private Sub subFixHertzAxes()
'clear the pciture box
pboxGraphScaleHertz.Cls
For I = 0 To 1 Step 0.1
    'set the point where the label is going to be placed
    '(you used to use locate in Qbasic and Gwbasic instead
    'of pset ,but the locate command has since gone in VB)
    pboxGraphScaleHertz.ForeColor = RGB(0, 0, 0)
    pboxGraphScaleHertz.PSet (15, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I)
    'set the text for the hertz axis
    pboxGraphScaleHertz.ForeColor = RGB(0, 255, 0)
    pboxGraphScaleHertz.Print (I * intMaxHertzScale%)
    'draw some lines to make it easier to see where you are
    'on the graph
    pboxGraphScaleHertz.Line (pboxGraphScaleHertz.Width - 800, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I)-(pboxGraphScaleHertz.Width, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I)
    pboxGraphScaleHertz.Line (pboxGraphScaleHertz.Width - 200, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I - pboxGraphScaleHertz.Height * 0.05)-(pboxGraphScaleHertz.Width, pboxGraphScaleHertz.Height - pboxGraphScaleHertz.Height * I - pboxGraphScaleHertz.Height * 0.05)
Next I
'draw the vertical line to separate the scale from the
'graph
pboxGraphScaleHertz.Line (pboxGraphScaleHertz.Width - 30, 0)-(pboxGraphScaleHertz.Width - 30, pboxGraphScaleHertz.Height)
End Sub
'----------------------------------------------------------

'change wave number scale
Private Sub hscrolGraphWaveNO_Change()
'draw the new wave no axes
pboxGraphWaveNOScale.Cls
For I = 0 To 3
    'set the point for the text to be placed in the scale
    pboxGraphWaveNOScale.ForeColor = RGB(0, 0, 0)
    pboxGraphWaveNOScale.PSet (pboxGraphWaveNOScale.Width / 3 * I, 100)
    'print the text in the picture box
    pboxGraphWaveNOScale.ForeColor = RGB(0, 255, 0)
    pboxGraphWaveNOScale.Print hscrolGraphWaveNO.Value + 100 * I
    'draw lines to divide up the graph (makes graph easier
    'to read
    pboxGraphWaveNOScale.Line (pboxGraphWaveNOScale.Width / 3 * I, 200)-(pboxGraphWaveNOScale.Width / 3 * I, 0)
    pboxGraphWaveNOScale.Line (pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 6), 100)-(pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 6), 0)
    pboxGraphWaveNOScale.Line (pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 12), 60)-(pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width / 12), 0)
    pboxGraphWaveNOScale.Line (pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width * 18 / 72), 60)-(pboxGraphWaveNOScale.Width / 3 * I + (pboxGraphWaveNOScale.Width * 18 / 72), 0)
Next I

'draw the horizontal line to separate the scale from the
'graph
pboxGraphWaveNOScale.Line (0, 15)-(pboxGraphWaveNOScale.Width, 15)
subDrawGraph
End Sub

Private Sub pboxLargeGraph_Click()
'if the point is at zero then set it to -1 (for use in the
'loop, zero is disguarded)
If intCurrentYCordS% = 0 Then intCurrentYCordS% = -1
'set the values of the frequency at this point
arrPointsForFreq&(intCurrentXCordS%) = intCurrentYCordS%
'draw the graph
subDrawGraph
End Sub

'draw the graph
Private Sub subDrawGraph()
On Error Resume Next
'set the values to non-zero (to state that is is the
'first time it is used
arrOldFreqVAlue&(0) = -1
arrOldFreqVAlue&(1) = -1
'set up the graph picture box's colour and refresh it
pboxLargeGraph.Cls
pboxLargeGraph.ForeColor = RGB(0, 255, 0)
'store the length of the divisions on the x-axis (for some
'reason they are not 15? but anyways this fixes that
'problem)
longintDivision# = pboxLargeGraph.Width / 300
'set temp integer to -5 (forces it to check for previous
'points)
intTempHolder% = -5
'This loop does most of the graphing
For I = 0 To 300
    'if value is zero, then skip drawing it (take this part
    'out and see what happens, makes it really hard to draw
    'a graph)
    If arrPointsForFreq&(I + hscrolGraphWaveNO.Value) = 0 Then GoTo 10
    'if the value is -1 then set it to zero, so that if the
    'user wants a zero value, then it will plot a zero
    'value
    If arrPointsForFreq&(I + hscrolGraphWaveNO.Value) = -1 Then arrPointsForFreq&(I + hscrolGraphWaveNO.Value) = 0
        'if temp integer is non-zero then do graphing as
        'per normal, otherwise find a previous value to
        'graph (otherwise the first line will be between
        'the first point and (0,0), try removing this if
        'statement and 'un-remark' the pbox...line
        'statement that is remarked below
        If intTempHolder% <> -5 Then
            'draw a line between the current and previous
            'point
            pboxLargeGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxLargeGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
        Else
            'find previous values (i.e. before
            'hscoll1.value)
            subFindPreviousValue
                If intTempHolder% <> -5 Then
                'if a point does exist before the current
                'point then draw a line between it and
                'the current point
                    pboxLargeGraph.Line ((intTempHolder% - hscrolGraphWaveNO.Value) * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(intTempHolder%) / intMaxHertzScale%) * pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
                Else
                    'set the fore colour to red
                    pboxLargeGraph.ForeColor = RGB(255, 0, 0)
                    'draw a line between the current point
                    'and (0,0)
                    pboxLargeGraph.Line (0, pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
                    'set the temp integer to non-zero (so
                    'that the sub doesn't keep checking
                    'for previous values, probably not
                    'necessary but will save a fraction
                    'of the processing time)
                    intTempHolder% = 1
                    'set the forecolour back to green
                    pboxLargeGraph.ForeColor = RGB(0, 255, 0)
                End If
        End If
'    pboxLargeGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxLargeGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
    'set the current values as the old values
    arrOldFreqVAlue&(0) = I
    arrOldFreqVAlue&(1) = arrPointsForFreq&(I + hscrolGraphWaveNO)
    'if the point is zero, set it back to the original -1
    'value
    If arrPointsForFreq&(I + hscrolGraphWaveNO) = 0 Then arrPointsForFreq&(I + hscrolGraphWaveNO) = -1
'the line number 10 (I hope everyone still understands line
'numbers, if not, it does the same as puting a 'goto start'
'then ':start' in a batch file (.bat) (or the other way
'round :-) ))
10
Next I
'find the next value that is not on the graph's current
'scale
If intTempHolder% = -5 Then
    subFindPreviousValue
    inttempHolder2% = intTempHolder%
    subFindNextValue
    If intTempHolder% <> -5 And inttempHolder2% <> -5 Then
        'if there is a point off the scale then draw a line to
        'it
        pboxLargeGraph.Line ((inttempHolder2% - hscrolGraphWaveNO.Value) * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(inttempHolder2%) / intMaxHertzScale%) * pboxLargeGraph.Height)-((intTempHolder% - hscrolGraphWaveNO.Value) * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(intTempHolder%) / intMaxHertzScale%) * pboxLargeGraph.Height)
    Else
        'if there isn't a line off the scale then draw a red
        'line to (pbox.width,0)
        pboxLargeGraph.ForeColor = RGB(255, 0, 0)
        pboxLargeGraph.Line (inttempHolder2% * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(inttempHolder2%) / intMaxHertzScale%) * pboxLargeGraph.Height)-(intTempHolder% * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(intTempHolder%) / intMaxHertzScale%) * pboxLargeGraph.Height)
    End If
Else
    subFindNextValue
    If intTempHolder% <> -5 Then
        'if there is a point off the scale then draw a line to
        'it
        pboxLargeGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxLargeGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxLargeGraph.Height)-(intTempHolder% * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(intTempHolder%) / intMaxHertzScale%) * pboxLargeGraph.Height)
    Else
        'if there isn't a line off the scale then draw a red
        'line to (pbox.width,0)
        pboxLargeGraph.ForeColor = RGB(255, 0, 0)
        pboxLargeGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxLargeGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxLargeGraph.Height)-(I * longintDivision#, pboxLargeGraph.Height - (arrPointsForFreq&(I + hscrolGraphWaveNO) / intMaxHertzScale%) * pboxLargeGraph.Height)
    End If
End If
'draw the thumb view of the WHOLE graph
subDrawThumbView
End Sub

Private Sub subFindPreviousValue()
'set the temp integer to zero
intTempHolder% = -5
'step backwards from the start of the current scale to zero
'(note:- I don't fully understand why you have to switch
'both the values around (the ones separated by the 'TO')
'and put step -1, no doubt its microsoft crazy logic)
For J = hscrolGraphWaveNO.Value To 0 Step -1
    'if a point has been found then store it in the temp
    'integer and exit the loop
    If arrPointsForFreq&(J) <> 0 Then
        intTempHolder% = J
        Exit For
    End If
Next J
End Sub

Private Sub subFindNextValue()
'set the temp integer to zero
intTempHolder% = -5
'step from the end of the scale to the very last value
'in the array
For J = 300 + hscrolGraphWaveNO.Value To 10000
    If arrPointsForFreq&(J) <> 0 Then
    'if a point has been found then store it in the temp
    'integer and exit the loop
        intTempHolder% = J
        Exit For
    End If
Next J
End Sub

Private Sub subDrawThumbView()
'set the values to non-zero (to state that is is the
'first time it is used
arrOldFreqVAlue&(0) = -1
arrOldFreqVAlue&(1) = -1
'set up the pbox
pboxMiniGraph.Cls
pboxMiniGraph.ForeColor = RGB(0, 155, 0)
'store the division lengths in vbpixels
longintDivision# = pboxMiniGraph.Width / 10000
'this loop draw the thumb view graph
For I = 0 To 10000
    'if the value of the current point is zero then skip it
    If arrPointsForFreq&(I) = 0 Then GoTo 10
    'if the current value is -1 then set it to zero
    If arrPointsForFreq&(I) = -1 Then arrPointsForFreq&(I) = 0
    If arrOldFreqVAlue&(0) <> -1 Then
        'if it is not the first point then draw a line
        'between the current point and the previous point
        pboxMiniGraph.Line (I * longintDivision#, pboxMiniGraph.Height - (arrPointsForFreq&(I) / intMaxHertzScale%) * pboxMiniGraph.Height)-(arrOldFreqVAlue&(0) * longintDivision#, pboxMiniGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxMiniGraph.Height)
    Else
        'if it is the first point then draw a red line
        'betweent the current point and (0,0)
        pboxMiniGraph.ForeColor = RGB(155, 0, 0)
        pboxMiniGraph.Line (I * longintDivision#, pboxMiniGraph.Height - (arrPointsForFreq&(I) / intMaxHertzScale%) * pboxMiniGraph.Height)-(arrOldFreqVAlue&(0) * longintDivision#, pboxMiniGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxMiniGraph.Height)
        pboxMiniGraph.ForeColor = RGB(0, 155, 0)
    End If
    'set the current point to the old points
    arrOldFreqVAlue&(0) = I
    arrOldFreqVAlue&(1) = arrPointsForFreq&(I)
    'if the current point is zero then set it back to -1
    If arrPointsForFreq&(I) = 0 Then arrPointsForFreq&(I) = -1
10
Next I
'draw a red line between the current point and
'(pbox.width,0)
pboxMiniGraph.ForeColor = RGB(155, 0, 0)
pboxMiniGraph.Line (arrOldFreqVAlue&(0) * longintDivision#, pboxMiniGraph.Height - (arrOldFreqVAlue&(1) / intMaxHertzScale%) * pboxMiniGraph.Height)-(10000 * longintDivision#, pboxMiniGraph.Height)
End Sub

'When mouse moves in the graph picture box
Private Sub pboxLargeGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'----------------------------------------------
'make the four lines move with the mouse
'----------------------------------------------
lineH1CrossFollowingMouse.X1 = -160
lineH1CrossFollowingMouse.X2 = X - 120
lineH1CrossFollowingMouse.Y1 = Y
lineH1CrossFollowingMouse.Y2 = Y
lineV1CrossFollowingMouse.X1 = X
lineV1CrossFollowingMouse.X2 = X
lineV1CrossFollowingMouse.Y1 = -160
lineV1CrossFollowingMouse.Y2 = Y - 120
lineH2CrossFollowingMouse.X1 = X + 120
lineH2CrossFollowingMouse.X2 = pboxLargeGraph.Width
lineH2CrossFollowingMouse.Y1 = Y
lineH2CrossFollowingMouse.Y2 = Y
lineV2CrossFollowingMouse.X1 = X
lineV2CrossFollowingMouse.X2 = X
lineV2CrossFollowingMouse.Y1 = Y + 120
lineV2CrossFollowingMouse.Y2 = pboxLargeGraph.Height
'----------------------------------------------
'display the coordinates of the mouse cursor
intCurrentXCordS% = hscrolGraphWaveNO.Value + Int((X / (pboxLargeGraph.Width - 15)) * (300))
intCurrentYCordS% = Int(intMaxHertzScale% - (intMaxHertzScale% * (Y / (pboxLargeGraph.Height - 15))))
labelCurrentCords.Caption = "(" + Str$(intCurrentXCordS%) + " ," + Str$(intCurrentYCordS%) + " )"
'display value of point at this spot
If arrPointsForFreq&(intCurrentXCordS%) = -1 Then labelCurrentValue.Caption = "Value = Silence" Else labelCurrentValue.Caption = "Value =" + Str$(arrPointsForFreq&(intCurrentXCordS%))
End Sub

'display the current frequency as a sine wave
Private Sub timerWavePreview_Timer()
pboxWavePreview.Cls
'draw a sine wave to represent the current frequency
For I = 0 To pboxWavePreview.Width Step 15
    pboxWavePreview.Line (I, (pboxWavePreview.Height / 2) + (pboxWavePreview.Height / 2 - 15) * Sin(intCurrentYCordS% / 10000 * I))-(I - 15, (pboxWavePreview.Height / 2) + (pboxWavePreview.Height / 2 - 15) * Sin(intCurrentYCordS% / 10000 * (I - 15)))
Next I
End Sub
