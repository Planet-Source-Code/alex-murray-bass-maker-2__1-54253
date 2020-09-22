VERSION 5.00
Begin VB.Form frmMakeWave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Make Wave"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtboxFileName 
      Height          =   315
      Left            =   1140
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3840
      Width           =   3555
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Options"
      Height          =   2475
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4695
      Begin VB.TextBox txtboxRectification 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Text            =   "20"
         ToolTipText     =   "This will multiply the (highest frequency) by (this number), the (highest freq)/2 will be multiplied by (this number)/2 and so on"
         Top             =   1725
         Width           =   1095
      End
      Begin VB.CheckBox chkCompensate 
         Caption         =   "Rectification for wave length Differences"
         Height          =   315
         Left            =   180
         TabIndex        =   9
         ToolTipText     =   $"Form3.frx":0000
         Top             =   1440
         Width           =   3435
      End
      Begin VB.CheckBox chkIgnorZero 
         Caption         =   "Ignore all zero Hertz values"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkCalcZero 
         Caption         =   "Calculate zero Hertz values"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtboxEachZero 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Text            =   "1000"
         Top             =   885
         Width           =   1095
      End
      Begin VB.Label labelDescription 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(suggested 10 - 50) "
         Height          =   435
         Left            =   3420
         TabIndex        =   12
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label labelRectification1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount of rectification"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         ToolTipText     =   "This will multiply the (highest frequency) by (this number), the (highest freq)/2 will be multiplied by (this number)/2 and so on"
         Top             =   1740
         Width           =   2535
      End
      Begin VB.Line lineSpacer2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   4680
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line lineSpacer 
         BorderColor     =   &H00FFFFFF&
         X1              =   15
         X2              =   4680
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label labelZeroTime 
         Caption         =   "Time for each zero value"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   900
         Width           =   1995
      End
      Begin VB.Label labelTimeMs 
         BackStyle       =   0  'Transparent
         Caption         =   "µs"
         Height          =   195
         Left            =   3660
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4860
      TabIndex        =   1
      Top             =   2220
      Width           =   1155
   End
   Begin VB.CommandButton cmdCompileWave 
      Caption         =   "Make Wave"
      Height          =   315
      Left            =   4860
      TabIndex        =   0
      Top             =   1860
      Width           =   1155
   End
End
Attribute VB_Name = "frmMakeWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'used to store the position of a point on the sine wave
Dim longintPointOnSineW#
'this store the returned values from a hex conversion
Dim strRetCharFromHexConversion$
'temp strings
Dim strTemp4$
Dim strTemp1$
Dim strTemp2$
Dim strTemp3$
'The exact file size of the wave file
Dim longintTotalFileSize&
'stores the value of pi
Dim Pi#
'a new array which has all the points stored in it
Dim arrNewPointsForFreq&(10000)
'store the start and end of a project
Dim intStart%
Dim intEnd%


Private Sub subCalculateNewPoints()
'set K=0 (i.e. the first point)
K = 0
subFindNextValue (K)

'if it cannot find a first point (i.e. no points exist)
'then exit sub
If strTemp3$ = "-1" Then
    Exit Sub
End If

'set the start of the project (i.e. the first point found)
intStart% = Val(strTemp3$)
'this loop calculates all the points on the lines created
'by the user on the graph
For K = intStart% To 10000
    'is there is no point then goto the next K
    If arrPointsForFreq&(K) = 0 Then GoTo 10
    'find the next non-zero value
    subFindNextValue (K + 1)
    'if there is no next nob-zero value then store the
    'current point and continue (otherwise the very last
    'point will be disguarded)
    If Val(strTemp3$) = -1 Then
        arrNewPointsForFreq&(K) = arrPointsForFreq&(K)
        intEnd% = K
        GoTo 10
    End If
    'if the values are set to -1 (i.e. user plotted zero
    'on the graph (comes up as silence)) then set them
    'to zero (so that zero values are not counted as null
    'values)
    If arrPointsForFreq&(K) = -1 Then arrPointsForFreq&(K) = 0
    If arrPointsForFreq&(Val(strTemp3$)) = -1 Then arrPointsForFreq&(Val(strTemp3$)) = 0
    OldK = K
    'if the two points are consecutive then there are no
    'points between, therefore continue with the loop
    If Val(strTemp3$) - (K + 1) = 0 Then
        arrNewPointsForFreq&(K) = arrPointsForFreq&(K)
    Else
        'if the two points are paralell to the x-axis, then
        'all the points between them are equal to the two
        'points
        If arrPointsForFreq&(K) = arrPointsForFreq&(Val(strTemp3$)) Then
            For G = K To Val(strTemp3$)
                arrNewPointsForFreq&(G) = arrPointsForFreq&(K)
            Next G
            K = Val(strTemp3$) - 1
        Else
            'if the current point is less than the next
            'point, therefore the values in between will
            'start from the current point and progress
            'linearly upward to the next point
            If arrPointsForFreq&(K) < arrPointsForFreq&(Val(strTemp3$)) Then
                For G = K To Val(strTemp3$)
                    arrNewPointsForFreq&(G) = arrPointsForFreq&(K) + (arrPointsForFreq&(Val(strTemp3$)) - arrPointsForFreq&(K)) * ((G - K) / (Val(strTemp3$) - K))
                Next G
                K = Val(strTemp3$) - 1
            Else
                'if the current point is more than the next
                'point, therefore the values in between
                'will start from the current point and
                'progress linearly downward to the next point
                For G = K To Val(strTemp3$)
                    arrNewPointsForFreq&(G) = arrPointsForFreq&(K) - (arrPointsForFreq&(K) - arrPointsForFreq&(Val(strTemp3$))) * ((G - K) / (Val(strTemp3$) - K))
                Next G
                K = Val(strTemp3$) - 1
            End If
        End If
    End If
    'if the points used were set to zero, then set them
    'back to -1
    If arrPointsForFreq&(Val(strTemp3$)) = 0 Then arrPointsForFreq&(Val(strTemp3$)) = -1
    If arrPointsForFreq&(OldK) = 0 Then arrPointsForFreq&(OldK) = -1
    'set the end of the project (it will be equal to the
    'last value)
    intEnd% = K
10
Next K

End Sub

Private Sub subFindNextValue(intStartFrom As Integer)
'set a temp string to '-1'
strTemp3$ = "-1"
'find the next point starting from intStart
For J = intStartFrom To 10000
    'if the value is non-zero, therefore it is the next
    'point
    If arrPointsForFreq&(J) <> 0 Then
        strTemp3$ = Str$(J)
        Exit For
    End If
Next J
'note: if no values are found, strtemp3$ will equal '-1'
End Sub

Private Sub chkCompensate_Click()
'enable/disable the rectification textbox
If chkCompensate.Value = 1 Then
    txtboxRectification.Enabled = True
Else
    txtboxRectification.Enabled = False
End If
End Sub

Private Sub chkIgnorZero_Click()
'enable/disable calculation of zero values
If chkIgnorZero.Value = 1 Then
    chkCalcZero.Value = 0
    txtboxEachZero.Enabled = False
End If
If chkIgnorZero.Value = 0 Then
    chkCalcZero.Value = 1
    txtboxEachZero.Enabled = True
End If
End Sub

Private Sub chkCalcZero_Click()
'enable/disable calculation of zero values
If chkCalcZero.Value = 1 Then
    chkIgnorZero.Value = 0
    txtboxEachZero.Enabled = True
End If
If chkCalcZero.Value = 0 Then
    chkIgnorZero.Value = 1
    txtboxEachZero.Enabled = False
End If
End Sub

Private Sub cmdCompileWave_Click()
On Error Resume Next

Me.Caption = "Make Wave - Making Wave... Please Wait..."
Me.Cls

'close all open files
Close

'Workout where the project starts from in the array, and
'calculate all the points on all the lines on the graph,
'and work out where the project finishes from
subCalculateNewPoints

'the file to save the wave data to
Open txtboxFilename.Text For Output As #1

'the value of pi, ie the curcumference of a circle diameter
'of 1
Pi# = 3.141592654

'write the start of the wave file and set file size to 110
Print #1, "RIFF";: longintTotalFileSize& = 110

'if Rectification is enabled, then set the amount of
'repetitions of each wave, if not, set the amount of reps
'to 1 (i.e. only loops once per wave)
If chkCompensate.Value = 1 Then
    intAmountofReps% = Val(txtboxRectification.Text)
Else
    intAmountofReps% = 1
End If

'calculate the size of the wave file
For I = intStart% To intEnd%
    'check to see if zero calculation is enabled
    If arrNewPointsForFreq&(I) = 0 And chkIgnorZero.Value = 1 Then
        GoTo 204
    Else
        'if it is zero then calculate the size of all zero
        'values as selected by the user (note when L is
        'repeated 22050 times, it is equal to 1 second)
        If arrNewPointsForFreq&(I) = 0 And chkIgnorZero.Value = 0 Then
            For L = 1 To (Val(txtboxEachZero.Text) * 0.02205)
                longintTotalFileSize& = longintTotalFileSize& + 4
            Next L
            GoTo 204
        End If
    End If
    
    'calculate the amount of reps to be done (note: the
    'repetitions are dependent on the frequency of the
    'wave)
    intSteps% = (intMaxHertzScale% / arrNewPointsForFreq&(I))
    If intSteps% = 0 Then intSteps% = intAmountofReps%
    'if the intsteps is somehow zero then set it to 1
    
    'this loop repeats the waves
    For B = 1 To intAmountofReps% Step intSteps%
        'for O= 0 to (pi*2) step
        '(Frequency * (2*Pi)/(amount of bytes for
        '1 second i.e. 22050, hence sample rate
        '22050hz)
        For O# = 0 To (Pi# * 2) Step (arrNewPointsForFreq&(I) * (2 * Pi# / 22050))
            'since the sample rate 22050 uses 2 bytes per
            'channel per point on the wave in the wave file
            ', therefore each point takes up 4 bytes. this
            'is used to calculate the total file size for
            'information to be written to the header
            longintTotalFileSize& = longintTotalFileSize& + 4
        Next O#
    Next B
204
Next I

'calculate the size of the whole wave files (inc. header)
strTemp4$ = Hex(longintTotalFileSize&)
strTemp4$ = Space(8 - Len(strTemp4$)) + strTemp4$
HexConverTT (Mid$(strTemp4$, 7, 2))
place1$ = strRetCharFromHexConversion$
HexConverTT (Mid$(strTemp4$, 5, 2))
place2$ = strRetCharFromHexConversion$
HexConverTT (Mid$(strTemp4$, 3, 2))
place3$ = strRetCharFromHexConversion$
HexConverTT (Mid$(strTemp4$, 1, 2))
Print #1, place1$ + place2$ + place3$ + strRetCharFromHexConversion$;

'write majority of header (i.e. sample rate, format e.t.c.)
Print #1, Chr$(&H57) + Chr$(&H41) + Chr$(&H56) + Chr$(&H45) + Chr$(&H66) + Chr$(&H6D) + Chr$(&H74) + Chr$(&H20) + Chr$(&H10) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + Chr$(&H1) + Chr$(&H0) + Chr$(&H2) + Chr$(&H0) + Chr$(&H22) + Chr$(&H56) + Chr$(&H0) + Chr$(&H0) + Chr$(&H88) + Chr$(&H58) + Chr$(&H1) + Chr$(&H0) + Chr$(&H4) + Chr$(&H0) + Chr$(&H10) + Chr$(&H0) + Chr$(&H64) + Chr$(&H61) + Chr$(&H74) + Chr$(&H61);

'calculate the size of the wave file without its header
strTemp4$ = Hex(longintTotalFileSize& - 110)
strTemp4$ = Space(8 - Len(strTemp4$)) + strTemp4$
HexConverTT (Mid$(strTemp4$, 7, 2))
place1$ = strRetCharFromHexConversion$
HexConverTT (Mid$(strTemp4$, 5, 2))
place2$ = strRetCharFromHexConversion$
HexConverTT (Mid$(strTemp4$, 3, 2))
place3$ = strRetCharFromHexConversion$: HexConverTT (Mid$(strTemp4$, 1, 2))
Print #1, place1$ + place2$ + place3$ + strRetCharFromHexConversion$;

'write the majority of the wave file
For I = intStart% To intEnd%
    'check to see if zero calculation is enabled
    If arrNewPointsForFreq&(I) = 0 And chkIgnorZero.Value = 1 Then
        GoTo 203
    Else
        'if it is zero then calculate the size of all zero
        'values as selected by the user (note when L is
        'repeated 22050 times, it is equal to 1 second)
        If arrNewPointsForFreq&(I) = 0 And chkIgnorZero.Value = 0 Then
            For L = 1 To (Val(txtboxEachZero.Text) * 0.02205) Step 1
                Print #1, Chr$(255) + Chr$(255) + Chr$(255) + Chr$(255);
            Next L
            GoTo 203
        End If
    End If
    'calculate the amount of reps to be done (note: the
    'repetitions are dependent on the frequency of the
    'wave)
    intSteps% = (intMaxHertzScale% / arrNewPointsForFreq&(I))
    If intSteps% = 0 Then intSteps% = intAmountofReps%
    'if the intsteps is somehow zero then set it to 1
    
    'this loop repeats the waves
    For B = 1 To intAmountofReps% Step intSteps%
        ' for O= 0 to (pi*2) step (Frequency * (2*Pi)/
        '(amount of bytes for 1 second i.e. 22050,
        'hence sample rate 22050hz)
        For O# = 0 To (Pi# * 2) Step (arrNewPointsForFreq&(I) * (2 * Pi# / 22050))
            'get the next point on the sine wave with
            'period 1/freq& and amplitude 30000
            longintPointOnSineW# = 30000 * Sin(O#)
            'if the value is less than or equal to zero
            'then add &HFFFF (modulus system)
            If Int(longintPointOnSineW#) <= 0 Then longintPointOnSineW# = 65535 + longintPointOnSineW#
            'convert the value to hex
            strTemp1$ = Hex(Int(longintPointOnSineW#))
            'if value returned was say 'A1' or 'F45' then
            'move them to the right hand side i.e. '  A1'
            'and ' F45' (or '00A1' and '0F45')
            strTemp1$ = Space(4 - Len(strTemp1$)) + strTemp1$
            'divide the hex value in half, ie into two lots
            'of two bytes
            strTemp2$ = Mid$(strTemp1$, 1, 2)
            strTemp3$ = Mid$(strTemp1$, 3, 2)
            'convert the hex value to a character
            HexConverTT (strTemp3$)
            'write the converted value to file
            Print #1, strRetCharFromHexConversion$;
            'set the value to a temp string because it will
            'otherwise be overwritten
            strTemp3$ = strRetCharFromHexConversion$
            'convert the other value to a character
            HexConverTT (strTemp2$)
            'write the converted values to file
            Print #1, strRetCharFromHexConversion$ + strTemp3$ + strRetCharFromHexConversion$;
        Next O#
    Next B
203
Next I

'to be written as a comment in the wave files
message1$ = "Made With Bass Maker 2"

'write the end of the wave file (e.g. comments e.t.c.)
Print #1, Chr$(&H4C) + Chr$(&H49) + Chr$(&H53) + Chr$(&H54) + Chr$(&H42) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + Chr$(&H49) + Chr$(&H4E) + Chr$(&H46) + Chr$(&H4F) + Chr$(&H49) + Chr$(&H53) + Chr$(&H46) + Chr$(&H54) + Chr$(&H35) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + message1$ + Space(52 - Len(message1$)) + Chr$(&H0) + Chr$(&H0);

'close all open files
Close

'notify the user that working is complete and return to the
'main form
Call MsgBox("The wave file has been successfully created", vbInformation, "Done")
Me.Caption = "Make Wave"
Me.Refresh
cmdCancel_Click
End Sub

Private Sub HexConverTT(Hexd$)
'separate the two byte hex value into two values
A1$ = Mid$(Hexd$, 1, 1): B1$ = Mid$(Hexd$, 2, 1)

'if their values are A,B,C,D,E, or F then set their values
'to 10,11,12,13,14,15 respectivly
If Asc(A1$) >= 65 Then A1$ = Str$(Asc(A1$) - 55)
If Asc(B1$) >= 65 Then B1$ = Str$(Asc(B1$) - 55)

'calculate its value in decimal and then convert it to a
'character
strRetCharFromHexConversion$ = Chr$((Val(A1$) * 16) + Val(B1$))
End Sub

'return to main form
Private Sub cmdCancel_Click()
Me.Hide
frmBASSMK2.Enabled = True
frmBASSMK2.Hide
frmBASSMK2.Show
End Sub

'-------------------------------------------------------
'a few error trapping subroutines
'-------------------------------------------------------
Private Sub txtboxEachZero_Change()
If Val(txtboxEachZero) < 50 Then txtboxEachZero = "50"
End Sub

Private Sub txtboxRectification_Change()
If Val(txtboxRectification) < 1 Then txtboxRectification = "1"
If Val(txtboxRectification) > 22050 Then txtboxRectification = "22050"
End Sub
'-------------------------------------------------------
