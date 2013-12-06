VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCODStats 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4500
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   8655
   ForeColor       =   &H00000000&
   Icon            =   "frmCODStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   27
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtShotsMissed 
      Height          =   285
      Left            =   4200
      TabIndex        =   11
      Text            =   "# Shots Missed"
      ToolTipText     =   "# Shots Missed"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtShotsHit 
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Text            =   "# Shots Hit"
      ToolTipText     =   "# Shots Hit"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   5880
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdiFile 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtLosses 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Text            =   "# Losses"
      ToolTipText     =   "# Losses"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtWins 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "# Wins"
      ToolTipText     =   "# Wins"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   7560
      TabIndex        =   28
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtInfo 
      Height          =   1215
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   3120
      Width           =   5175
   End
   Begin VB.TextBox txtHeadshots 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "# Headshots"
      ToolTipText     =   "# Headshots"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtDays 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "Days Played"
      ToolTipText     =   "Days Played"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtNumPoints 
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Text            =   "# Points Total"
      ToolTipText     =   "# Points Total"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtAssists 
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Text            =   "# Assists"
      ToolTipText     =   "# Assists"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtMinutes 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "Minutes Played"
      ToolTipText     =   "Minutes Played"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtHours 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "Hours Played"
      ToolTipText     =   "Hours Played"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H00000000&
      Caption         =   "10"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   6480
      TabIndex        =   24
      Top             =   3120
      Width           =   495
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H00000000&
      Caption         =   "9"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   6480
      TabIndex        =   23
      Top             =   2880
      Width           =   375
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H00000000&
      Caption         =   "8"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   6480
      TabIndex        =   22
      Top             =   2640
      Width           =   375
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H00000000&
      Caption         =   "7"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   6480
      TabIndex        =   21
      Top             =   2400
      Width           =   375
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H00000000&
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   6480
      TabIndex        =   20
      Top             =   2160
      Width           =   375
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H00000000&
      Caption         =   "5"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   5880
      TabIndex        =   19
      Top             =   3360
      Width           =   375
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H80000012&
      Caption         =   "4"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5880
      TabIndex        =   18
      Top             =   3120
      Width           =   375
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H80000012&
      Caption         =   "3"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   5880
      TabIndex        =   17
      Top             =   2880
      Width           =   375
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H80000012&
      Caption         =   "2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5880
      TabIndex        =   16
      Top             =   2640
      Width           =   375
   End
   Begin VB.OptionButton optPrestige 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5880
      TabIndex        =   15
      Top             =   2400
      Width           =   375
   End
   Begin VB.Frame fraPrestige 
      BackColor       =   &H00000000&
      Caption         =   "Prestige and Level"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   5640
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Text            =   "Level"
         ToolTipText     =   "Level"
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   5640
      TabIndex        =   26
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtDeaths 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "# Deaths"
      ToolTipText     =   "# Deaths"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtKills 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "# Kills"
      ToolTipText     =   "# Kills"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Image imgPrestige 
      Height          =   1650
      Left            =   7440
      Top             =   2040
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actions"
      Begin VB.Menu mnuCalculate 
         Caption         =   "Calculate"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "Game"
      Begin VB.Menu mnuCOD4 
         Caption         =   "COD4"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuCOD5 
         Caption         =   "COD5"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuCOD6 
         Caption         =   "MW2"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmCODStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Kills, Deaths, Assits, and Headshots
Private dblKills As Double
Private dblDeaths As Double
Private dblHeadshots As Double
Private dblAssists As Double
Private dblRatio As Double
Private dblDifference As Double
Private dblHeadshotsPerKill As Double
Private dblAssistsPerKill As Double

'Games: Wins and Losses
Private dblGames As Double
Private dblWins As Double
Private dblLosses As Double
Private dblWinLossRatio As Double

'Values Per Match
Private dblKillsPerMatch As Double
Private dblDeathsPerMatch As Double
Private dblHeadshotsPerMatch As Double
Private dblAssistsPerMatch As Double
Private dblShotsPerMatch As Double
Private dblShotsHitPerMatch As Double
Private dblShotsMissedPerMatch As Double
Private dblMinutesPerMatch As Double
Private dblPointsPerMatch As Double

'Values Per Minute
Private dblKillsPerMinute As Double
Private dblDeathsPerMinute As Double
Private dblHeadshotsPerMinute As Double
Private dblAssistsPerMinute As Double
Private dblShotsPerMinute As Double
Private dblShotsHitPerMinute As Double
Private dblShotsMissedPerMinute As Double
Private dblPointsPerMinute As Double

'Values Per Level
Private dblKillsPerLevel As Double
Private dblDeathsPerLevel As Double
Private dblHeadshotsPerLevel As Double
Private dblAssistsPerLevel As Double
Private dblLevelPercentage As Double
Private dblHoursPerLevel As Double
Private dblMinutesPerLevel As Double
Private dblSecondsPerLevel As Double
Private dblShotsPerLevel As Double
Private dblShotsHitPerLevel As Double
Private dblShotsMissedPerLevel As Double
Private dblGamesPerLevel As Double
Private dblGamesWonPerLevel As Double
Private dblGamesLostPerLevel As Double
Private dblPointsPerLevel As Double

'Values Per Prestige
Private dblDaysPerPrestige As Double
Private dblHoursPerPrestige As Double
Private dblMinutesPerPrestige As Double
Private dblSecondsPerPrestige As Double
Private dblKillsPerPrestige As Double
Private dblDeathsPerPrestige As Double
Private dblHeadshotsPerPrestige As Double
Private dblAssistsPerPrestige As Double
Private dblPointsPerPrestige As Double
Private dblShotsPerPrestige As Double
Private dblShotsHitPerPrestige As Double
Private dblShotsMissedPerPrestige As Double
Private dblGamesPerPrestige As Double
Private dblGamesWonPerPrestige As Double
Private dblGamesLostPerPrestige As Double

'Shots Fired
Private dblTotalShots As Double
Private dblShotsHit As Double
Private dblShotsMissed As Double
Private dblPercentHit As Double
Private dblPercentMissed As Double

'Time
Private dblTime As Double
Private dblYears As Double
Private dblMonths As Double
Private dblWeeks As Double
Private dblDays As Double
Private dblHours As Double
Private dblMinutes As Double
Private dblSeconds As Double

'Calculated Time
Private dblCalcDays As Double
Private dblCalcHours As Double

'Prestige, Points, and Level
Private intPrestigeNumber As Integer 'Actual Prestige Level for Preventing Dividing by 0. 0 is Actually the 1st Prestige.
Private dblPrestigePercentage As String
Private dblLevel As Double
Private dblTotalLevel As Double
Private dblPoints As Double

'GUI Items
Private intOption As Integer
Private intPrestige As Integer 'Visual Prestige Level for Saving and Loading and Assigning Image Visablity.
Private intGUI As Integer

'Save and Open
Private intFile As Integer

'Mode Switching
Private strImagePath As String
Private strGameMode As String
Private Sub cmdCalculate_Click()
    Call subCalculate
End Sub
Private Sub cmdPrint_Click()
    Call subPrint
End Sub
Private Sub cmdReset_Click()
    Call subReset
End Sub
Sub subPrint()
    Printer.Print txtInfo.Text
    MsgBox "Page will print when program is closed.", vbInformation, "Print OK!"
    cmdPrint.Enabled = False
End Sub
Sub subReset()
'Clear ALL Variables
    dblKills = 0
    dblDeaths = 0
    dblHeadshots = 0
    dblAssists = 0
    dblRatio = 0
    dblDifference = 0
    dblHeadshotsPerKill = 0
    dblAssistsPerKill = 0
    dblGames = 0
    dblWins = 0
    dblLosses = 0
    dblWinLossRatio = 0
    dblKillsPerMatch = 0
    dblDeathsPerMatch = 0
    dblHeadshotsPerMatch = 0
    dblAssistsPerMatch = 0
    dblShotsPerMatch = 0
    dblShotsHitPerMatch = 0
    dblShotsMissedPerMatch = 0
    dblMinutesPerMatch = 0
    dblPointsPerMatch = 0
    dblKillsPerMinute = 0
    dblDeathsPerMinute = 0
    dblHeadshotsPerMinute = 0
    dblAssistsPerMinute = 0
    dblShotsPerMinute = 0
    dblShotsHitPerMinute = 0
    dblShotsMissedPerMinute = 0
    dblPointsPerMinute = 0
    dblTotalLevel = 0
    dblKillsPerLevel = 0
    dblDeathsPerLevel = 0
    dblHeadshotsPerLevel = 0
    dblAssistsPerLevel = 0
    dblLevelPercentage = 0
    dblHoursPerLevel = 0
    dblMinutesPerLevel = 0
    dblSecondsPerLevel = 0
    dblShotsPerLevel = 0
    dblShotsHitPerLevel = 0
    dblShotsMissedPerLevel = 0
    dblGamesPerLevel = 0
    dblGamesWonPerLevel = 0
    dblGamesLostPerLevel = 0
    dblDaysPerPrestige = 0
    dblHoursPerPrestige = 0
    dblMinutesPerPrestige = 0
    dblSecondsPerPrestige = 0
    dblKillsPerPrestige = 0
    dblDeathsPerPrestige = 0
    dblHeadshotsPerPrestige = 0
    dblAssistsPerPrestige = 0
    dblPointsPerPrestige = 0
    dblShotsPerPrestige = 0
    dblShotsHitPerPrestige = 0
    dblShotsMissedPerPrestige = 0
    dblGamesPerPrestige = 0
    dblGamesWonPerPrestige = 0
    dblGamesLostPerPrestige = 0
    dblTotalShots = 0
    dblShotsHit = 0
    dblShotsMissed = 0
    dblPercentHit = 0
    dblPercentMissed = 0
    dblTime = 0
    dblYears = 0
    dblMonths = 0
    dblWeeks = 0
    dblDays = 0
    dblHours = 0
    dblMinutes = 0
    dblSeconds = 0
    dblCalcDays = 0
    dblCalcHours = 0
    intPrestigeNumber = 0
    dblPrestigePercentage = 0
    dblLevel = 0
    dblPoints = 0
    intOption = 0
    intPrestige = 0
    intGUI = 0
    
    'Clear ALL File Paths
    cdiFile.FileName = ""
    
    'Hide ALL Prestige Images
    imgPrestige.Picture = LoadPicture("")

    'Deselect ALL Option Buttons
    For intOption = 0 To 10
        optPrestige(intOption).Value = False
    Next
    
    'Reset ALL Text Boxes
    txtKills.Text = "# Kills"
    txtDeaths.Text = "# Deaths"
    txtHeadshots.Text = "# Headshots"
    txtAssists.Text = "# Assists"
    txtNumPoints.Text = "# Points Total"
    txtDays.Text = "Days Played"
    txtHours.Text = "Hours Played"
    txtMinutes.Text = "Minutes Played"
    txtWins.Text = "# Wins"
    txtLosses.Text = "# Losses"
    txtShotsHit.Text = "# Shots Hit"
    txtShotsMissed.Text = "# Shots Missed"
    txtLevel.Text = "Level"
    txtInfo.Text = ""
    
    'Reset Buttons
    cmdPrint.Enabled = False
End Sub
Private Sub Form_Load()
    intFile = FreeFile
    strImagePath = App.Path & "\Images\"
    Call subModeCOD6
End Sub
Private Sub mnuAbout_Click()
    MsgBox "Version: " & App.Major & "." & App.Minor & vbCrLf & "Build Date: 02/28/10" & vbCrLf & _
    "Program Written by Jackson (obogobo)." & vbCrLf & vbCrLf & "With Thanks to... " & vbCrLf & _
    "Jamie (Shell Shokked): Concepts / Ideas" & vbCrLf & "Philippe (TheBelgianDips): Beta Testing" & vbCrLf & _
    "Pat (TheFRY364): Beta Testing" & vbCrLf & vbCrLf & "Infinity Ward, Treyarch, Activision" & vbCrLf & "And the creators of the various images." & vbCrLf & vbCrLf & _
    "Team [XviD] FTW!", vbInformation, "About"
End Sub
Private Sub mnuCalculate_Click()
    Call subCalculate
End Sub
Private Sub mnuCOD4_Click()
    Call subReset
    Call subModeCOD4
End Sub
Private Sub mnuCOD5_Click()
    Call subReset
    Call subModeCOD5
End Sub
Private Sub mnuCOD6_Click()
    Call subReset
    Call subModeCOD6
End Sub
Private Sub mnuExit_Click()
    Unload Me
End Sub
Private Sub mnuPrint_Click()
    Call subPrint
End Sub
Private Sub mnuReset_Click()
    Call subReset
End Sub
Private Sub mnuSave_Click()
    cdiFile.FileName = ""
    cdiFile.InitDir = App.Path
    cdiFile.ShowSave
    
    If cdiFile.FileName <> "" Then
        Open cdiFile.FileName For Output As #intFile
        MsgBox "Wrote to " & cdiFile.FileName, vbInformation, "Save OK!"
        Print #intFile, dblKills, dblDeaths, dblHeadshots, dblAssists, dblPoints, dblDays, dblHours, dblMinutes, dblWins, dblLosses, dblShotsHit, dblShotsMissed, dblLevel, intPrestige
        Close #intFile
    End If
End Sub
Private Sub mnuOpen_Click()
    cdiFile.FileName = ""
    cdiFile.InitDir = App.Path
    cdiFile.ShowOpen
    
    If cdiFile.FileName <> "" Then
        Open cdiFile.FileName For Input As #intFile
        MsgBox "Read from " & cdiFile.FileName, vbInformation, "Load OK!"
        Input #intFile, dblKills, dblDeaths, dblHeadshots, dblAssists, dblPoints, dblDays, dblHours, dblMinutes, dblWins, dblLosses, dblShotsHit, dblShotsMissed, dblLevel, intPrestige
        Close #intFile
        txtKills.Text = dblKills
        txtDeaths.Text = dblDeaths
        txtHeadshots.Text = dblHeadshots
        txtAssists.Text = dblAssists
        txtDays.Text = dblDays
        txtHours.Text = dblHours
        txtMinutes.Text = dblMinutes
        txtNumPoints.Text = dblPoints
        txtWins.Text = dblWins
        txtLosses.Text = dblLosses
        txtShotsHit.Text = dblShotsHit
        txtShotsMissed.Text = dblShotsMissed
        txtLevel.Text = dblLevel
        optPrestige(intPrestige).Value = True
    End If
End Sub
Private Sub optPrestige_Click(Index As Integer)
    intPrestige = 0
    intPrestigeNumber = 0

    intPrestige = optPrestige(Index).Index
    intPrestigeNumber = intPrestige + 1
    
    If intPrestige <> 0 Then
        imgPrestige.Picture = LoadPicture(strImagePath & strGameMode & "-Prestige" & optPrestige(Index).Index & ".jpg")
    Else
        imgPrestige.Picture = LoadPicture("")
    End If
End Sub
Private Sub txtAssists_GotFocus()
    If txtAssists.Text = "# Assists" Then
        txtAssists.Text = ""
    End If
End Sub
Private Sub txtAssists_LostFocus()
    If txtAssists.Text = "" Then
        txtAssists.Text = "# Assists"
    End If
End Sub
Private Sub txtDays_Change()
    txtInfo.Text = ""
    dblDays = 0
    If IsNumeric(txtDays.Text) = True Then
        dblDays = txtDays.Text
    End If
End Sub
Private Sub txtDays_GotFocus()
    If txtDays.Text = "Days Played" Then
        txtDays.Text = ""
    End If
End Sub
Private Sub txtDays_LostFocus()
    If txtDays.Text = "" Then
        txtDays.Text = "Days Played"
    End If
End Sub
Private Sub txtDeaths_LostFocus()
    If txtDeaths.Text = "" Then
        txtDeaths.Text = "# Deaths"
    End If
End Sub
Private Sub txtDeaths_GotFocus()
    If txtDeaths.Text = "# Deaths" Then
        txtDeaths.Text = ""
    End If
End Sub
Private Sub txtHeadshots_GotFocus()
    If txtHeadshots.Text = "# Headshots" Then
        txtHeadshots.Text = ""
    End If
End Sub
Private Sub txtHeadshots_LostFocus()
    If txtHeadshots.Text = "" Then
        txtHeadshots.Text = "# Headshots"
    End If
End Sub
Private Sub txtHours_Change()
    txtInfo.Text = ""
    dblHours = 0
    If IsNumeric(txtHours.Text) = True Then
        dblHours = txtHours.Text
    End If
End Sub
Private Sub txtHours_GotFocus()
    If txtHours.Text = "Hours Played" Then
        txtHours.Text = ""
    End If
End Sub
Private Sub txtHours_LostFocus()
    If txtHours.Text = "" Then
        txtHours.Text = "Hours Played"
    End If
End Sub
Private Sub txtKills_GotFocus()
    If txtKills.Text = "# Kills" Then
        txtKills.Text = ""
    End If
End Sub
Private Sub txtKills_LostFocus()
    If txtKills.Text = "" Then
        txtKills.Text = "# Kills"
    End If
End Sub
Private Sub txtLevel_Change()
    txtInfo.Text = ""
    dblLevel = 0
    If IsNumeric(txtLevel.Text) = True Then
        If txtLevel.Text >= 1 And txtLevel.Text <= 55 And strGameMode = "COD4" Then
            dblLevel = txtLevel.Text
        ElseIf txtLevel.Text >= 1 And txtLevel.Text <= 65 And strGameMode = "COD5" Then
            dblLevel = txtLevel.Text
        ElseIf txtLevel.Text >= 1 And txtLevel.Text <= 70 And strGameMode = "COD6" Then
            dblLevel = txtLevel.Text
        Else
            Call LevelCheck
        End If
    Else
        Call LevelCheck
    End If
End Sub
Private Sub txtLevel_GotFocus()
    If txtLevel.Text = "Level" Then
        txtLevel.Text = ""
    End If
End Sub
Private Sub txtLevel_LostFocus()
    If txtLevel.Text = "" Then
        txtLevel.Text = "Level"
    End If
End Sub
Private Sub txtLosses_Change()
    txtInfo.Text = ""
    dblLosses = 0
    If IsNumeric(txtLosses.Text) = True Then
        dblLosses = txtLosses.Text
    End If
End Sub
Private Sub txtLosses_GotFocus()
    If txtLosses.Text = "# Losses" Then
        txtLosses.Text = ""
    End If
End Sub
Private Sub txtLosses_LostFocus()
    If txtLosses.Text = "" Then
        txtLosses.Text = "# Losses"
    End If
End Sub
Private Sub txtMinutes_Change()
    txtInfo.Text = ""
    dblMinutes = 0
    If IsNumeric(txtMinutes.Text) = True Then
        dblMinutes = txtMinutes.Text
    End If
End Sub
Private Sub txtAssists_Change()
    txtInfo.Text = ""
    dblAssists = 0
    If IsNumeric(txtAssists.Text) = True Then
        dblAssists = txtAssists.Text
    End If
End Sub
Private Sub txtDeaths_Change()
    txtInfo.Text = ""
    dblDeaths = 0
    If IsNumeric(txtDeaths.Text) = True Then
        dblDeaths = txtDeaths.Text
    End If
End Sub
Private Sub txtHeadshots_Change()
    txtInfo.Text = ""
    dblHeadshots = 0
    If IsNumeric(txtHeadshots.Text) = True Then
        dblHeadshots = txtHeadshots.Text
    End If
End Sub
Private Sub txtKills_Change()
    txtInfo.Text = ""
    dblKills = 0
    
    If IsNumeric(txtKills.Text) = True Then
        dblKills = txtKills.Text
    End If
End Sub
Private Sub txtMinutes_GotFocus()
    If txtMinutes.Text = "Minutes Played" Then
        txtMinutes.Text = ""
    End If
End Sub
Private Sub txtMinutes_LostFocus()
    If txtMinutes.Text = "" Then
        txtMinutes.Text = "Minutes Played"
    End If
End Sub
Private Sub txtNumPoints_Change()
    txtInfo.Text = ""
    dblPoints = 0
    If IsNumeric(txtNumPoints.Text) = True Then
        dblPoints = txtNumPoints.Text
    End If
End Sub
Private Sub txtNumPoints_GotFocus()
    If txtNumPoints.Text = "# Points Total" Then
        txtNumPoints.Text = ""
    End If
End Sub
Private Sub txtNumPoints_LostFocus()
    If txtNumPoints.Text = "" Then
        txtNumPoints.Text = "# Points Total"
    End If
End Sub
Private Sub txtShotsHit_Change()
    txtInfo.Text = ""
    dblShotsHit = 0
    If IsNumeric(txtShotsHit.Text) = True Then
        dblShotsHit = txtShotsHit.Text
    End If
End Sub
Private Sub txtShotsHit_GotFocus()
    If txtShotsHit.Text = "# Shots Hit" Then
        txtShotsHit.Text = ""
    End If
End Sub
Private Sub txtShotsHit_LostFocus()
    If txtShotsHit.Text = "" Then
        txtShotsHit.Text = "# Shots Hit"
    End If
End Sub
Private Sub txtShotsMissed_Change()
    txtInfo.Text = ""
    dblShotsMissed = 0
    If IsNumeric(txtShotsMissed.Text) = True Then
        dblShotsMissed = txtShotsMissed.Text
    End If
End Sub
Private Sub txtShotsMissed_GotFocus()
    If txtShotsMissed.Text = "# Shots Missed" Then
        txtShotsMissed.Text = ""
    End If
End Sub
Private Sub txtShotsMissed_LostFocus()
    If txtShotsMissed.Text = "" Then
        txtShotsMissed.Text = "# Shots Missed"
    End If
End Sub
Private Sub txtWins_Change()
    txtInfo.Text = ""
    dblWins = 0
    If IsNumeric(txtWins.Text) = True Then
        dblWins = txtWins.Text
    End If
End Sub
Sub subModeCOD4()
    strGameMode = "COD4"
    With frmCODStats
        .Icon = LoadPicture(strImagePath & strGameMode & ".ico")
        .Caption = "COD4: Modern Warfare, Stats Calculator"
        .Picture = LoadPicture(strImagePath & strGameMode & "-Background.jpg")
        .BackColor = vbBlack
    End With
    
    For intGUI = 0 To 10
        optPrestige(intGUI).BackColor = vbBlack
        optPrestige(intGUI).ForeColor = vbWhite
    Next
    
    fraPrestige.BackColor = vbBlack
    fraPrestige.ForeColor = vbWhite
    
    cdiFile.DialogTitle = "Select COD4 Stats File..."
    cdiFile.Filter = "Call of Duty 4 Data Files (*.cod4)|*cod4"
    cdiFile.DefaultExt = "cod4"
    
    mnuCOD4.Checked = True
    mnuCOD5.Checked = False
    mnuCOD6.Checked = False
End Sub
Sub subModeCOD5()
    strGameMode = "COD5"
    With frmCODStats
        .Icon = LoadPicture(strImagePath & strGameMode & ".ico")
        .Caption = "COD5: World at War, Stats Calculator"
        .Picture = LoadPicture(strImagePath & strGameMode & "-Background.jpg")
        .BackColor = vbWhite
    End With
    
    For intGUI = 0 To 10
        optPrestige(intGUI).BackColor = vbWhite
        optPrestige(intGUI).ForeColor = vbBlack
    Next
    
    fraPrestige.BackColor = vbWhite
    fraPrestige.ForeColor = vbBlack
    
    cdiFile.DialogTitle = "Select COD5 Stats File..."
    cdiFile.Filter = "Call of Duty 5 Data Files (*.cod5)|*.cod5"
    cdiFile.DefaultExt = "cod5"
    
    mnuCOD4.Checked = False
    mnuCOD5.Checked = True
    mnuCOD6.Checked = False
End Sub
Sub subModeCOD6()
    strGameMode = "COD6"
    With frmCODStats
        .Icon = LoadPicture(strImagePath & strGameMode & ".ico")
        .Caption = "COD6: Modern Warfare 2, Stats Calculator"
        .Picture = LoadPicture(strImagePath & strGameMode & "-Background.jpg")
        .BackColor = vbBlack
    End With
    
    For intGUI = 0 To 10
        optPrestige(intGUI).BackColor = vbBlack
        optPrestige(intGUI).ForeColor = vbWhite
    Next
    
    fraPrestige.BackColor = vbBlack
    fraPrestige.ForeColor = vbWhite
    
    cdiFile.DialogTitle = "Select MW2 Stats File..."
    cdiFile.Filter = "Modern Warfare 2 Data Files (*.mw2)|*mw2"
    cdiFile.DefaultExt = "mw2"
    
    mnuCOD4.Checked = False
    mnuCOD5.Checked = False
    mnuCOD6.Checked = True
End Sub
Sub subCalculate()
    'Prevent Crash and Burn...
    If dblKills <> 0 And dblDeaths <> 0 And dblAssists <> 0 And dblHeadshots <> 0 And dblPoints <> 0 And dblWins <> 0 And dblLosses <> 0 And dblShotsHit <> 0 And dblShotsMissed <> 0 And intPrestigeNumber <> 0 And dblLevel <> 0 Then
        
        'Calculate Statistics
        dblRatio = dblKills / dblDeaths
        dblDifference = dblKills - dblDeaths
        dblHeadshotsPerKill = dblHeadshots / dblKills
        dblAssistsPerKill = dblAssists / dblKills
        dblGames = dblWins + dblLosses
        dblWinLossRatio = dblWins / dblLosses
        dblKillsPerMatch = dblKills / dblGames
        dblDeathsPerMatch = dblDeaths / dblGames
        dblHeadshotsPerMatch = dblHeadshots / dblGames
        dblAssistsPerMatch = dblAssists / dblGames
        dblPointsPerMatch = dblPoints / dblGames
        dblTime = (dblDays * 1440) + (dblHours * 60) + dblMinutes
        dblYears = dblTime * 0.00000190132588
        dblMonths = dblTime * 0.0000228159105
        dblWeeks = dblTime * 0.0000992063492
        dblCalcDays = dblTime * 0.000694444444
        dblCalcHours = dblTime * 0.0166666667
        dblSeconds = dblTime * 60
        dblKillsPerMinute = dblKills / dblTime
        dblDeathsPerMinute = dblDeaths / dblTime
        dblHeadshotsPerMinute = dblHeadshots / dblTime
        dblAssistsPerMinute = dblAssists / dblTime
        dblPointsPerMinute = dblPoints / dblTime
        dblPointsPerMatch = dblPoints / dblGames
        dblTotalShots = dblShotsHit + dblShotsMissed
        dblPercentHit = dblShotsHit / dblTotalShots
        dblPercentMissed = dblShotsMissed / dblTotalShots
        dblShotsPerMatch = dblTotalShots / dblGames
        dblShotsPerMinute = dblTotalShots / dblTime
        dblShotsHitPerMatch = dblShotsHit / dblGames
        dblShotsMissedPerMatch = dblShotsMissed / dblGames
        dblShotsHitPerMinute = dblShotsHit / dblTime
        dblShotsMissedPerMinute = dblShotsMissed / dblTime
        
        If strGameMode = "COD4" Then
            dblTotalLevel = (intPrestige * 55) + dblLevel
        ElseIf strGameMode = "COD5" Then
            dblTotalLevel = (intPrestige * 65) + dblLevel
        ElseIf strGameMode = "COD6" Then
            dblTotalLevel = (intPrestige * 70) + dblLevel
        End If
        
        dblKillsPerLevel = dblKills / dblTotalLevel
        dblDeathsPerLevel = dblDeaths / dblTotalLevel
        dblHeadshotsPerLevel = dblHeadshots / dblTotalLevel
        dblAssistsPerLevel = dblAssists / dblTotalLevel
        
        If strGameMode = "COD4" Then
            dblLevelPercentage = dblTotalLevel / 605
        ElseIf strGameMode = "COD5" Then
            dblLevelPercentage = dblTotalLevel / 715
        ElseIf strGameMode = "COD6" Then
            dblLevelPercentage = dblTotalLevel / 770
        End If
        
        dblHoursPerLevel = dblCalcHours / dblTotalLevel
        dblMinutesPerLevel = dblTime / dblTotalLevel
        dblSecondsPerLevel = dblSeconds / dblTotalLevel
        dblShotsPerLevel = dblTotalShots / dblTotalLevel
        dblShotsHitPerLevel = dblShotsHit / dblTotalLevel
        dblShotsMissedPerLevel = dblShotsMissed / dblTotalLevel
        dblGamesPerLevel = dblGames / dblTotalLevel
        dblGamesWonPerLevel = dblWins / dblTotalLevel
        dblGamesLostPerLevel = dblLosses / dblTotalLevel
        dblPointsPerLevel = dblPoints / dblTotalLevel
        dblMinutesPerMatch = dblTime / dblGames
        dblDaysPerPrestige = dblCalcDays / intPrestigeNumber
        dblHoursPerPrestige = dblCalcHours / intPrestigeNumber
        dblMinutesPerPrestige = dblTime / intPrestigeNumber
        dblSecondsPerPrestige = dblSeconds / intPrestigeNumber
        dblKillsPerPrestige = dblKills / intPrestigeNumber
        dblDeathsPerPrestige = dblDeaths / intPrestigeNumber
        dblHeadshotsPerPrestige = dblHeadshots / intPrestigeNumber
        dblAssistsPerPrestige = dblAssists / intPrestigeNumber
        dblPointsPerPrestige = dblPoints / intPrestigeNumber
        dblShotsPerPrestige = dblTotalShots / intPrestigeNumber
        dblShotsHitPerPrestige = dblShotsHit / intPrestigeNumber
        dblShotsMissedPerPrestige = dblShotsMissed / intPrestigeNumber
        dblGamesPerPrestige = dblGames / intPrestigeNumber
        dblGamesWonPerPrestige = dblWins / intPrestigeNumber
        dblGamesLostPerPrestige = dblLosses / intPrestigeNumber
        dblPrestigePercentage = intPrestigeNumber / 11 '0 to 10 is Actually 11 Prestige Levels
        
        'Format Statistics
        dblRatio = FormatNumber(dblRatio, 4)
        dblHeadshotsPerKill = FormatNumber(dblHeadshotsPerKill, 2)
        dblAssistsPerKill = FormatNumber(dblAssistsPerKill, 2)
        dblWinLossRatio = FormatNumber(dblWinLossRatio, 3)
        dblKillsPerMatch = FormatNumber(dblKillsPerMatch, 2)
        dblDeathsPerMatch = FormatNumber(dblDeathsPerMatch, 2)
        dblHeadshotsPerMatch = FormatNumber(dblHeadshotsPerMatch, 2)
        dblAssistsPerMatch = FormatNumber(dblAssistsPerMatch, 2)
        dblPointsPerMatch = FormatNumber(dblPointsPerMatch, 3)
        dblYears = FormatNumber(dblYears, 4)
        dblMonths = FormatNumber(dblMonths, 2)
        dblWeeks = FormatNumber(dblWeeks, 2)
        dblCalcDays = FormatNumber(dblCalcDays, 2)
        dblCalcHours = FormatNumber(dblCalcHours, 2)
        dblKillsPerMinute = FormatNumber(dblKillsPerMinute, 2)
        dblDeathsPerMinute = FormatNumber(dblDeathsPerMinute, 2)
        dblHeadshotsPerMinute = FormatNumber(dblHeadshotsPerMinute, 2)
        dblAssistsPerMinute = FormatNumber(dblAssistsPerMinute, 2)
        dblPointsPerMinute = FormatNumber(dblPointsPerMinute, 3)
        dblPointsPerMatch = FormatNumber(dblPointsPerMatch, 3)
        dblPercentHit = FormatNumber(dblPercentHit, 3)
        dblPercentMissed = FormatNumber(dblPercentMissed, 3)
        dblShotsPerMatch = FormatNumber(dblShotsPerMatch, 2)
        dblShotsPerMinute = FormatNumber(dblShotsPerMinute, 2)
        dblShotsHitPerMatch = FormatNumber(dblShotsHitPerMatch, 2)
        dblShotsMissedPerMatch = FormatNumber(dblShotsMissedPerMatch, 2)
        dblShotsHitPerMinute = FormatNumber(dblShotsHitPerMinute, 2)
        dblShotsMissedPerMinute = FormatNumber(dblShotsMissedPerMinute, 2)
        dblMinutesPerMatch = FormatNumber(dblMinutesPerMatch, 2)
        dblTotalLevel = FormatNumber(dblTotalLevel, 2)
        dblKillsPerLevel = FormatNumber(dblKillsPerLevel, 2)
        dblDeathsPerLevel = FormatNumber(dblDeathsPerLevel, 2)
        dblHeadshotsPerLevel = FormatNumber(dblHeadshotsPerLevel, 2)
        dblAssistsPerLevel = FormatNumber(dblAssistsPerLevel, 2)
        dblLevelPercentage = FormatNumber(dblLevelPercentage, 2)
        dblHoursPerLevel = FormatNumber(dblHoursPerLevel, 2)
        dblMinutesPerLevel = FormatNumber(dblMinutesPerLevel, 2)
        dblSecondsPerLevel = FormatNumber(dblSecondsPerLevel, 2)
        dblShotsPerLevel = FormatNumber(dblShotsPerLevel, 2)
        dblShotsHitPerLevel = FormatNumber(dblShotsHitPerLevel, 2)
        dblShotsMissedPerLevel = FormatNumber(dblShotsMissedPerLevel, 2)
        dblGamesPerLevel = FormatNumber(dblGamesPerLevel, 2)
        dblGamesWonPerLevel = FormatNumber(dblGamesWonPerLevel, 2)
        dblGamesLostPerLevel = FormatNumber(dblGamesLostPerLevel, 2)
        dblPointsPerLevel = FormatNumber(dblPointsPerLevel, 3)
        dblDaysPerPrestige = FormatNumber(dblDaysPerPrestige, 2)
        dblHoursPerPrestige = FormatNumber(dblHoursPerPrestige, 2)
        dblMinutesPerPrestige = FormatNumber(dblMinutesPerPrestige, 2)
        dblSecondsPerPrestige = FormatNumber(dblSecondsPerPrestige, 2)
        dblKillsPerPrestige = FormatNumber(dblKillsPerPrestige, 2)
        dblDeathsPerPrestige = FormatNumber(dblDeathsPerPrestige, 2)
        dblHeadshotsPerPrestige = FormatNumber(dblHeadshotsPerPrestige, 2)
        dblAssistsPerPrestige = FormatNumber(dblAssistsPerPrestige, 2)
        dblPointsPerPrestige = FormatNumber(dblPointsPerPrestige, 3)
        dblShotsPerPrestige = FormatNumber(dblShotsPerPrestige, 2)
        dblShotsHitPerPrestige = FormatNumber(dblShotsHitPerPrestige, 2)
        dblShotsMissedPerPrestige = FormatNumber(dblShotsMissedPerPrestige, 2)
        dblGamesPerPrestige = FormatNumber(dblGamesPerPrestige, 2)
        dblGamesWonPerPrestige = FormatNumber(dblGamesWonPerPrestige, 2)
        dblGamesLostPerPrestige = FormatNumber(dblGamesLostPerPrestige, 2)
        dblPrestigePercentage = FormatPercent(dblPrestigePercentage, 1)
        
        'Display Info
        txtInfo.Text = strGameMode & " Statistics:" & vbCrLf & vbCrLf & "Total Kills: " & dblKills & vbCrLf & "Total Deaths: " & dblDeaths & vbCrLf & "Total Headshots: " & dblHeadshots & vbCrLf & "Total Assists: " & dblAssists & vbCrLf & "Total Points: " & dblPoints & vbCrLf & "Kill to Death Ratio: " & dblRatio & vbCrLf & "Kill to Death Difference: " & dblDifference & vbCrLf & "Headshots per Kill: " & dblHeadshotsPerKill & vbCrLf & "Assists per Kill: " & dblAssistsPerKill & vbCrLf & vbCrLf & _
        "Average Kills per Match: " & dblKillsPerMatch & vbCrLf & "Average Deaths per Match: " & dblDeathsPerMatch & vbCrLf & "Average Headshots per Match: " & dblHeadshotsPerMatch & vbCrLf & "Average Assists per Match: " & dblAssistsPerMatch & vbCrLf & "Average Minutes per Match: " & dblMinutesPerMatch & vbCrLf & "Average Points per Match: " & dblPointsPerMatch & vbCrLf & vbCrLf & _
        "Average Kills per Minute: " & dblKillsPerMinute & vbCrLf & "Average Deaths per Minute: " & dblDeathsPerMinute & vbCrLf & "Average Headshots per Minute: " & dblHeadshotsPerMinute & vbCrLf & "Average Assists per Minute: " & dblAssistsPerMinute & vbCrLf & "Average Points per Minute: " & dblPointsPerMinute & vbCrLf & vbCrLf & _
        "Total Shots Taken: " & dblTotalShots & vbCrLf & "Percent of Shots Hit: " & dblPercentHit & vbCrLf & "Percent of Shots Missed: " & dblPercentMissed & vbCrLf & _
        "Average Shots Fired per Match: " & dblShotsPerMatch & vbCrLf & "Average Shots Hit per Match: " & dblShotsHitPerMatch & vbCrLf & "Average Shots Missed per Match: " & dblShotsMissedPerMatch & vbCrLf & vbCrLf & _
        "Average Shots Fired per Minute: " & dblShotsPerMinute & vbCrLf & "Average Shots Hit per Minute: " & dblShotsHitPerMinute & vbCrLf & "Average Shots Missed per Minute: " & dblShotsMissedPerMinute & vbCrLf & vbCrLf & _
        "Total Time Played..." & vbCrLf & "Years: " & dblYears & vbCrLf & "Months: " & dblMonths & vbCrLf & "Weeks: " & dblWeeks & vbCrLf & "Days: " & dblCalcDays & vbCrLf & "Hours: " & dblCalcHours & vbCrLf & "Minutes: " & dblTime & vbCrLf & "Seconds: " & dblSeconds & vbCrLf & vbCrLf & _
        "Total Games Played: " & dblGames & vbCrLf & "Games Won: " & dblWins & vbCrLf & "Games Lost: " & dblLosses & vbCrLf & "Win to Loss Ratio: " & dblWinLossRatio & vbCrLf & vbCrLf & _
        "Prestige Level: " & intPrestige & vbCrLf & "Prestige Percentage: " & dblPrestigePercentage & vbCrLf & "Average Days per Prestige: " & dblDaysPerPrestige & vbCrLf & "Average Hours per Prestige: " & dblHoursPerPrestige & vbCrLf & "Average Minutes per Prestige: " & dblMinutesPerPrestige & vbCrLf & "Average Seconds per Prestige: " & dblSecondsPerPrestige & vbCrLf & _
        "Average Kills per Prestige: " & dblKillsPerPrestige & vbCrLf & "Average Deaths per Prestige: " & dblDeathsPerPrestige & vbCrLf & "Average Headshots per Prestige: " & dblHeadshotsPerPrestige & vbCrLf & "Average Assists per Prestige: " & dblAssistsPerPrestige & vbCrLf & "Average Points per Prestige: " & dblPointsPerPrestige & vbCrLf & "Average Shots Fired per Prestige: " & dblShotsPerPrestige & vbCrLf & "Average Shots Hit per Prestige: " & dblShotsHitPerPrestige & vbCrLf & "Average Shots Missed per Prestige: " & dblShotsMissedPerPrestige & vbCrLf & "Average Matches Played per Prestige: " & dblGamesPerPrestige & vbCrLf & "Average Matches Won per Prestige: " & dblGamesWonPerPrestige & vbCrLf & "Average Matches Lost per Prestige: " & dblGamesLostPerPrestige & vbCrLf & vbCrLf & _
        "Total Level: " & dblTotalLevel & vbCrLf & "Average Kills per Level: " & dblKillsPerLevel & vbCrLf & "Average Deaths per Level: " & dblDeathsPerLevel & vbCrLf & "Average Headshots per Level: " & dblHeadshotsPerLevel & vbCrLf & "Average Assists per Level: " & dblAssistsPerLevel & vbCrLf & "Average Points per Level: " & dblPointsPerLevel & vbCrLf & "Level Percentage: " & dblLevelPercentage & vbCrLf & "Average Hours per Level: " & dblHoursPerLevel & vbCrLf & "Average Minutes per Level: " & dblMinutesPerLevel & vbCrLf & "Average Seconds per Level: " & dblSecondsPerLevel & vbCrLf & "Average Shots Fired per Level: " & dblShotsPerLevel & vbCrLf & "Average Shots Hit per Level: " & dblShotsHitPerLevel & vbCrLf & "Average Shots Missed per Level: " & dblShotsMissedPerLevel & vbCrLf & "Average Games Played per Level: " & dblGamesPerLevel & vbCrLf & "Average Games Won per Level: " & dblGamesWonPerLevel & vbCrLf & "Average Games Lost per Level: " & dblGamesLostPerLevel
        
        'Enable Printing
        cmdPrint.Enabled = True
    Else
        'Call subReset
        txtInfo.Text = "Enter a valid numeric value into each field."
    End If
End Sub
Private Sub txtWins_GotFocus()
    If txtWins.Text = "# Wins" Then
        txtWins.Text = ""
    End If
End Sub
Private Sub txtWins_LostFocus()
    If txtWins.Text = "" Then
        txtWins.Text = "# Wins"
    End If
End Sub
Sub LevelCheck()
    If strGameMode = "COD4" Then
        txtInfo.Text = "Enter a level between 1 and 55."
    ElseIf strGameMode = "COD5" Then
        txtInfo.Text = "Enter a level between 1 and 65."
    ElseIf strGameMode = "COD6" Then
        txtInfo.Text = "Enter a level between 1 and 70."
    End If
End Sub
