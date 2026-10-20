'bbbbbbbbb          aaaaa       ttttttttttttttttt   mmmm       mmmm       aaaaa       nnnn         nnnn   
'bbbbbbbbbb       aaaa aaaa     ttttttttttttttttt   mmmmm     mmmmm     aaaa aaaa     nnnnn        nnnn
'bbbbbbbbbbb     aaaa   aaaa    ttttttttttttttttt   mmmmmm   mmmmmm    aaaa   aaaa    nnnnnn       nnnn 
'bbb      bbb   aaaa     aaaa        ttttttt        mmmmmmm mmmmmmm   aaaa     aaaa   nnnnnnn      nnnn
'bbb      bbb   aaaa     aaaa        ttttttt        mmmmmmmmmmmmmmm   aaaa     aaaa   nnnn nnn     nnnn
'bbb      bbb   aaaa     aaaa        ttttttt        mmmm  mmm  mmmm   aaaa     aaaa   nnnn  nnn    nnnn
'bbbbbbbbbbb    aaaa     aaaa        ttttttt        mmmm   m   mmmm   aaaa     aaaa   nnnn   nnn   nnnn
'bbbbbbbbbbb    aaaaaaaaaaaaa        ttttttt        mmmm       mmmm   aaaaaaaaaaaaa   nnnn    nnn  nnnn
'bbb      bbb   aaaaaaaaaaaaa        ttttttt        mmmm       mmmm   aaaaaaaaaaaaa   nnnn     nnn nnnn
'bbb      bbb   aaaa     aaaa        ttttttt        mmmm       mmmm   aaaa     aaaa   nnnn      nnnnnnn
'bbb      bbb   aaaa     aaaa        ttttttt        mmmm       mmmm   aaaa     aaaa   nnnn       nnnnnn
'bbbbbbbbbbb    aaaa     aaaa        ttttttt        mmmm       mmmm   aaaa     aaaa   nnnn        nnnnn
'bbbbbbbbbb     aaaa     aaaa        ttttttt        mmmm       mmmm   aaaa     aaaa   nnnn         nnnn
'bbbbbbbbb      aaaa     aaaa        ttttttt        mmmm       mmmm   aaaa     aaaa   nnnn         nnnn
'                                             ....`    `...`   `....`....``....`   ....`   `.  `.    ``    .....
'Original Table by Ricky Dunigan AKA Dunri Wikan 45                              
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/



'  Credits
'Thank you to my wife and kids for letting me get away for a few hours (or a lot of hours) a day to work on this project
'Thank you to John Memeo for keeping me motivated and listening to every little idea I had and testing
'This is my 2nd original table
'Thanks to Scotty Wick for providing the framework and answering my questions

'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' Compiled and built by Scottywic
' Code snippets by JP Salas, Flupper, DJRobX, NailBuster, ninuzzu, HauntFreaks rothbauerw & many more.
' I'll attempt to make ammends for anyone left out.
' For details using this framework check https://www.youtube.com/channel/UC_vroB2Uboi8GytbQSIepPQ


	Option Explicit
	Randomize


'**************************
'   PinUp Player USER Config
'**************************

dim PuPDMDDriverType: PuPDMDDriverType=2   ' 0=LCD DMD, 1=RealDMD 2=FULLDMD (large/High LCD)
dim useRealDMDScale : useRealDMDScale=0    ' 0 or 1 for RealDMD scaling.  Choose which one you prefer.
dim useDMDVideos    : useDMDVideos=True   ' true or false to use DMD splash videos.

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  Player Options
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'orbital scoreboard
	
Dim osbactive:osbactive = 1 'set to 0 for off, 1 for only player 1 to be sent, 2 for all scores to be sent.
	Dim osbid:osbid ="DunriWikan45" ' your orbital scoreboard login name
	Dim osbkey:osbkey="c4dc9669-559b-11ea-a042-42010a8a06b6" ' your orbital scoreboard api key
	Dim osbdefinit:osbdefinit = "RGD" ' your default initials to use 
	'standard player rules
	turnoffrules = 1 ' change to 1 to take off the backglass helper rules text during a game
	turnonultradmd = 0 ' change to 1 to turn on ultradmd
	toppervideo = 0 'set to 1 to turn on the topper
	ballrolleron = 1 ' set to 0 to turn off the ball roller if you use the "c" key in your cabinet






'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  FRAMEWORK Options
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  

	Const typefont = "Raleway Medium"
	Const numberfont = "Bebas Neue"
	Const zoomfont = "Fundamental  Brigade"
	Const zoombgfont = "Fundamental 3D  Brigade" ' needs to be an outline of the zoomfont
	Const cGameName = "BatmanTAS"
	Const TableName = "batmantas"
	Const myVersion = "1.0.0"
	


'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'FRAMEWORK BASE CODE STARTS HERE
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  VARIABLES
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 	

	'Constructions
	Const BallSize = 50
	Const MaxPlayers = 1
	Const BallSaverTime = 15 
	Const MaxMultiplier = 6 
	Const MaxMultiballs = 5
	Const bpgcurrent = 5

	' Define Global Variables

	Dim toppervideo
	Dim ballrolleron
	Dim turnonultradmd
	Dim turnoffrules
	Dim PlayersPlayingGame
	Dim CurrentPlayer
	Dim Credits
	Dim BonusPoints(4)
	Dim BonusHeldPoints(4)
	Dim BonusMultiplier(4)
	Dim bBonusHeld
	Dim BallsRemaining(4)
	Dim ExtraBallsAwards(4)
	Dim Score(4)
	Dim HighScore(4)
	Dim HighScoreName(4)
	Dim WaffleScore(4)
	Dim WaffleScoreName(4)
	Dim Jackpot
	Dim SuperJackpot
	Dim Tilt
	Dim TiltSensitivity
	Dim Tilted
	Dim TotalGamesPlayed
	Dim mBalls2Eject
	Dim SkillshotValue(4)
	Dim bAutoPlunger
	Dim bInstantInfo
	Dim bromconfig
	Dim bAttractMode
	Dim LastSwitchHit
	Dim BallsOnPlayfield
	Dim BallsInHole
	Dim bFreePlay
	Dim bGameInPlay
	Dim bOnTheFirstBall
	Dim bBallInPlungerLane
	Dim bBallSaverActive
	Dim bBallSaverReady
	Dim bMultiBallMode
	Dim bMusicOn
	Dim bSkillshotReady
	Dim bExtraBallWonThisBall
	Dim bJustStarted

	LoadCoreFiles
	Sub LoadCoreFiles
		On Error Resume Next
		ExecuteGlobal GetTextFile("core.vbs")
		If Err Then MsgBox "Can't open core.vbs"
		On Error Goto 0
	End Sub

	Dim EnableBallControl
	EnableBallControl = False 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  Controller VBS stuff, but with b2s not started
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'***Controller.vbs version 1.2***'
'

Const directory = "HKEY_CURRENT_USER\SOFTWARE\Visual Pinball\Controller\"
Dim objShell
Dim PopupMessage
Dim B2SController
Dim Controller
Const DOFContactors = 1
Const DOFKnocker = 2
Const DOFChimes = 3
Const DOFBell = 4
Const DOFGear = 5
Const DOFShaker = 6
Const DOFFlippers = 7
Const DOFTargets = 8
Const DOFDropTargets = 9
Const DOFOff = 0
Const DOFOn = 1
Const DOFPulse = 2

Dim DOFeffects(9)
Dim B2SOn
Dim B2SOnALT

Sub LoadEM
	LoadController("EM")
End Sub

Sub LoadPROC(VPMver, VBSfile, VBSver)
	LoadVBSFiles VPMver, VBSfile, VBSver
	LoadController("PROC")
End Sub

Sub LoadVPM(VPMver, VBSfile, VBSver)
	LoadVBSFiles VPMver, VBSfile, VBSver
	LoadController("VPM")
End Sub

Sub LoadVPMALT(VPMver, VBSfile, VBSver)
	LoadVBSFiles VPMver, VBSfile, VBSver
	LoadController("VPMALT")
End Sub

Sub LoadVBSFiles(VPMver, VBSfile, VBSver)
	On Error Resume Next
	If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
	ExecuteGlobal GetTextFile(VBSfile)
	If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description	
	InitializeOptions
End Sub

Sub LoadVPinMAME
	Set Controller = CreateObject("VPinMAME.Controller")
	If Err Then MsgBox "Can't load VPinMAME." & vbNewLine & Err.Description
	If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
	If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
	On Error Goto 0
End Sub

Sub LoadController(TableType)
	Dim FileObj
	Dim DOFConfig
	Dim TextStr2
	Dim tempC
	Dim count
	Dim ISDOF
	Dim Answer
	
	B2SOn = False
	B2SOnALT = False
	tempC = 0
	on error resume next
	Set objShell = CreateObject("WScript.Shell")
	objShell.RegRead(directory & "ForceDisableB2S")
	If Err.number <> 0 Then
		PopupMessage = "This latest version of Controller.vbs stores its settings in the registry. To adjust the values, you must use VP 10.2 (or newer) and setup your configuration in the DOF section of the -Keys, Nudge and DOF- dialog of Visual Pinball."
		objShell.RegWrite directory & "ForceDisableB2S",0, "REG_DWORD"
		objShell.RegWrite directory & "DOFContactors",2, "REG_DWORD"
		objShell.RegWrite directory & "DOFKnocker",2, "REG_DWORD"
		objShell.RegWrite directory & "DOFChimes",2, "REG_DWORD"
		objShell.RegWrite directory & "DOFBell",2, "REG_DWORD"
		objShell.RegWrite directory & "DOFGear",2, "REG_DWORD"
		objShell.RegWrite directory & "DOFShaker",2, "REG_DWORD"
		objShell.RegWrite directory & "DOFFlippers",2, "REG_DWORD"
		objShell.RegWrite directory & "DOFTargets",2, "REG_DWORD"
		objShell.RegWrite directory & "DOFDropTargets",2, "REG_DWORD"
		MsgBox PopupMessage
	End If
	tempC = objShell.RegRead(directory & "ForceDisableB2S")
	DOFeffects(1)=objShell.RegRead(directory & "DOFContactors")
	DOFeffects(2)=objShell.RegRead(directory & "DOFKnocker")
	DOFeffects(3)=objShell.RegRead(directory & "DOFChimes")
	DOFeffects(4)=objShell.RegRead(directory & "DOFBell")
	DOFeffects(5)=objShell.RegRead(directory & "DOFGear")
	DOFeffects(6)=objShell.RegRead(directory & "DOFShaker")
	DOFeffects(7)=objShell.RegRead(directory & "DOFFlippers")
	DOFeffects(8)=objShell.RegRead(directory & "DOFTargets")
	DOFeffects(9)=objShell.RegRead(directory & "DOFDropTargets")
	Set objShell = nothing

	If TableType = "PROC" or TableType = "VPMALT" Then
		If TableType = "PROC" Then
			Set Controller = CreateObject("VPROC.Controller")
			If Err Then MsgBox "Can't load PROC"
		Else
			LoadVPinMAME
		End If		
		If tempC = 0 Then
			On Error Resume Next
			If Controller is Nothing Then
				Err.Clear
			Else
				Set B2SController = CreateObject("B2S.Server")
				If B2SController is Nothing Then
					Err.Clear
				Else
					B2SController.B2SName = B2ScGameName
					B2SController.Run()
					On Error Goto 0
					B2SOn = True
					B2SOnALT = True
				End If
			End If
		End If
	Else
		If tempC = 0 Then
			On Error Resume Next
			Set Controller = CreateObject("B2S.Server")
			If Controller is Nothing Then
				Err.Clear
				If TableType = "VPM" Then 
					LoadVPinMAME
				End If
			Else
				Controller.B2SName = cGameName
				If TableType = "EM" Then
					Controller.Run()
				End If
				On Error Goto 0
				B2SOn = True
			End If
		Else
			If TableType = "VPM" Then 
				LoadVPinMAME
			End If
		End If
		Set DOFConfig=Nothing
		Set FileObj=Nothing
	End If
End sub

Function SoundFX (Sound, Effect)
	If ((Effect = 0 And B2SOn) Or DOFeffects(Effect)=1) Then
		SoundFX = ""
	Else
		SoundFX = Sound
	End If
End Function

Function SoundFXDOF (Sound, DOFevent, State, Effect)
	If DOFeffects(Effect)=1 Then
		SoundFXDOF = ""
		DOF DOFevent, State
	ElseIf DOFeffects(Effect)=2 Then
		SoundFXDOF = Sound
		DOF DOFevent, State
	Else
		SoundFXDOF = Sound
	End If
End Function

Function SoundFXDOFALT (Sound, DOFevent, State, Effect)
	If DOFeffects(Effect)=1 Then
		SoundFXDOFALT = ""
		DOFALT DOFevent, State
	ElseIf DOFeffects(Effect)=2 Then
		SoundFXDOFALT = Sound
		DOFALT DOFevent, State
	Else
		SoundFXDOFALT = Sound
	End If
End Function


Sub DOF(DOFevent, State)
	If B2SOn Then
		If State = 2 Then
			Controller.B2SSetData DOFevent, 1:Controller.B2SSetData DOFevent, 0
		Else
			Controller.B2SSetData DOFevent, State
		End If
	End If
End Sub

Sub DOFALT(DOFevent, State)
	If B2SOnALT Then
		If State = 2 Then
			B2SController.B2SSetData DOFevent, 1:B2SController.B2SSetData DOFevent, 0
		Else
			B2SController.B2SSetData DOFevent, State
		End If
	End If
End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   TABLE INITS & MATHS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  


	Sub Table1_Init()
		resetbackglass
		DMD_Init
		LoadEM
		Dim i
		Randomize
		Loadhs
		bAttractMode = False
		bOnTheFirstBall = False
		bBallInPlungerLane = False
		bBallSaverActive = False
		bBallSaverReady = False
		bMultiBallMode = False
		bGameInPlay = False
		bAutoPlunger = False
		bMusicOn = True
		BallsOnPlayfield = 0
		BallsInHole = 0
		LastSwitchHit = ""
		Tilt = 0
		TiltSensitivity = 6
		Tilted = False
		bBonusHeld = False
		bJustStarted = True
		GiOff
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":4, ""FS"":1 }"
		StartAttractMode
	End Sub


	'********************
	' MATHS
	'********************

	Function RndNum(min,max)
	 RndNum = Int(Rnd()*(max-min+1))+min     ' Sets a random number between min AND max
	End Function


		Dim BIP
		BIP = 0







'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   KEYS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  


	Sub Table1_KeyDown(ByVal Keycode)

if keycode = LeftFlipperKey Then
LeftFlipper.RotateToEnd
ShiftLanesLeft ' for lanes
'PlaySound "FlipperUp"
End If

If keycode = RightFlipperKey Then
RightFlipper.RotateToEnd
ShiftLanesRight ' for lanes
'PlaySound "FlipperUp"
End If

	If ballrolleron = 1 then
		if keycode = 46 then ' C Key
			 If contball = 1 Then
				  contball = 0
			 Else
				  contball = 1
			 End If
		End If
	End If
	if keycode = 48 then 'B Key
		 If bcboost = 1 Then
			  bcboost = bcboostmulti
		 Else
			  bcboost = 1
		 End If
	End If
	if keycode = 203 then bcleft = 1 ' Left Arrow
	if keycode = 200 then bcup = 1 ' Up Arrow
	if keycode = 208 then bcdown = 1 ' Down Arrow
	if keycode = 205 then bcright = 1 ' Right Arrow


		If keycode = PlungerKey Then
			PlaySoundAt "fx_plungerpull", Plunger
			Plunger.Pullback
		End If

		If hsbModeActive = True Then
			EnterHighScoreKey(keycode)
			elseif bGameInPlay Then
			If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25:CheckTilt
			If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25:CheckTilt
			If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound SoundFX("fx_nudge",0), 0, 1, 1, 0.25:CheckTilt
			If NOT Tilted Then
			If keycode = LeftFlipperKey Then
				LeftFlipper.RotateToEnd
				PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
				ldown = 1
				checkdown
				If bSkillshotReady = False Then
				RotateLaneLightsLeft
				End If
			End If
			If keycode = RightFlipperKey Then 
				RightFlipper.RotateToEnd
				PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
				rdown = 1
				checkdown
				If bSkillshotReady = False Then
				RotateLaneLightsRight
				End If
			End If

			If keycode = StartGameKey Then
				If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then
						PlayersPlayingGame = PlayersPlayingGame + 1
						PuPlayer.playlistplayex pCallouts,"audiocallouts","player1.wav",80,1
						chilloutthemusic

						If PlayersPlayingGame = 2 Then
							PuPlayer.LabelSet pBackglass,"Play2","PLAYER 2",1,"{'mt':2}"
							pUpdateScores
						PuPlayer.playlistplayex pCallouts,"audiocallouts","player2.wav",80,1
						chilloutthemusic
						End If

						If PlayersPlayingGame = 3 Then
							PuPlayer.LabelSet pBackglass,"Play3","PLAYER 3",1,"{'mt':2}"
							pUpdateScores
						PuPlayer.playlistplayex pCallouts,"audiocallouts","player3.wav",80,1
						chilloutthemusic
						End If

						If PlayersPlayingGame = 4 Then
							PuPlayer.LabelSet pBackglass,"Play4","PLAYER 4",1,"{'mt':2}"
							pUpdateScores	
						PuPlayer.playlistplayex pCallouts,"audiocallouts","player4.wav",80,1
						chilloutthemusic
						End If

						TotalGamesPlayed = TotalGamesPlayed + 1
						savegp
				End If
			End If
			End If
			Else
			If NOT Tilted Then

				If keycode = LeftFlipperKey Then helptime.enabled = true:DMDintroloop:introtime = 0
				If keycode = RightFlipperKey Then helptime.enabled = true:DMDintroloop:introtime = 0
					If keycode = StartGameKey Then
							If(BallsOnPlayfield = 0) Then
								ResetForNewGame()
							End If
					End If
			End If
			End If 

	'****************
	' Testing Keys  
	'****************
'	if keycode = "3" then
'		'addscore 18000000
'		bumps(CurrentPlayer) = bumps(CurrentPlayer) + 180
'		BumperRewards
'		CheckHighscore
'	End If
'
'	if keycode = "3" then
'		lm1.State = 1
'		lm2.State = 1
'		lm3.State = 1
'		lm5.State = 1
'		CheckMONSTER
'		skipscene
'	End If
	if keycode = "3" then
'	addscore 50000000

	End If
	End Sub


	Sub Table1_KeyUp(ByVal keycode)
	
    'Manual Ball Control

	if keycode = 203 then bcleft = 0 ' Left Arrow
	if keycode = 200 then bcup = 0 ' Up Arrow
	if keycode = 208 then bcdown = 0 ' Down Arrow
	if keycode = 205 then bcright = 0 ' Right Arrow

		If keycode = PlungerKey Then
			PlaySoundAt "fx_plunger", Plunger
			Plunger.Fire
		End If

		' Table specific

		If bGameInPLay and hsbModeActive <> True Then


			If keycode = LeftFlipperKey Then
				ldown = 0
				LeftFlipper.RotateToStart
				PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
			End If
			If keycode = RightFlipperKey Then
				rdown = 0
				RightFlipper.RotateToStart
				PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
			End If
		Else
			If keycode = LeftFlipperKey Then helptime.enabled = false
			If keycode = RightFlipperKey Then helptime.enabled = false
		End If

	End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  MANUAL BALLCONTROL
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
	' 

	Sub StartControl_Hit()
		 Set ControlBall = ActiveBall
		 contballinplay = False
	End Sub

	Sub StopControl_Hit()
		 contballinplay = False
	End Sub

	Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
	Dim bcvel, bcyveloffset, bcboostmulti

	bcboost = 1 'Do Not Change - default setting
	bcvel = 4 'Controls the speed of the ball movement
	bcyveloffset = -0.01 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
	bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)

	Sub BallControl_Timer()
		 If Contball and ContBallInPlay then
			  If bcright = 1 Then
				   ControlBall.velx = bcvel*bcboost
			  ElseIf bcleft = 1 Then
				   ControlBall.velx = - bcvel*bcboost
			  Else
				   ControlBall.velx=0
			  End If

			 If bcup = 1 Then
				  ControlBall.vely = -bcvel*bcboost
			 ElseIf bcdown = 1 Then
				  ControlBall.vely = bcvel*bcboost
			 Else
				  ControlBall.vely= bcyveloffset
			 End If
		 End If
	End Sub






'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   SHADOWS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  


	'*****************************************
	'	ninuzzu's	FLIPPER SHADOWS
	'*****************************************

	sub FlipperTimer_Timer()
		FlipperLSh.RotZ = LeftFlipper.currentangle
		FlipperRSh.RotZ = RightFlipper.currentangle

	End Sub

	'*****************************************
	'	ninuzzu's	BALL SHADOW
	'*****************************************
	Dim BallShadow
	BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

	Sub BallShadowUpdate_timer()
		Dim BOT, b
		BOT = GetBalls
		' hide shadow of deleted balls
		If UBound(BOT)<(tnob-1) Then
			For b = (UBound(BOT) + 1) to (tnob-1)
				BallShadow(b).visible = 0
			Next
		End If
		' exit the Sub if no balls on the table
		If UBound(BOT) = -1 Then Exit Sub
		' render the shadow for each ball
		For b = 0 to UBound(BOT)
			If BOT(b).X < Table1.Width/2 Then
				BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
			Else
				BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
			End If
			ballShadow(b).Y = BOT(b).Y + 12
			If BOT(b).Z > 20 Then
				BallShadow(b).visible = 1
			Else
				BallShadow(b).visible = 0
			End If
		Next
	End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   TILT
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  


	'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

	Sub CheckTilt                                    'Called when table is nudged
		Tilt = Tilt + TiltSensitivity                'Add to tilt count
		TiltDecreaseTimer.Enabled = True
		If(Tilt> TiltSensitivity) AND(Tilt <15) Then 'show a warning
			pNote "CAREFUL!","TILT WARNING"
				'PlaySound "buzz"
			PuPlayer.playlistplayex pBackglass,"videotilt","",100,1
			DOF 131, DOFPulse
			DOF 311, DOFPulse  'DOF MX - Tilt Warning
		End if
		If Tilt> 15 Then 'If more that 15 then TILT the table
			Tilted = True
			pNote "TILT",""
				'PlaySound "powerdownn"
			PuPlayer.playlistplayex pBackglass,"videotilt","",100,4
			DOF 310, DOFPulse   'DOF MX - TILT
			DOF 127, DOFOff   'DOF - Beacon - OFF
			DisableTable True
			tilttableclear.enabled = true
			TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
		End If
	End Sub

	Dim tilttime:tilttime = 0

	sub tilttableclear_timer
		tilttime = tilttime + 1
		Select Case tilttime
			Case 10
				tableclearing
		End Select
	End Sub

	Sub tableclearing

	End Sub

	Sub posttiltreset

	End Sub

	Sub TiltDecreaseTimer_Timer
		' DecreaseTilt
		If Tilt> 0 Then
			Tilt = Tilt - 0.1
		Else
			TiltDecreaseTimer.Enabled = False
		End If
	End Sub

	Sub DisableTable(Enabled)
		If Enabled Then
			GiOff
			LightSeqTilt.Play SeqAllOff
			LeftFlipper.RotateToStart
			RightFlipper.RotateToStart
			LeftSlingshot.Disabled = 1
			RightSlingshot.Disabled = 1
			PuPlayer.playresume 4
			PuPlayer.playlistplayex pAudio,"audiomodes","clear.mp3",100,1
			PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100, 1
		Else
			GiOn
			LightSeqTilt.StopPlay
			LeftSlingshot.Disabled = 0
			RightSlingshot.Disabled = 0
		End If
	End Sub

	Sub TiltRecoveryTimer_Timer()
		If(BallsOnPlayfield = 0) Then
			EndOfBall()
			TiltRecoveryTimer.Enabled = False
		End If
	End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   SOUND FUNCTIONS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  


	'*********************************************************************
	'                 Positional Sound Playback Functions
	'*********************************************************************

	' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
	' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
	' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
	Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
		PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
	End Sub

	' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
	Sub PlaySoundAt(soundname, tableobj)
		PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
	End Sub

	Sub PlaySoundAtBall(soundname)
		PlaySoundAt soundname, ActiveBall
	End Sub


	'*********************************************************************
	'                     Supporting Ball & Sound Functions
	'*********************************************************************

	Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
		Dim tmp
		tmp = tableobj.y * 2 / table1.height-1
		If tmp > 0 Then
			AudioFade = Csng(tmp ^10)
		Else
			AudioFade = Csng(-((- tmp) ^10) )
		End If
	End Function

	Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
		Dim tmp
		tmp = tableobj.x * 2 / table1.width-1
		If tmp > 0 Then
			AudioPan = Csng(tmp ^10)
		Else
			AudioPan = Csng(-((- tmp) ^10) )
		End If
	End Function

	Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
		Vol = Csng(BallVel(ball) ^2 / 2000)
	End Function

	Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
		Pitch = BallVel(ball) * 20
	End Function

	Function BallVel(ball) 'Calculates the ball speed
		BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
	End Function


	'*****************************************
	'      JP's VP10 Rolling Sounds
	'*****************************************

	Const tnob = 5 ' total number of balls
	ReDim rolling(tnob)
	InitRolling

	Sub InitRolling
		Dim i
		For i = 0 to tnob
			rolling(i) = False
		Next
	End Sub

	Sub RollingTimer_Timer()
		Dim BOT, b
		BOT = GetBalls

		' stop the sound of deleted balls
		For b = UBound(BOT) + 1 to tnob
			rolling(b) = False
			StopSound("fx_ballrolling" & b)
		Next

		' exit the sub if no balls on the table
		If UBound(BOT) = -1 Then Exit Sub

		' play the rolling sound for each ball
		For b = 0 to UBound(BOT)
			If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
				rolling(b) = True
				PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
			Else
				If rolling(b) = True Then
					StopSound("fx_ballrolling" & b)
					rolling(b) = False
				End If
			End If
		Next
	End Sub

	'**********************
	' Ball Collision Sound
	'**********************

	Sub OnBallBallCollision(ball1, ball2, velocity)
		PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
	End Sub

	'**************************************
	' Explanation of the collision routine
	'**************************************

	' The collision is built in VP.
	' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they 
	' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
	' depending of the speed of the collision.


	Sub Pins_Hit (idx)
		PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
	End Sub

	Sub Targets_Hit (idx)
		PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
	End Sub

	Sub Metals_Thin_Hit (idx)
		PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Sub

	Sub Metals_Medium_Hit (idx)
		PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Sub

	Sub Metals2_Hit (idx)
		PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Sub

	Sub Gates_Hit (idx)
		PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Sub

	Sub Spinner_Spin
		PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
	End Sub

	Sub Rubbers_Hit(idx)
		dim finalspeed
		finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
		If finalspeed > 20 then 
			PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		End if
		If finalspeed >= 6 AND finalspeed <= 20 then
			RandomSoundRubber()
		End If
	End Sub

	Sub Posts_Hit(idx)
		dim finalspeed
		finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
		If finalspeed > 16 then 
			PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		End if
		If finalspeed >= 6 AND finalspeed <= 16 then
			RandomSoundRubber()
		End If
	End Sub

	Sub RandomSoundRubber()
		Select Case Int(Rnd*3)+1
			Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
			Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
			Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		End Select
	End Sub

	Sub LeftFlipper_Collide(parm)
		RandomSoundFlipper()
	End Sub

	Sub RightFlipper_Collide(parm)
		RandomSoundFlipper()
	End Sub

	Sub RandomSoundFlipper()
		Select Case Int(Rnd*3)+1
			Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
			Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
			Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		End Select
	End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   START GAME, END GANE
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  

'
	Sub ResetForNewGame()
		Dim i
		bGameInPLay = True
		StopAttractMode
		GiOn

pDMDStartGame


questionmarksreset
riddlerreset
sw003Reset
sw004Reset
sw4Reset
resettargets
batrampreset
jokerrampreset

LightTrig6.State = 0
LightTrig7.State = 0
LightTrig8.State = 0
LightTrig9.State = 0


penguingl.state = lightstateblinking
freezegl.state = lightstateblinking
poisongl.state = lightstateblinking
banegl.state = lightstateblinking
jokergl.state = lightstateblinking
batgl.state = lightstateblinking
riddlergl.state = lightstateblinking
symbolgl.state = 0

sw777.IsDropped = False
sw77.IsDropped = False
sw7.IsDropped = False

LightTrig5.State = LightStateOff
LightTrig4.State = LightStateOff
LightTrig3.State = LightStateOff
LightTrig2.State = LightStateOff
LightTrig1.State = LightStateOn

goalwall.IsDropped = 0

		TotalGamesPlayed = TotalGamesPlayed + 1
		savegp
		CurrentPlayer = 1
		PlayersPlayingGame = 1
		bOnTheFirstBall = True
		For i = 1 To MaxPlayers
			Score(i) = 0
			BonusPoints(i) = 0
			BonusHeldPoints(i) = 0
			BonusMultiplier(i) = 1
			BallsRemaining(i) = 5
			ExtraBallsAwards(i) = 0
		Next
		Tilt = 0
		Game_Init()
		vpmtimer.addtimer 1500, "FirstBall '"
	End Sub


	Sub EndOfGame()
osbtempscore = Score(CurrentPlayer)
			if osbkey="" Then
			Else
			SubmitOSBScore
			end If
		PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100, 1
		pNote "GAME OVER","PLAY AGAIN"
		PuPlayer.playlistplayex pBackglass,"videogameover","",100,1
		StartAttractMode
		introposition = 0
		bGameInPLay = False
		bJustStarted = False
		Dim i
		GiOff

pDMDGameOver
questionmarksreset
sw003Reset
sw004Reset
sw4Reset
batrampreset
jokerrampreset

	sw2.IsDropped = False
	sw22.IsDropped = False
	sw222.IsDropped = False 

	Light013.State = LightStateBlinking
	Light011.State = LightStateBlinking
	Light012.State = LightStateBlinking
	LightA.State = LightStateBlinking
	LightB.State = LightStateBlinking
	LightC.State = LightStateBlinking

	LightTrig6.State = LightStateBlinking
	LightTrig7.State = LightStateBlinking
	LightTrig8.State = LightStateBlinking
	LightTrig9.State = LightStateBlinking

LightTrig5.State = LightStateBlinking
LightTrig4.State = LightStateBlinking
LightTrig3.State = LightStateBlinking
LightTrig2.State = LightStateBlinking
LightTrig1.State = LightStateBlinking

penguingl.state = LightStateBlinking
freezegl.state = LightStateBlinking
poisongl.state = LightStateBlinking
banegl.state = LightStateBlinking
jokergl.state = LightStateBlinking
batgl.state = LightStateBlinking
riddlergl.state = LightStateBlinking
symbolgl.state = LightStateBlinking


	End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   ULTRADMD SCRIPTING
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  

	Dim UltraDMD

	' DMD using UltraDMD calls

	Sub DMD(background, toptext, bottomtext, duration)
		If turnonultradmd = 0 then exit sub
		UltraDMD.DisplayScene00 background, toptext, 15, bottomtext, 15, 14, duration, 14
		UltraDMDTimer.Enabled = 1 'to show the score after the animation/message
	End Sub

	Sub DMDScore
		If turnonultradmd = 0 then exit sub
		UltraDMD.SetScoreboardBackgroundImage "scoreboard-background.jpg", 15, 7
		UltraDMD.DisplayScoreboard PlayersPlayingGame, CurrentPlayer, Score(1), Score(2), Score(3), Score(4), "Player " & CurrentPlayer, "Ball " & Balls
	End Sub

	Sub DMDScoreNow
		If turnonultradmd = 0 then exit sub
		DMDFlush
		DMDScore
	End Sub

	Sub DMDFLush
		If turnonultradmd = 0 then exit sub
		UltraDMDTimer.Enabled = 0
		UltraDMD.CancelRendering
	End Sub

	Sub DMDScrollCredits(background, text, duration)
		If turnonultradmd = 0 then exit sub
		UltraDMD.ScrollingCredits background, text, 15, 14, duration, 14
	End Sub

	Sub DMDId(id, background, toptext, bottomtext, duration)
		If turnonultradmd = 0 then exit sub
		UltraDMD.DisplayScene00ExwithID id, False, background, toptext, 15, 0, bottomtext, 15, 0, 14, duration, 14
	End Sub

	Sub DMDMod(id, toptext, bottomtext, duration)
		If turnonultradmd = 0 then exit sub
		UltraDMD.ModifyScene00Ex id, toptext, bottomtext, duration
	End Sub

	Sub UltraDMDTimer_Timer 'used for the attrackmode and the instant info.
		If turnonultradmd = 0 then exit sub
		If bInstantInfo Then
			InstantInfo
		ElseIf bAttractMode Then
		ElseIf NOT UltraDMD.IsRendering Then
			DMDScoreNow
		ElseIf bromconfig Then
			romconfig
		End If
	End Sub

	Sub DMD_Init
		If turnonultradmd = 0 then exit sub
		Set UltraDMD = CreateObject("UltraDMD.DMDObject")
		If UltraDMD is Nothing Then
			MsgBox "No UltraDMD found.  This table will NOT run without it."
			Exit Sub
		End If

		UltraDMD.Init
		If turnonultradmd = 0 then exit sub
		If Not UltraDMD.GetMajorVersion = 1 Then
			MsgBox "Incompatible Version of UltraDMD found."
			Exit Sub
		End If

		If UltraDMD.GetMinorVersion <1 Then
		If turnonultradmd = 0 then exit sub
			MsgBox "Incompatible Version of UltraDMD found. Please update to version 1.1 or newer."
			Exit Sub
		End If

		Dim fso:Set fso = CreateObject("Scripting.FileSystemObject")
		Dim curDir:curDir = fso.GetAbsolutePathName(".")

		Dim DirName
		DirName = curDir& "\" &TableName& ".UltraDMD"

		If Not fso.FolderExists(DirName) Then _
				Msgbox "UltraDMD userfiles directory '" & DirName & "' does not exist." & CHR(13) & "No graphic images will be displayed on the DMD"
		UltraDMD.SetProjectFolder DirName

		' wait for the animation to end
		While UltraDMD.IsRendering = True
		WEnd

	End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   Pinup Active Backglass
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  

'**************************
'   PinUp Player Config
'   Change HasPuP = True if using PinUp Player Videos
'**************************

	Dim HasPup:HasPuP = True

	Dim PuPlayer

	Const pTopper=0
	Const pDMD=1
	Const pBackglass=2
	Const pPlayfield=3
	Const pMusic=4
	Const pAudio=7
	Const pCallouts=8

'pages
Const pDMDBlank=0
Const pScores=1
Const pBigLine=2
Const pThreeLines=3
Const pTwoLines=4
Const pTargerLetters=5

'dmdType
Const pDMDTypeLCD=0
Const pDMDTypeReal=1
Const pDMDTypeFULL=2







dim PUPDMDObject  'for realtime mirroring.
Dim pDMDlastchk: pDMDLastchk= -1    'performance of updates
Dim pDMDCurPage: pDMDCurPage= 0     'default page is empty.
Dim pInAttract : pInAttract=false   'pAttract mode



	if HasPuP Then
	on error resume next
	Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay") 
	on error goto 0
	if not IsObject(PuPlayer) then HasPuP = False
	end If

	if HasPuP Then

	PuPlayer.Init pBackglass,cGameName
	PuPlayer.Init pMusic,cGameName
	PuPlayer.Init pAudio,cGameName
	PuPlayer.Init pCallouts,cGameName
	If toppervideo = 1 Then
	PuPlayer.Init pTopper,cGameName
	End If

	PuPlayer.SetScreenex pBackglass,0,0,0,0,0       'Set PuPlayer DMD TO Always ON    <screen number> , xpos, ypos, width, height, POPUP
	PuPlayer.SetScreenex pAudio,0,0,0,0,2
	PuPlayer.hide pAudio
	PuPlayer.SetScreenex pMusic,0,0,0,0,2
	PuPlayer.hide pMusic
	PuPlayer.SetScreenex pCallouts,0,0,0,0,2
	PuPlayer.hide pCallouts

	Sub chilloutthemusic
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":10 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":10 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":10 }"
		vpmtimer.addtimer 2200, "turnitbackup'"
	End Sub

	Sub chilloutthemusic2
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":30 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":30 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":30 }"
		vpmtimer.addtimer 1000, "turnitbackup3'"
	End Sub

	Sub chilloutthemusic5
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":10 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":10 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":00 }"
		vpmtimer.addtimer 12000, "turnitbackup5'"
	End Sub


	Sub chilloutthemusic4
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":0 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":0 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":0 }"
		vpmtimer.addtimer 20000, "turnitbackup2'"
	End Sub


	Sub turnitbackup
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":99 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":99 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":99 }"
	End Sub

	Sub turnitbackup2
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":99 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":99 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":99 }"
	End Sub

	Sub turnitbackup3
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":99 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":99 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":99 }"
	End Sub

	Sub turnitbackup5
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 2, ""FN"":11, ""VL"":99 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 4, ""FN"":11, ""VL"":99 }"
		PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 7, ""FN"":11, ""VL"":99 }"
	End Sub

	PuPlayer.playlistadd pBackglass, "videogameover", 1, 0
	PuPlayer.playlistadd pMusic,"audioattract", 1 , 0
	PuPlayer.playlistadd pAudio,"audioevents", 1 , 0
	PuPlayer.playlistadd pCallouts,"audiocallouts", 1 , 0
	PuPlayer.playlistadd pBackglass,"backglass", 1 , 0
	PuPlayer.playlistadd pBackglass,"PuPOverlays", 1 , 0
	PuPlayer.playlistadd pBackglass,"videoattract", 1 , 0
	PuPlayer.playlistadd pBackglass,"videodrain", 1 , 0
    PuPlayer.playlistadd pDMD,"PUPFrames", 1 , 0
	PuPlayer.playlistadd pBackglass,"pencap", 1 , 0
    PuPlayer.playlistadd pDMD,"pendmd", 1 , 0
 	PuPlayer.playlistadd pTopper,"pentop", 1 , 0
 	PuPlayer.playlistadd pTopper,"pentopjail", 1 , 0
	PuPlayer.playlistadd pBackglass,"ivycap", 1 , 0
    PuPlayer.playlistadd pDMD,"ivydmd", 1 , 0
 	PuPlayer.playlistadd pTopper,"ivytop", 1 , 0
 	PuPlayer.playlistadd pTopper,"ivytopjail", 1 , 0
	PuPlayer.playlistadd pBackglass,"frecap", 1 , 0
    PuPlayer.playlistadd pDMD,"fredmd", 1 , 0
 	PuPlayer.playlistadd pTopper,"fretop", 1 , 0
 	PuPlayer.playlistadd pTopper,"fretopjail", 1 , 0
	PuPlayer.playlistadd pBackglass,"carcap", 1 , 0
    PuPlayer.playlistadd pDMD,"cardmd", 1 , 0
 	PuPlayer.playlistadd pTopper,"cartop", 1 , 0
 	PuPlayer.playlistadd pTopper,"cartopjail", 1 , 0
	PuPlayer.playlistadd pBackglass,"jokcap", 1 , 0
    PuPlayer.playlistadd pDMD,"jokdmd", 1 , 0
 	PuPlayer.playlistadd pTopper,"joktop", 1 , 0
 	PuPlayer.playlistadd pTopper,"joktopjail", 1 , 0
    PuPlayer.playlistadd pDMD,"batarangdmd", 1 , 0
 	PuPlayer.playlistadd pTopper,"batarang", 1 , 0
	PuPlayer.playlistadd pBackglass,"ridcap", 1 , 0
    PuPlayer.playlistadd pDMD,"riddmd", 1 , 0
 	PuPlayer.playlistadd pTopper,"ridtop", 1 , 0
 	PuPlayer.playlistadd pTopper,"ridtopjail", 1 , 0
 	PuPlayer.playlistadd pTopper,"ridtop2", 1 , 0
	PuPlayer.playlistadd pBackglass,"ridback", 1 , 0
	PuPlayer.playlistadd pCallouts,"catwomancallouts", 1 , 0
	PuPlayer.playlistadd pCallouts,"robincallouts", 1 , 0
	PuPlayer.playlistadd pCallouts,"harleycallouts", 1 , 0
	PuPlayer.playlistadd pBackglass,"bancap", 1 , 0
    PuPlayer.playlistadd pDMD,"bandmd", 1 , 0
 	PuPlayer.playlistadd pTopper,"bantop", 1 , 0
 	PuPlayer.playlistadd pTopper,"bantopjail", 1 , 0
	PuPlayer.playlistadd pBackglass,"rogcap", 1 , 0
    PuPlayer.playlistadd pDMD,"rogdmd", 1 , 0
 	PuPlayer.playlistadd pTopper,"rogtop", 1 , 0
 	PuPlayer.playlistadd pTopper,"rogtopjail", 1 , 0
	PuPlayer.playlistadd pBackglass,"20mil", 1 , 0
    PuPlayer.playlistadd pDMD,"20mildmd", 1 , 0

	If toppervideo = 1 Then
 	PuPlayer.playlistadd pTopper,"topper", 1 , 0
	End If

	'Set Background video on DMD
		PuPlayer.playlistplayex pBackglass,"topper","topper.mp4",0,1
		PuPlayer.playlistplayex pBackglass,"PuPOverlays","overlay.png",0,1
		PuPlayer.SetBackground pBackglass,1	

		PuPlayer.playlistplayex pDMD,"PuPFrames","default.png",0,1
		PuPlayer.SetBackground pDMD,1


	End if

	If toppervideo = 1 Then
		PuPlayer.playlistplayex pTopper,"DMDAttract","dmdbackground.mp4",0,1
		PuPlayer.SetBackground pTopper,1	
	End If


	PuPlayer.LabelInit pBackglass

	'Setup Pages.  Note if you use fonts they must be in FONTS folder of the pupVideos\tablename\FONTS
	'syntax - PuPlayer.LabelNew <screen# or pDMD>,<Labelname>,<fontName>,<size%>,<colour>,<rotation>,<xAlign>,<yAlign>,<xpos>,<ypos>,<PageNum>,<visible>

	'Page 1 (default score display)
	PuPlayer.LabelNew pBackglass,"Play1",typefont,			3,16777215  ,0,0,1,23,81,1,0
	PuPlayer.LabelNew pBackglass,"Play1score",numberfont,	3,16777215  ,0,0,1,31,81,1,0
	PuPlayer.LabelNew pBackglass,"Play2",typefont,			3,16777215  ,0,0,1,23,85,1,0
	PuPlayer.LabelNew pBackglass,"Play2score",numberfont,	3,16777215  ,0,0,1,31,85,1,0
	PuPlayer.LabelNew pBackglass,"Play3",typefont,			3,16777215  ,0,0,1,23,89,1,0
	PuPlayer.LabelNew pBackglass,"Play3score",numberfont,	3,16777215  ,0,0,1,31,89,1,0
	PuPlayer.LabelNew pBackglass,"Play4",typefont,			3,16777215  ,0,0,1,23,93,1,0
	PuPlayer.LabelNew pBackglass,"Play4score",numberfont,	3,16777215  ,0,0,1,31,93,1,0
	PuPlayer.LabelNew pBackglass,"curscore",numberfont,		7,16777215	,0,1,1,50,87,1,1
	PuPlayer.LabelNew pBackglass,"Ball",typefont,			2,16777215 	,0,1,1,50,81,1,1
	PuPlayer.LabelNew pBackglass,"curplayer",typefont,		3,0		 	,0,1,1,50,93,1,1
	PuPlayer.LabelNew pBackglass,"hstitle",typefont,		3,444665 	,0,1,1,65,89,1,1
	PuPlayer.LabelNew pBackglass,"hs",numberfont,			3,16777215 	,0,1,1,65,93,1,1
	PuPlayer.LabelNew pBackglass,"gptitle",typefont,		3,8421504 	,0,1,1,77,89,1,1
	PuPlayer.LabelNew pBackglass,"gp",numberfont,			3,16777215 	,0,1,1,77,93,1,1
	PuPlayer.LabelNew pBackglass,"notetitle",numberfont,	3,16777215  ,0,1,1,50,12,1,1
	PuPlayer.LabelNew pBackglass,"notecopy",typefont,		2,16777215 	,0,1,1,50,15,1,1
	PuPlayer.LabelNew pBackglass,"titlebg",zoombgfont,		9,0  		,0,1,1,50,50,1,1
	PuPlayer.LabelNew pBackglass,"title",zoomfont,			9,16777215 	,0,1,1,50,50,1,1
	PuPlayer.LabelNew pBackglass,"titlebg2",zoombgfont,		6,0  		,0,1,1,50,50,1,1
	PuPlayer.LabelNew pBackglass,"title2",zoomfont,			6,16777215 	,0,1,1,50,50,1,1
	PuPlayer.LabelNew pBackglass,"lefttitle",numberfont,	3,0			,0,1,1,11,82,1,1
	PuPlayer.LabelNew pBackglass,"lefttimer",numberfont,	8,16777215  ,0,1,1,11,87,1,1
	PuPlayer.LabelNew pBackglass,"leftdetail",typefont,		2,0	  		,0,1,1,11,93,1,1
	PuPlayer.LabelNew pBackglass,"righttitle",numberfont,	3,0			,0,1,1,89,82,1,1
	PuPlayer.LabelNew pBackglass,"righttimer",numberfont,	8,16777215  ,0,1,1,89,87,1,1
	PuPlayer.LabelNew pBackglass,"rightdetail",typefont,	2,0	  		,0,1,1,89,93,1,1
	PuPlayer.LabelNew pBackglass,"high1name",typefont,		5,16777215  ,0,1,1,22,30,1,1
	PuPlayer.LabelNew pBackglass,"high1score",numberfont,	5,16777215  ,0,1,1,36,30,1,1
	PuPlayer.LabelNew pBackglass,"high2name",typefont,		5,16777215  ,0,1,1,22,38,1,1
	PuPlayer.LabelNew pBackglass,"high2score",numberfont,	5,16777215  ,0,1,1,36,38,1,1
	PuPlayer.LabelNew pBackglass,"high3name",typefont,		5,16777215  ,0,1,1,22,46,1,1
	PuPlayer.LabelNew pBackglass,"high3score",numberfont,	5,16777215  ,0,1,1,36,46,1,1
	PuPlayer.LabelNew pBackglass,"high4name",typefont,		5,16777215  ,0,1,1,22,54,1,1
	PuPlayer.LabelNew pBackglass,"high4score",numberfont,	5,16777215  ,0,1,1,36,54,1,1
	PuPlayer.LabelNew pBackglass,"HighScore",typefont,		6,16777215	,0,0,1,20,30,1,1
	PuPlayer.LabelNew pBackglass,"HighScoreL1",numberfont,	8,16777215	,0,0,1,20,40,1,1
	PuPlayer.LabelNew pBackglass,"HighScoreL2",numberfont,	8,16777215	,0,0,1,24,40,1,1
	PuPlayer.LabelNew pBackglass,"HighScoreL3",numberfont,	8,16777215	,0,0,1,28,40,1,1
	PuPlayer.LabelNew pBackglass,"HighScoreL4",numberfont,	4,16777215	,0,0,1,20,50,1,1



	Sub resetbackglass
	Loadhs
	PuPlayer.LabelShowPage pBackglass,1,0,""
	PuPlayer.playlistplayex pBackglass,"scene","layout.png",0,1
	PuPlayer.SetBackground pBackglass,1
	PuPlayer.LabelSet pBackglass,"hstitle","HIGH SCORE",1,""
	PuPlayer.LabelSet pBackglass,"hs","" & FormatNumber(HighScore(0),0),1,""
	PuPlayer.LabelSet pBackglass,"gptitle","GAMES",1,""
	PuPlayer.LabelSet pBackglass,"gp","" & FormatNumber(TotalGamesPlayed,0),1,""
	PuPlayer.LabelSet pBackglass,"Ball","PRESS\rSTART",1,""

	PuPlayer.LabelSet pBackglass,"lefttitle","DATA ITEM",1,""
	PuPlayer.LabelSet pBackglass,"lefttimer","455",1,""
	PuPlayer.LabelSet pBackglass,"leftdetail","PRESS\rSTART",1,""
	PuPlayer.LabelSet pBackglass,"righttitle","DATA ITEM",1,""
	PuPlayer.LabelSet pBackglass,"righttimer","455",1,""
	PuPlayer.LabelSet pBackglass,"rightdetail","PRESS\rSTART",1,""

	End Sub


	Dim titlepos
	titlepos = 0
	titletimer.enabled = 0
	dim title
	title = ""
	dim subtitle
	subtitle = ""

	Sub titletimer_timer
		titlepos = titlepos + 1
		Select Case titlepos
			Case 1
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':2565927, 'size': 0, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 0, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':2565927, 'size': 0, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 0, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 2
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':5066061, 'size': 0.6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 0.6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':5066061, 'size': 0.4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 0.4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 3
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':7960953, 'size': 1.2, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 1.2, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':7960953, 'size': 0.8, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 0.8, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 4
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':9671571, 'size': 1.8, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 1.8, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':9671571, 'size': 1.2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 1.2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 5
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':11842740, 'size': 2.4, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 2.4, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':11842740, 'size': 1.6, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 1.6, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 6
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':13224393, 'size': 3, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 3, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':13224393, 'size': 2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 7
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':14671839, 'size': 3.6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 3.6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':14671839, 'size': 2.4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 2.4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 8
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':15790320, 'size': 4.2, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 4.2, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':15790320, 'size': 2.8, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 2.8, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 9
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':16316664, 'size': 4.8, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 4.8, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':16316664, 'size': 3.2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 3.2, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 10
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':16777215, 'size': 5.4, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 5.4, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':16777215, 'size': 3.6, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 3.6, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 11
				PuPlayer.LabelSet pBackglass,"title",title,1,"{'mt':2,'color':16777215, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg",title,1,"{'mt':2,'color':0, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2",subtitle,1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2",subtitle,1,"{'mt':2,'color':0, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			Case 150
				PuPlayer.LabelSet pBackglass,"title","",1,"{'mt':2,'color':16777215, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg","",1,"{'mt':2,'color':0, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"title2","",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				PuPlayer.LabelSet pBackglass,"titlebg2","",1,"{'mt':2,'color':0, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
				titlepos = 0
				titletimer.enabled = 0
		End Select
	End Sub

	
	Sub pNote(msgText,msg2text)
		title = msgText
		subtitle = msg2text
		If titlepos = 0 Then
			titletimer.enabled = 1
		Else
			titlepos = 0
			titletimer.enabled = 1
			PuPlayer.LabelSet pBackglass,"title","",1,"{'mt':2,'color':16777215, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
			PuPlayer.LabelSet pBackglass,"titlebg","",1,"{'mt':2,'color':0, 'size': 6, 'xpos': 50, 'xalign': 1, 'ypos': 36, 'yalign': 0}"
			PuPlayer.LabelSet pBackglass,"title2","",1,"{'mt':2,'color':16777215, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
			PuPlayer.LabelSet pBackglass,"titlebg2","",1,"{'mt':2,'color':0, 'size': 4, 'xpos': 50, 'xalign': 1, 'ypos': 45, 'yalign': 0}"
		End If
	End Sub




	Sub currentplayerbackglass

	End Sub


	Sub pUpdateScores
		PuPlayer.LabelSet pBackglass,"curscore",FormatNumber(Score(CurrentPlayer),0),1,""
		PuPlayer.LabelSet pBackglass,"curplayer","Player " & CurrentPlayer,1,""
		If CurrentPlayer = 1 Then
			PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(CurrentPlayer),0),1,"{'mt':2,'color':16777215, 'size': 2.4 }"
			PuPlayer.LabelSet pBackglass,"Play1","Player 1",1,"{'mt':2,'color':16777215 }"
			'make other scores red (inactive)
			If PlayersPlayingGame = 2 Then
				PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play2","Player 2",1,"{'mt':2,'color':8421504}"
			End If
			If PlayersPlayingGame = 3 Then
				PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play2","Player 2",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play3","Player 3",1,"{'mt':2,'color':8421504}"
			End If
			If PlayersPlayingGame = 4 Then
				PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play2","Player 2",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play3","Player 3",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play4score","" & FormatNumber(Score(4),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play4","Player 4",1,"{'mt':2,'color':8421504}"
			End If
		End If
		If CurrentPlayer = 2 Then
			PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(CurrentPlayer),0),1,"{'mt':2,'color':16777215, 'size': 2.4 }"
			PuPlayer.LabelSet pBackglass,"Play2","Player 2",1,"{'mt':2,'color':16777215 }"
			If PlayersPlayingGame = 2 Then
				PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play1","Player 1",1,"{'mt':2,'color':8421504}"
			End If
			If PlayersPlayingGame = 3 Then
				PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play1","Player 1",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play3","Player 3",1,"{'mt':2,'color':8421504}"
			End If
			If PlayersPlayingGame = 4 Then
				PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play1","Player 1",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play3","Player 3",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play4score","" & FormatNumber(Score(4),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play4","Player 4",1,"{'mt':2,'color':8421504}"
			End If
		End If
		If CurrentPlayer = 3 Then
			PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(CurrentPlayer),0),1,"{'mt':2,'color':16777215, 'size': 2.4 }"
			PuPlayer.LabelSet pBackglass,"Play3","Player 3",1,"{'mt':2,'color':16777215 }"
			If PlayersPlayingGame = 3 Then
				PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play1","Player 1",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play2","Player 2",1,"{'mt':2,'color':8421504}"
			End If
			If PlayersPlayingGame = 4 Then
				PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play1","Player 1",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play2","Player 2",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play4score","" & FormatNumber(Score(4),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play4","Player 4",1,"{'mt':2,'color':8421504}"
			End If
		End If
		If CurrentPlayer = 4 Then
			PuPlayer.LabelSet pBackglass,"Play4score","" & FormatNumber(Score(CurrentPlayer),0),1,"{'mt':2,'color':16777215, 'size': 2.4 }"
			PuPlayer.LabelSet pBackglass,"Play4","Player 4",1,"{'mt':2,'color':16777215 }"
			If PlayersPlayingGame = 4 Then
				PuPlayer.LabelSet pBackglass,"Play1score","" & FormatNumber(Score(1),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play1","Player 1",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play2score","" & FormatNumber(Score(2),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play2","Player 2",1,"{'mt':2,'color':8421504}"
				PuPlayer.LabelSet pBackglass,"Play3score","" & FormatNumber(Score(3),0),1,"{'mt':2,'color':8421504, 'size': 2.4 }"
				PuPlayer.LabelSet pBackglass,"Play3","Player 3",1,"{'mt':2,'color':8421504}"
			End If
		End If
	PuPlayer.LabelSet pBackglass,"Ball","Ball "  &  bpgcurrent - BallsRemaining(CurrentPlayer) + 1 ,1,""
	end Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  High Scores
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  

	' load em up


	Dim hschecker:hschecker = 0

	Sub Loadhs
		Dim x
		x = LoadValue(TableName, "HighScore1")
		If(x <> "") Then HighScore(0) = CDbl(x) Else HighScore(0) = 13000000000 End If

		x = LoadValue(TableName, "HighScore1Name")
		If(x <> "") Then HighScoreName(0) = x Else HighScoreName(0) = "SBW" End If

		x = LoadValue(TableName, "HighScore2")
		If(x <> "") then HighScore(1) = CDbl(x) Else HighScore(1) = 12000000000 End If

		x = LoadValue(TableName, "HighScore2Name")
		If(x <> "") then HighScoreName(1) = x Else HighScoreName(1) = "JPS" End If

		x = LoadValue(TableName, "HighScore3")
		If(x <> "") then HighScore(2) = CDbl(x) Else HighScore(2) = 11000000000 End If

		x = LoadValue(TableName, "HighScore3Name")
		If(x <> "") then HighScoreName(2) = x Else HighScoreName(2) = "011" End If

		x = LoadValue(TableName, "HighScore4")
		If(x <> "") then HighScore(3) = CDbl(x) Else HighScore(3) = 10000000000 End If

		x = LoadValue(TableName, "Credits")
		If(x <> "") then Credits = CInt(x) Else Credits = 0 End If

		x = LoadValue(TableName, "TotalGamesPlayed")
		If(x <> "") then TotalGamesPlayed = CInt(x) Else TotalGamesPlayed = 0 End If

		If hschecker = 0 Then
		checkorder
		End If
	End Sub

	Dim hs3,hs2,hs1,hs0,hsn3,hsn2,hsn1,hsn0


	Sub checkorder
		hschecker = 1
		hs3 = HighScore(3)
		hs2 = HighScore(2)
		hs1 = HighScore(1)
		hs0 = HighScore(0)
		hsn3 = HighScoreName(3)
		hsn2 = HighScoreName(2)
		hsn1 = HighScoreName(1)
		hsn0 = HighScoreName(0)
		If hs3 > hs0 Then
			HighScore(0) = hs3
			HighScoreName(0) = hsn3	
			HighScore(1) = hs0
			HighScoreName(1) = hsn0	
			HighScore(2) = hs1
			HighScoreName(2) = hsn1	
			HighScore(3) = hs2
			HighScoreName(3) = hsn2

		ElseIf hs3 > hs1 Then
			HighScore(0) = hs0
			HighScoreName(0) = hsn0	
			HighScore(1) = hs3
			HighScoreName(1) = hsn3	
			HighScore(2) = hs1
			HighScoreName(2) = hsn1	
			HighScore(3) = hs2
			HighScoreName(3) = hsn2
		ElseIf hs3 > hs2 Then
			HighScore(0) = hs0
			HighScoreName(0) = hsn0	
			HighScore(1) = hs1
			HighScoreName(1) = hsn1	
			HighScore(2) = hs3
			HighScoreName(2) = hsn3	
			HighScore(3) = hs2
			HighScoreName(3) = hsn2
		ElseIf hs3 < hs2 Then
			HighScore(0) = hs0
			HighScoreName(0) = hsn0	
			HighScore(1) = hs1
			HighScoreName(1) = hsn1	
			HighScore(2) = hs2
			HighScoreName(2) = hsn2	
			HighScore(3) = hs3
			HighScoreName(3) = hsn3
		End If

		savehs
	End Sub


	Sub Savehs
		SaveValue TableName, "HighScore1", HighScore(0)
		SaveValue TableName, "HighScore1Name", HighScoreName(0)
		SaveValue TableName, "HighScore2", HighScore(1)
		SaveValue TableName, "HighScore2Name", HighScoreName(1)
		SaveValue TableName, "HighScore3", HighScore(2)
		SaveValue TableName, "HighScore3Name", HighScoreName(2)
		SaveValue TableName, "HighScore4", HighScore(3)
		SaveValue TableName, "HighScore4Name", HighScoreName(3)
	End Sub


	Sub Savegp
		SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
		vpmtimer.addtimer 1000, "Loadhs'"
	End Sub


	' Initials

	Dim hsbModeActive:hsbModeActive = False
	Dim hsEnteredName
	Dim hsEnteredDigits(3)
	Dim hsCurrentDigit
	Dim hsValidLetters
	Dim hsCurrentLetter
	Dim hsLetterFlash

	' Check the scores to see if you got one

	Sub CheckHighscore()
		Dim tmp
		tmp = Score(CurrentPlayer)
osbtempscore = Score(CurrentPlayer)
			if osbkey="" Then
			Else
			SubmitOSBScore
			end If

		If tmp > HighScore(3) Then
			AwardSpecial
			vpmtimer.addtimer 2000, "PlaySound ""vo_contratulationsgreatscore"" '"
			HighScore(3) = tmp
			'enter player's name
			HighScoreEntryInit()
			DOF 403, DOFPulse   'DOF MX - Hi Score
		Else
			EndOfBallComplete
		End If
	End Sub





	Sub HighScoreEntryInit()
		hsbModeActive = True
		PlaySound "vo_enteryourinitials"

		hsEnteredDigits(1) = "A"
		hsEnteredDigits(2) = " "
		hsEnteredDigits(3) = " "

		hsCurrentDigit = 1

		pNote "YOU GOT","A HIGH SCORE!"
		PuPlayer.playlistplayex pCallouts,"audiocallouts","yougotahighscore.wav",100,1
		chilloutthemusic
		PuPlayer.playlistplayex pBackglass,"videohighscore","hs1.mov",100,7
		PuPlayer.SetLoop 2,1
		Playsound "bellhs"
		HighScoreDisplayName()
		HighScorelabels	
	End Sub

	' flipper moving around the letters

	Sub EnterHighScoreKey(keycode)
		If keycode = LeftFlipperKey Then
			Playsound "fx_Previous"
				If hsletter = 0 Then
					hsletter = 26
				Else
					hsLetter = hsLetter - 1
				End If
				HighScoreDisplayName()
		End If

		If keycode = RightFlipperKey Then
			Playsound "fx_Next"
				If hsletter = 26 Then
					hsletter = 0
				Else
					hsLetter = hsLetter + 1
				End If
				HighScoreDisplayName()
		End If

		If keycode = StartGameKey or keycode = PlungerKey Then
			PlaySound "success"
				If hsCurrentDigit = 3 Then
					If hsletter = 0 Then
						hsCurrentDigit = hsCurrentDigit -1
					Else
						assignletter
						vpmtimer.addtimer 700, "HighScoreCommitName()'"
						PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,7
					End If
				End If
				If hsCurrentDigit < 3 Then
					If hsletter = 0 Then
						If hsCurrentDigit = 1 Then
						Else
							hsCurrentDigit = hsCurrentDigit -1
						End If
					Else
						assignletter
						hsCurrentDigit = hsCurrentDigit + 1
						HighScoreDisplayName()

					End If
				End If
		End if
	End Sub

	Dim hsletter
	hsletter = 1

	dim hsdigit:hsdigit = 1

	Sub assignletter
		if hscurrentdigit = 1 Then
			hsdigit = 1
		End If
		if hscurrentdigit = 2 Then
			hsdigit = 2
		End If
		if hscurrentdigit = 3 Then
			hsdigit = 3
		End If
		If hsletter = 1 Then 
			hsEnteredDigits(hsdigit) = "A"
		End If
		If hsletter = 2 Then 
			hsEnteredDigits(hsdigit) = "B"
		End If
		If hsletter = 3 Then 
			hsEnteredDigits(hsdigit) = "C"
		End If
		If hsletter = 4 Then 
			hsEnteredDigits(hsdigit) = "D"
		End If
		If hsletter = 5 Then 
			hsEnteredDigits(hsdigit) = "E"
		End If
		If hsletter = 6 Then 
			hsEnteredDigits(hsdigit) = "F"
		End If
		If hsletter = 7 Then 
			hsEnteredDigits(hsdigit) = "G"
		End If
		If hsletter = 8 Then 
			hsEnteredDigits(hsdigit) = "H"
		End If
		If hsletter = 9 Then 
			hsEnteredDigits(hsdigit) = "I"
		End If
		If hsletter = 10 Then 
			hsEnteredDigits(hsdigit) = "J"
		End If
		If hsletter = 11 Then 
			hsEnteredDigits(hsdigit) = "K"
		End If
		If hsletter = 12 Then 
			hsEnteredDigits(hsdigit) = "L"
		End If
		If hsletter = 13 Then 
			hsEnteredDigits(hsdigit) = "M"
		End If
		If hsletter = 14 Then 
			hsEnteredDigits(hsdigit) = "N"
		End If
		If hsletter = 15 Then 
			hsEnteredDigits(hsdigit) = "O"
		End If
		If hsletter = 16 Then 
			hsEnteredDigits(hsdigit) = "P"
		End If
		If hsletter = 17 Then 
			hsEnteredDigits(hsdigit) = "Q"
		End If
		If hsletter = 18 Then 
			hsEnteredDigits(hsdigit) = "R"
		End If
		If hsletter = 19 Then 
			hsEnteredDigits(hsdigit) = "S"
		End If
		If hsletter = 20 Then 
			hsEnteredDigits(hsdigit) = "T"
		End If
		If hsletter = 21 Then 
			hsEnteredDigits(hsdigit) = "U"
		End If
		If hsletter = 22 Then 
			hsEnteredDigits(hsdigit) = "V"
		End If
		If hsletter = 23 Then 
			hsEnteredDigits(hsdigit) = "W"
		End If
		If hsletter = 24 Then 
			hsEnteredDigits(hsdigit) = "X"
		End If
		If hsletter = 25 Then 
			hsEnteredDigits(hsdigit) = "Y"
		End If
		If hsletter = 26 Then 
			hsEnteredDigits(hsdigit) = "Z"
		End If

	End Sub

	Sub HighScorelabels
		PuPlayer.LabelSet pBackglass,"HighScore","YOU GOT A\rHIGH SCORE!",1,""
		PuPlayer.LabelSet pBackglass,"HighScoreL1","A",1,""
		PuPlayer.LabelSet pBackglass,"HighScoreL2"," ",1,""
		PuPlayer.LabelSet pBackglass,"HighScoreL3"," ",1,""
		PuPlayer.LabelSet pBackglass,"HighScoreL4",Score(CurrentPlayer),1,""
		hsletter = 1
	End Sub

	Sub HighScoreDisplayName()

		Select case hsLetter
		Case 0
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","<",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","<",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","<",1,""
		Case 1
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","A",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","A",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","A",1,""
		Case 2
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","B",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","B",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","B",1,""
		Case 3
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","C",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","C",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","C",1,""
		Case 4
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","D",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","D",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","D",1,""
		Case 5
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","E",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","E",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","E",1,""
		Case 6
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","F",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","F",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","F",1,""
		Case 7
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","G",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","G",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","G",1,""
		Case 8
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","H",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","H",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","H",1,""
		Case 9
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","I",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","I",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","I",1,""
		Case 10
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","J",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","J",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","J",1,""
		Case 11
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","K",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","K",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","K",1,""
		Case 12
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","L",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","L",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","L",1,""
		Case 13
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","M",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","M",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","M",1,""
		Case 14
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","N",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","N",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","N",1,""
		Case 15
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","O",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","O",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","O",1,""
		Case 16
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","P",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","P",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","P",1,""
		Case 17
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","Q",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","Q",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","Q",1,""
		Case 18
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","R",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","R",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","R",1,""
		Case 19
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","S",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","S",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","S",1,""
		Case 20
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","T",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","T",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","T",1,""
		Case 21
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","U",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","U",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","U",1,""
		Case 22
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","V",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","V",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","V",1,""
		Case 23
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","W",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","W",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","W",1,""
		Case 24
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","X",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","X",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","X",1,""
		Case 25
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","Y",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","Y",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","Y",1,""
		Case 26
			if(hsCurrentDigit = 1) then PuPlayer.LabelSet pBackglass,"HighScoreL1","Z",1,""
			if(hsCurrentDigit = 2) then PuPlayer.LabelSet pBackglass,"HighScoreL2","Z",1,""
			if(hsCurrentDigit = 3) then PuPlayer.LabelSet pBackglass,"HighScoreL3","Z",1,""
		End Select
	End Sub

	' post the high score letters


	Sub HighScoreCommitName()
		PuPlayer.playlistplayex pBackglass,"videobarb","clear.mov",100,7
		hsEnteredName = hsEnteredDigits(1) & hsEnteredDigits(2) & hsEnteredDigits(3)
		HighScoreName(3) = hsEnteredName
		checkorder
		EndOfBallComplete()
		PuPlayer.LabelSet pBackglass,"HighScore","",1,""
		PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
		PuPlayer.LabelSet pBackglass,"HighScoreL2"," ",1,""
		PuPlayer.LabelSet pBackglass,"HighScoreL3"," ",1,""
	End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  ATTRACT MODE
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 


	Sub StartAttractMode()
		DOF 323, DOFOn   'DOF MX - Attract Mode ON
		bAttractMode = True
		UltraDMDTimer.Enabled = 1
'		StartLightSeq
		'ShowTableInfo
		DMDintroloop
		StartRainbow alights
		StartRainbow2 GI
		DMDattract.Enabled = 1
		intromover.enabled = true
		ruleshelperoff
	End Sub

	Sub StopAttractMode()
		DOF 323, DOFOff   'DOF MX - Attract Mode Off
		bAttractMode = False
		DMDScoreNow
		LightSeqAttract.StopPlay
		LightSeqAttract2.StopPlay
		StopRainbow alights
		StopRainbow2 GI
		ResetAllLightsColor
		DMDattract.Enabled = 0
		intromover.enabled = false
		
	'StopSong
	End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  CINEMATIC SKIPPING
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  


	Dim ldown:ldown = 0
	Dim rdown:rdown = 0

	Sub checkdown
		If ldown + rdown = 2 Then
			skipscene
		End If
	End Sub

	Sub skipscene

	End Sub






'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  LIGHTING / RAINBOW LIGHTS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 


	'********************************************************************************************
	' Only for VPX 10.2 and higher.
	' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
	' When TotalPeriod done, light or flasher will be set to FinalState value where
	' Final State values are:   0=Off, 1=On, 2=Return to previous State
	'********************************************************************************************

	Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version

		If TypeName(MyLight) = "Light" Then

			If FinalState = 2 Then
				FinalState = MyLight.State 'Keep the current light state
			End If
			MyLight.BlinkInterval = BlinkPeriod
			MyLight.Duration 2, TotalPeriod, FinalState
		ElseIf TypeName(MyLight) = "Flasher" Then

			Dim steps

			' Store all blink information
			steps = Int(TotalPeriod / BlinkPeriod + .5) 'Number of ON/OFF steps to perform
			If FinalState = 2 Then                      'Keep the current flasher state
				FinalState = ABS(MyLight.Visible)
			End If
			MyLight.UserValue = steps * 10 + FinalState 'Store # of blinks, and final state

			' Start blink timer and create timer subroutine
			MyLight.TimerInterval = BlinkPeriod
			MyLight.TimerEnabled = 0
			MyLight.TimerEnabled = 1
			ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
		End If
	End Sub

	'******************************************
	' Change light color - simulate color leds
	' changes the light color and state
	' 10 colors: red, orange, amber, yellow...
	'******************************************
	' in this table this colors are use to keep track of the progress during the acts and battles

	'colors
	Dim red, orange, amber, yellow, darkgreen, green, blue, darkblue, purple, white, base

	red = 10
	orange = 9
	amber = 8
	yellow = 7
	darkgreen = 6
	green = 5
	blue = 4
	darkblue = 3
	purple = 2
	white = 1
	base = 11

	Sub SetLightColor(n, col, stat)
		Select Case col
			Case red
				n.color = RGB(18, 0, 0)
				n.colorfull = RGB(255, 0, 0)
			Case orange
				n.color = RGB(18, 3, 0)
				n.colorfull = RGB(255, 64, 0)
			Case amber
				n.color = RGB(193, 49, 0)
				n.colorfull = RGB(255, 153, 0)
			Case yellow
				n.color = RGB(18, 18, 0)
				n.colorfull = RGB(255, 255, 0)
			Case darkgreen
				n.color = RGB(0, 8, 0)
				n.colorfull = RGB(0, 64, 0)
			Case green
				n.color = RGB(0, 18, 0)
				n.colorfull = RGB(0, 255, 0)
			Case blue
				n.color = RGB(0, 18, 18)
				n.colorfull = RGB(0, 255, 255)
			Case darkblue
				n.color = RGB(0, 8, 8)
				n.colorfull = RGB(0, 64, 64)
			Case purple
				n.color = RGB(128, 0, 128)
				n.colorfull = RGB(255, 0, 255)
			Case white
				n.color = RGB(255, 252, 224)
				n.colorfull = RGB(193, 91, 0)
			Case white
				n.color = RGB(255, 252, 224)
				n.colorfull = RGB(193, 91, 0)
			Case base
				n.color = RGB(255, 197, 143)
				n.colorfull = RGB(255, 255, 236)
		End Select
		If stat <> -1 Then
			n.State = 0
			n.State = stat
		End If
	End Sub

	Sub ResetAllLightsColor ' Called at a new game
		SetLightColor l13, red, -1		
	End Sub

	Sub UpdateBonusColors
	End Sub

	'*************************
	' Rainbow Changing Lights
	'*************************

	Dim RGBStep, RGBFactor, rRed, rGreen, rBlue, RainbowLights

	Sub StartRainbow(n)
		set RainbowLights = n
		RGBStep = 0
		RGBFactor = 5
		rRed = 255
		rGreen = 0
		rBlue = 0
		RainbowTimer.Enabled = 1
	End Sub

	Dim RGBStep2, RGBFactor2, rRed2, rGreen2, rBlue2, RainbowLights2
	Sub StartRainbow2(n)
		set RainbowLights2 = n
		RGBStep2 = 0
		RGBFactor2 = 5
		rRed2 = 255
		rGreen2 = 0
		rBlue2 = 0
		RainbowTimer1.Enabled = 1
	End Sub

	Sub StopRainbow(n)
		Dim obj
		RainbowTimer.Enabled = 0
		RainbowTimer.Enabled = 0
			For each obj in RainbowLights
				SetLightColor obj, "white", 0
			Next
	End Sub

	Sub StopRainbow2(n)
		Dim obj
		RainbowTimer1.Enabled = 0
			For each obj in RainbowLights2
				SetLightColor obj, "white", 0
				obj.state = 1
				obj.Intensity = 15
			Next
	End Sub

	Sub RainbowTimer_Timer 'rainbow led light color changing
		Dim obj
		Select Case RGBStep
			Case 0 'Green
				rGreen = rGreen + RGBFactor
				If rGreen > 255 then
					rGreen = 255
					RGBStep = 1
				End If
			Case 1 'Red
				rRed = rRed - RGBFactor
				If rRed < 0 then
					rRed = 0
					RGBStep = 2
				End If
			Case 2 'Blue
				rBlue = rBlue + RGBFactor
				If rBlue > 255 then
					rBlue = 255
					RGBStep = 3
				End If
			Case 3 'Green
				rGreen = rGreen - RGBFactor
				If rGreen < 0 then
					rGreen = 0
					RGBStep = 4
				End If
			Case 4 'Red
				rRed = rRed + RGBFactor
				If rRed > 255 then
					rRed = 255
					RGBStep = 5
				End If
			Case 5 'Blue
				rBlue = rBlue - RGBFactor
				If rBlue < 0 then
					rBlue = 0
					RGBStep = 0
				End If
		End Select
			For each obj in RainbowLights
				obj.color = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
				obj.colorfull = RGB(rRed, rGreen, rBlue)
			Next
	End Sub

	Sub RainbowTimer1_Timer 'rainbow led light color changing
		Dim obj
		Select Case RGBStep2
			Case 0 'Green
				rGreen2 = rGreen2 + RGBFactor2
				If rGreen2 > 255 then
					rGreen2 = 255
					RGBStep2 = 1
				End If
			Case 1 'Red
				rRed2 = rRed2 - RGBFactor2
				If rRed2 < 0 then
					rRed2 = 0
					RGBStep2 = 2
				End If
			Case 2 'Blue
				rBlue2 = rBlue2 + RGBFactor2
				If rBlue2 > 255 then
					rBlue2 = 255
					RGBStep2 = 3
				End If
			Case 3 'Green
				rGreen2 = rGreen2 - RGBFactor2
				If rGreen2 < 0 then
					rGreen2 = 0
					RGBStep2 = 4
				End If
			Case 4 'Red
				rRed2 = rRed2 + RGBFactor2
				If rRed2 > 255 then
					rRed2 = 255
					RGBStep2 = 5
				End If
			Case 5 'Blue
				rBlue2 = rBlue2 - RGBFactor2
				If rBlue2 < 0 then
					rBlue2 = 0
					RGBStep2 = 0
				End If
		End Select
			For each obj in RainbowLights2
				obj.color = RGB(rRed2 \ 10, rGreen2 \ 10, rBlue2 \ 10)
				obj.colorfull = RGB(rRed2, rGreen2, rBlue2)
			Next
	End Sub


'	Sub StartLightSeq()
'		Dim a
'		For each a in alights
'			a.State = 1
'		Next
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 50, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqCircleOutOn, 15, 2
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 25, 1
'		LightSeqAttract.UpdateInterval = 10
'		LightSeqAttract.Play SeqCircleOutOn, 15, 3
'		LightSeqAttract.UpdateInterval = 5
'		LightSeqAttract.Play SeqRightOn, 50, 1
'		LightSeqAttract.UpdateInterval = 5
'		LightSeqAttract.Play SeqLeftOn, 50, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqRightOn, 50, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 50, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqRightOn, 40, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 40, 1
'		LightSeqAttract.UpdateInterval = 10
'		LightSeqAttract.Play SeqRightOn, 30, 1
'		LightSeqAttract.UpdateInterval = 10
'		LightSeqAttract.Play SeqLeftOn, 30, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqRightOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqRightOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 15, 1
'		LightSeqAttract.UpdateInterval = 10
'		LightSeqAttract.Play SeqCircleOutOn, 15, 3
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqRightOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 25, 1
'		LightSeqAttract.UpdateInterval = 5
'		LightSeqAttract.Play SeqStripe1VertOn, 50, 2
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqCircleOutOn, 15, 2
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqStripe1VertOn, 50, 3
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqRightOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqCircleOutOn, 15, 2
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqStripe2VertOn, 50, 3
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqRightOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 25, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqStripe1VertOn, 25, 3
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqStripe2VertOn, 25, 3
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqUpOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqDownOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqRightOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqRightOn, 15, 1
'		LightSeqAttract.UpdateInterval = 8
'		LightSeqAttract.Play SeqLeftOn, 15, 1
'		For each a in GI
'			a.State = 1
'			a.Intensity = 80
'		Next
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 50, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqCircleOutOn, 15, 2
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 10
'		LightSeqAttract2.Play SeqCircleOutOn, 15, 3
'		LightSeqAttract2.UpdateInterval = 5
'		LightSeqAttract2.Play SeqRightOn, 50, 1
'		LightSeqAttract2.UpdateInterval = 5
'		LightSeqAttract2.Play SeqLeftOn, 50, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqRightOn, 50, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 50, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqRightOn, 40, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 40, 1
'		LightSeqAttract2.UpdateInterval = 10
'		LightSeqAttract2.Play SeqRightOn, 30, 1
'		LightSeqAttract2.UpdateInterval = 10
'		LightSeqAttract2.Play SeqLeftOn, 30, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqRightOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqRightOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 10
'		LightSeqAttract2.Play SeqCircleOutOn, 15, 3
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqRightOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 5
'		LightSeqAttract2.Play SeqStripe1VertOn, 50, 2
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqCircleOutOn, 15, 2
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqStripe1VertOn, 50, 3
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqRightOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqCircleOutOn, 15, 2
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqStripe2VertOn, 50, 3
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqRightOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 25, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqStripe1VertOn, 25, 3
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqStripe2VertOn, 25, 3
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqUpOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqDownOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqRightOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqRightOn, 15, 1
'		LightSeqAttract2.UpdateInterval = 8
'		LightSeqAttract2.Play SeqLeftOn, 15, 1
'	End Sub

'	Sub LightSeqAttract_PlayDone()
'		StartLightSeq()
'	End Sub

'	Sub LightSeqTilt_PlayDone()
'		LightSeqTilt.Play SeqAllOff
'	End Sub

'	Sub LightSeqSkillshot_PlayDone()
'		LightSeqSkillshot.Play SeqAllOff
'	End Sub



	'**********************
	'     GI effects
	' independent routine
	' it turns on the gi
	' when there is a ball
	' in play
	'**********************

	Dim OldGiState
	OldGiState = -1   'start witht the Gi off

	Sub ChangeGi(col) 'changes the gi color
		Dim bulb
		For each bulb in GI
			SetLightColor bulb, col, -1
		Next
	End Sub

	Sub GIUpdateTimer_Timer
		Dim tmp, obj
		tmp = Getballs
		If UBound(tmp) <> OldGiState Then
			OldGiState = Ubound(tmp)
			If UBound(tmp) = 3 Then 'we have 4 captive balls on the table (-1 means no balls, 0 is the first ball, 1 is the second..)
				'GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
			Else
				'Gion
			End If
		End If
	End Sub

	Sub GiOn
		DOF 126, DOFOn
		Dim bulb
		For each bulb in GI
			SetLightColor bulb, base, -1
			bulb.State = 1
		Next
	End Sub

	Sub GiOff
		DOF 126, DOFOff
		Dim bulb
		For each bulb in GI
			bulb.State = 0
		Next
	End Sub


	' GI & light sequence effects


	Sub GiEffect(n)
		Select Case n
			Case 0 'all off
				'LightSeqGi.Play SeqAlloff
			Case 1 'all blink
				'LightSeqGi.UpdateInterval = 4
				'LightSeqGi.Play SeqBlinking, , 5, 100
			Case 2 'random
				'LightSeqGi.UpdateInterval = 10
				'LightSeqGi.Play SeqRandom, 5, , 1000
			Case 3 'upon
				'LightSeqGi.UpdateInterval = 4
				'LightSeqGi.Play SeqUpOn, 5, 1
			Case 4 ' left-right-left
				'LightSeqGi.UpdateInterval = 5
				'LightSeqGi.Play SeqLeftOn, 10, 1
				'LightSeqGi.UpdateInterval = 5
				'LightSeqGi.Play SeqRightOn, 10, 1
		End Select
	End Sub

	Sub LightEffect(n)
		Select Case n
			Case 0 ' all off
				'LightSeqInserts.Play SeqAlloff
			Case 1 'all blink
				'LightSeqInserts.UpdateInterval = 4
				'LightSeqInserts.Play SeqBlinking, , 5, 100
			Case 2 'random
				'LightSeqInserts.UpdateInterval = 10
				'LightSeqInserts.Play SeqRandom, 5, , 1000
			Case 3 'upon
				'LightSeqInserts.UpdateInterval = 4
				'LightSeqInserts.Play SeqUpOn, 10, 1
			Case 4 ' left-right-left
				'LightSeqInserts.UpdateInterval = 5
				'LightSeqInserts.Play SeqLeftOn, 10, 1
				'LightSeqInserts.UpdateInterval = 5
				'LightSeqInserts.Play SeqRightOn, 10, 1
			Case 5 'random
				'LightSeqbumper.UpdateInterval = 4
				'LightSeqbumper.Play SeqBlinking, , 5, 10
			Case 6 'random
				'LightSeqRSling.UpdateInterval = 4
				'LightSeqRSling.Play SeqBlinking, , 5, 6
			Case 7 'random
				'LightSeqLSling.UpdateInterval = 4
				'LightSeqLSling.Play SeqBlinking, , 5, 6
			Case 8 'random
				'LightSeqBack.UpdateInterval = 4
				'LightSeqBack.Play SeqBlinking, , 5, 6
			Case 12 'random
				'LightSeqlr.UpdateInterval = 4
				'LightSeqlr.Play SeqBlinking, , 5, 10
		End Select
	End Sub

	' Flasher Effects using lights

	Dim FEStep, FEffect
	FEStep = 0
	FEffect = 0

	Sub FlashEffect(n)
		Select case n
			Case 0 ' all off
				LightSeqFlasher.Play SeqAlloff
			Case 1 'all blink
				LightSeqFlasher.UpdateInterval = 4
				LightSeqFlasher.Play SeqBlinking, , 5, 100
			Case 2 'random
				LightSeqFlasher.UpdateInterval = 10
				LightSeqFlasher.Play SeqRandom, 5, , 1000
			Case 3 'upon
				LightSeqFlasher.UpdateInterval = 4
				LightSeqFlasher.Play SeqUpOn, 10, 1
			Case 4 ' left-right-left
				LightSeqFlasher.UpdateInterval = 5
				LightSeqFlasher.Play SeqLeftOn, 10, 1
				LightSeqFlasher.UpdateInterval = 5
				LightSeqFlasher.Play SeqRightOn, 10, 1
			Case 5 ' top flashers blink fast
		End Select
	End Sub







	'****************************
	' Flashers - Thanks Flupper
	'****************************

	Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4, FlashLevel5, FlashLevel6
'	Flasherlight4.IntensityScale = 0

	'*** right red flasher ***
	Sub Flasherflash3_Timer()
		dim flashx3, matdim
		If not Flasherflash3.TimerEnabled Then 
			Flasherflash3.TimerEnabled = True
			Flasherflash3.visible = 1
			Flasherlit3.visible = 1
		End If
		flashx3 = FlashLevel3 * FlashLevel3 * FlashLevel3
		Flasherflash3.opacity = 1500 * flashx3
		Flasherlit3.BlendDisableLighting = 10 * flashx3
		Flasherbase3.BlendDisableLighting =  flashx3
		Flasherlight4.IntensityScale = flashx3
		matdim = Round(10 * FlashLevel3)
		Flasherlit3.material = "domelit" & matdim
		FlashLevel3 = FlashLevel3 * 0.9 - 0.01
		If FlashLevel3 < 0.15 Then
			Flasherlit3.visible = 0
		Else
			Flasherlit3.visible = 1
		end If
		If FlashLevel3 < 0 Then
			Flasherflash3.TimerEnabled = False
			Flasherflash3.visible = 0
		End If
	End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   SCORING FUNCTIONS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  

	Sub AddScore(points)
		If(Tilted = False) Then
			' add the points to the current players score variable
			Score(CurrentPlayer) = Score(CurrentPlayer) + points
			DMDScore
			pUpdateScores
		End if
	End Sub

	Sub AwardExtraBall()
		If NOT bExtraBallWonThisBall Then
			LightShootAgain.State = 1
			flashflash.Enabled = True
			LightSeqFlasher.UpdateInterval = 150
			LightSeqFlasher.Play SeqRandom, 10, , 10000
			ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
			bExtraBallWonThisBall = True
			DOF 402, DOFPulse   'DOF MX - Extra Ball
			'extraballready(CurrentPlayer) = 0
			If bMultiBallMode = false Then
			PuPlayer.playlistplayex pBackglass,"videoextraball","extraballsm.mov",100,1
				PuPlayer.playlistplayex pCallouts,"audiocallouts","extraball.wav",100,1
		chilloutthemusic
			End If
		Else
		AddScore 4000000
		END If
	End Sub

	Sub AwardSpecial()
		Credits = Credits + 1
		DOF 140, DOFOn
		PlaySound SoundFXDOF("knocker",136,DOFPulse,DOFKnocker)		
		DOF 115, DOFPulse
		GiEffect 1
		LightEffect 1
		DOF 400, DOFPulse   'DOF MX - Special
	End Sub


	Sub AwardSkillshot()
		Dim i
		DOF 125, DOFPulse
		ResetSkillShotTimer_Timer
		AddScore SkillshotValue(CurrentPLayer)
		GiEffect 1
		LightEffect 2
		LightEffect 9
		DOF 401, DOFPulse   'DOF MX - Skillshot
		PuPlayer.playlistplayex pCallouts,"audiocallouts","skillshot.wav",100,1
		chilloutthemusic
		PuPlayer.playlistplayex pBackglass,"videoskillshot","",100,3
		pNote "SKILLSHOT",SkillShotValue(CurrentPLayer)
		SkillShotValue(CurrentPLayer) = SkillShotValue(CurrentPLayer) + 1000000
	End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'   BALL FUNCTIONS & DRAINS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  


	Function Balls
		Dim tmp
		tmp = bpgcurrent - BallsRemaining(CurrentPlayer) + 1
		If tmp> bpgcurrent Then
			Balls = bpgcurrent
		Else
			Balls = tmp
		End If
	End Function


	Sub FirstBall
		ResetForNewPlayerBall()
		CreateNewBall()
	End Sub


	Sub ResetForNewPlayerBall()
		PuPlayer.playresume 4
		If PlayersPlayingGame > 1 Then
			If CurrentPlayer = 1 Then
				PuPlayer.playlistplayex pBackglass,"videoplayers","player1sm1.mov",40,1
						PuPlayer.playlistplayex pCallouts,"audiocallouts","player1.wav",70,1
		chilloutthemusic
			Elseif currentplayer = 2 Then
				PuPlayer.playlistplayex pBackglass,"videoplayers","player2sm.mov",40,1
						PuPlayer.playlistplayex pCallouts,"audiocallouts","player2.wav",70,1
		chilloutthemusic
			Elseif currentplayer = 3 Then
				PuPlayer.playlistplayex pBackglass,"videoplayers","player3sm.mov",40,1
						PuPlayer.playlistplayex pCallouts,"audiocallouts","player3.wav",70,1
		chilloutthemusic
			Elseif currentplayer = 4 Then
				PuPlayer.playlistplayex pBackglass,"videoplayers","player4sm.mov",40,1
						PuPlayer.playlistplayex pCallouts,"audiocallouts","player4.wav",70,1
		chilloutthemusic
			End If
		Else
			pNote "BALL " & Balls,"LAUNCH BALL"
		PlaySound "batman"
		End If
		AddScore 0
		BonusPoints(CurrentPlayer) = 0
		bBonusHeld = False
		bExtraBallWonThisBall = False
		ResetNewBallLights()
		ResetNewBallVariables
		bBallSaverReady = True
		bSkillShotReady = True
	End Sub


	Sub CreateNewBall()
		BallRelease.CreateSizedball BallSize / 2
		BallsOnPlayfield = BallsOnPlayfield + 1
		PlaySoundAt SoundFXDOF("fx_Ballrel", 114, DOFPulse, DOFContactors), BallRelease
		BallRelease.Kick 90, 4
		If BallsOnPlayfield > 1 Then
			bMultiBallMode = True
				DOF 127, DOFOn    'Beacon ON
			bAutoPlunger = True
		End If

	End Sub


	Sub AddMultiball(nballs)
		mBalls2Eject = mBalls2Eject + nballs
		CreateMultiballTimer.Enabled = True
	End Sub

	Sub CreateMultiballTimer_Timer()
		If bBallInPlungerLane Then
			Exit Sub
		Else
			If BallsOnPlayfield <MaxMultiballs Then
				CreateNewBall()
				mBalls2Eject = mBalls2Eject -1
				If mBalls2Eject = 0 Then 
					Me.Enabled = False
				End If
			Else 
				mBalls2Eject = 0
				Me.Enabled = False
			End If
		End If
	End Sub

	Sub EndOfBall()
		PuPlayer.playlistplayex pAudio,"audiomultiballs","clear.mp3",100,1


		bMultiBallMode = False
		bOnTheFirstBall = False
		If NOT Tilted Then
			vpmtimer.addtimer 500, "EndOfBall2 '"
		Else 
			vpmtimer.addtimer 500, "EndOfBall2 '"
		End If
	End Sub

	
	Sub EndOfBall2()
		Tilted = False
		Tilt = 0
		DisableTable False
		tilttableclear.enabled = False
		tilttime = 0
		If(ExtraBallsAwards(CurrentPlayer) <> 0) Then
			ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1
			If(ExtraBallsAwards(CurrentPlayer) = 0) Then
				LightShootAgain.State = 0
			End If
			LightSeqFlasher.UpdateInterval = 150
			LightSeqFlasher.Play SeqRandom, 10, , 2000

			CreateNewBall()
			ResetForNewPlayerBall
			PuPlayer.playlistplayex pBackglass,"videoextraball","isthatsupposedtoimpress.mov",100,2
			pNote "SHOOT AGAIN","SAME PLAYER"
				PuPlayer.playlistplayex pCallouts,"audiocallouts","shootagain.wav",100,1
		chilloutthemusic
		Else
			BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1
			If(BallsRemaining(CurrentPlayer) <= 0) Then
				EndOfBallComplete()
			Else
				EndOfBallComplete()
			End If
		End If
	End Sub

	Sub EndOfBallComplete()
		ResetNewBallVariables
		Dim NextPlayer
		If(PlayersPlayingGame> 1) Then
			NextPlayer = CurrentPlayer + 1
			If(NextPlayer> PlayersPlayingGame) Then
				NextPlayer = 1
			End If
		Else
			NextPlayer = CurrentPlayer
		End If
		If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then
			EndOfGame()
		Else
			CurrentPlayer = NextPlayer
			AddScore 0
			ResetForNewPlayerBall()
			CreateNewBall()
			If PlayersPlayingGame> 1 Then
				PlaySound "vo_player" &CurrentPlayer
			End If
		End If
	End Sub

	Sub Balldrained
		DOF 312, DOFPulse
		PuPlayer.playlistplayex pMusic,"audioclear","clear.mp3",100,3
		PuPlayer.playlistplayex pBackglass,"videodrain","",100,3
		PuPlayer.playlistplayex pTopper,"topper","topperballlost.mp4",100,3
		PuPlayer.playlistplayex pDMD,"topper","balllost.mp4",0,3

	End Sub

	Sub Drain_Hit()
		Drain.DestroyBall
		BallsOnPlayfield = BallsOnPlayfield - 1
		PlaySoundAt "fx_drain", Drain
		If Tilted Then
			StopEndOfBallMode
		End If
		If(bGameInPLay = True) AND(Tilted = False) Then
			If(bBallSaverActive = True) Then
				AddMultiball 1
				bAutoPlunger = True
				If bMultiBallMode = False Then
					Ballsaved
					PuPlayer.playlistplayex pCallouts,"audiocallouts","ballsaved.wav",100,1
		chilloutthemusic
					'pNote "BALL SAVED",""
					PuPlayer.playlistplayex pBackglass,"videoballsaved","",100,1
				End If
			Else
				If(BallsOnPlayfield = 1) Then
					If(bMultiBallMode = True) then
						bMultiBallMode = False
						DOF 127, DOFOff   'DOF - Beacon - OFF
						ChangeGi "white"
					End If
					bMultiBallMode = False		
					ChangeGi "white"
			End If
				If(BallsOnPlayfield = 0) Then
					bMultiBallMode = False
					ChangeGi "white"
					vpmtimer.addtimer 2000, "Balldrained '"
					vpmtimer.addtimer 10000, "EndOfBall '"
					StopEndOfBallMode
				End If
			End If
		End If
	End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  BALL SAVE & LAUNCH
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 

	Sub ballsavestarttrigger_hit
		If(bBallSaverReady = True) AND(15 <> 0) And(bBallSaverActive = False) Then
			EnableBallSaver 15
		End If
	End Sub

	Sub swPlungerRest_Hit()
		PlaySoundAt "fx_sensor", ActiveBall
		bBallInPlungerLane = True
		If bAutoPlunger Then
			DOF 114, DOFPulse		
			DOF 115, DOFPulse
			DOF 318, DOFPulse
			bAutoPlunger = False
		End If
		DOF 141, DOFOn
		DOF 317, DOFOn
		If bSkillShotReady Then
			swPlungerRest.TimerEnabled = 1
			UpdateSkillshot()
		End If
		LastSwitchHit = "swPlungerRest"
	End Sub


	Sub swPlungerRest_UnHit()
		bBallInPlungerLane = False
		swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
		If bSkillShotReady Then
			ResetSkillShotTimer.Enabled = 1
		End If
		DOF 317, DOFOff   'DOF MX - Ball is ready to Launch - OFF
	End Sub


	Sub swPlungerRest_Timer
		'DMD "start-5.wmv", "", "",  6000
		swPlungerRest.TimerEnabled = 0
	End Sub

	Sub EnableBallSaver(seconds)
		bBallSaverActive = True
		bBallSaverReady = False
		BallSaverTimerExpired.Interval = 1000 * seconds
		BallSaverTimerExpired.Enabled = True
		BallSaverSpeedUpTimer.Interval = 1000 * seconds -(1000 * seconds) / 3
		BallSaverSpeedUpTimer.Enabled = True
		' if you have a ball saver light you might want to turn it on at this point (or make it flash)
		LightShootAgain.BlinkInterval = 160
		LightShootAgain.State = 2
	End Sub

	' The ball saver timer has expired.  Turn it off AND reset the game flag
	'
	Sub BallSaverTimerExpired_Timer()
		'debug.print "Ballsaver ended"
		BallSaverTimerExpired.Enabled = False
		' clear the flag
		Dim waittime
		waittime = 4000
		vpmtimer.addtimer waittime, "ballsavegrace'"
		' if you have a ball saver light then turn it off at this point
		If bExtraBallWonThisBall = True Then
			LightShootAgain.State = 1
		Else
			LightShootAgain.State = 0
		End If
	End Sub

	Sub ballsavegrace
		bBallSaverActive = False
	End Sub

	Sub BallSaverSpeedUpTimer_Timer()
		'debug.print "Ballsaver Speed Up Light"
		BallSaverSpeedUpTimer.Enabled = False
		' Speed up the blinking
		LightShootAgain.BlinkInterval = 80
		LightShootAgain.State = 2
	End Sub

'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'FRAMEWORK BASE CODE END HERE
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************
'********************************************************************************************************************************************

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  ATTRACT MODE
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 


	Dim introposition
	introposition = 0

	Sub DMDintroloop
		PuPlayer.LabelSet pBackglass,"modetitle","",1,"{'mt':2,'color':16777215, 'size': 0, 'xpos': 80.7, 'xalign': 1, 'ypos': 72.6, 'yalign': 0}"
		introtime = 0	
		introposition = introposition + 1
		Select Case introposition
		Case 1
			PuPlayer.LabelSet pBackglass,"HighScore","",1,""
			PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
			PuPlayer.LabelSet pBackglass,"HighScoreL2","",1,""
			PuPlayer.LabelSet pBackglass,"HighScoreL3","",1,""
			PuPlayer.playlistplayex pMusic,"audiobg","",99,1
			PuPlayer.playlistplayex pBackglass,"videoattract","batman.mp4",100,1
'			PuPlayer.LabelSet pBackglass,"notetitle","Need Rules Help?",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","Hold both flippers for 2 seconds while in attract mode for rules overlay.",1,""
			PuPlayer.LabelSet pBackglass,"high1name","",1,""
			PuPlayer.LabelSet pBackglass,"high1score","",1,""
			PuPlayer.LabelSet pBackglass,"high2name","",1,""
			PuPlayer.LabelSet pBackglass,"high2score","",1,""
			PuPlayer.LabelSet pBackglass,"high3name","",1,""
			PuPlayer.LabelSet pBackglass,"high3score","",1,""
			PuPlayer.LabelSet pBackglass,"high4name","",1,""
			PuPlayer.LabelSet pBackglass,"high4score","",1,""
		Case 2
			loadhs
			PuPlayer.playlistplayex pMusic,"audiobg","",99,1
			PuPlayer.LabelSet pBackglass,"HighScore","",1,""
			PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
			PuPlayer.LabelSet pBackglass,"HighScoreL2","",1,""
			PuPlayer.LabelSet pBackglass,"HighScoreL3","",1,""
			PuPlayer.LabelSet pBackglass,"notetitle","Got A killer good High Score?",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","Post it to #stlevpx on instagram to brag.",1,""
			PuPlayer.playlistplayex pBackglass,"videoattract","batman.mp4",100,1
			PuPlayer.LabelSet pBackglass,"high1name",HighScoreName(0),1,""
			PuPlayer.LabelSet pBackglass,"high1score",FormatNumber(HighScore(0),0),1,""
			PuPlayer.LabelSet pBackglass,"high2name",HighScoreName(1),1,""
			PuPlayer.LabelSet pBackglass,"high2score",FormatNumber(HighScore(1),0),1,""
			PuPlayer.LabelSet pBackglass,"high3name",HighScoreName(2),1,""
			PuPlayer.LabelSet pBackglass,"high3score",FormatNumber(HighScore(2),0),1,""
			PuPlayer.LabelSet pBackglass,"high4name",HighScoreName(3),1,""
			PuPlayer.LabelSet pBackglass,"high4score",FormatNumber(HighScore(3),0),1,""
		Case 3
			PuPlayer.playlistplayex pBackglass,"videoattract","batman.mp4",100,1
			PuPlayer.LabelSet pBackglass,"notetitle","Did You Know?",1,""
			PuPlayer.LabelSet pBackglass,"HighScore","",1,""
			PuPlayer.LabelSet pBackglass,"HighScoreL1","",1,""
			PuPlayer.LabelSet pBackglass,"HighScoreL2","",1,""
			PuPlayer.LabelSet pBackglass,"HighScoreL3","",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","You can change the music to 80s jams in the script? Check it out..",1,""
			PuPlayer.LabelSet pBackglass,"high1name","",1,""
			PuPlayer.LabelSet pBackglass,"high1score","",1,""
			PuPlayer.LabelSet pBackglass,"high2name","",1,""
			PuPlayer.LabelSet pBackglass,"high2score","",1,""
			PuPlayer.LabelSet pBackglass,"high3name","",1,""
			PuPlayer.LabelSet pBackglass,"high3score","",1,""
			PuPlayer.LabelSet pBackglass,"high4name","",1,""
			PuPlayer.LabelSet pBackglass,"high4score","",1,""
		Case 4
			PuPlayer.LabelSet pBackglass,"notetitle","Want to code your own original table?",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","Check Scottywics YouTube Channel for helpful guides.",1,""
			PuPlayer.playlistplayex pBackglass,"videoattract","batman.mp4",100,1

		Case 5
			PuPlayer.playlistplayex pBackglass,"videoattract","batman.mp4",100,1
			PuPlayer.LabelSet pBackglass,"notetitle","Be Excellent to Each Other",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","There is no place for homophobia, fascism, sexism, racism or hate.",1,""
			introposition = 0
	End Select
	End Sub


	Dim introtime
	introtime = 0

	Sub intromover_timer
		introtime = introtime + 1
		If introposition = 1 Then
			If introtime = 47 Then
				DMDintroloop
			End If
		End If
		If introposition = 2 Then
			If introtime = 13 Then
				DMDintroloop
			End If
		End If
		If introposition = 3 Then
			If introtime = 13 Then
				DMDintroloop
			End If
		End If
		If introposition = 4 Then
			If introtime = 9 Then
				DMDintroloop
			End If
		End If
		If introposition = 0 Then
			If introtime = 9 Then
				introposition = 0
				DMDintroloop
			End If
		End If
	End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  TABLE VARIABLES
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 


	' Game Variables
	Dim LaneBonus
	Dim TargetBonus
	Dim RampBonus
	Dim OrbitBonus
	Dim spinvalue
	Dim finalflips
	Dim Saves
	Dim Drains
	Dim inmode


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  GAME STARTING & RESETS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 

	Sub Game_Init() 'called at the start of a new game
		Dim i
		For i = 0 to 4
			SkillshotValue(i) = 1000000 ' increases by 1000000 each time it is collected
		Next
		lrflashtime.Enabled = False
		bExtraBallWonThisBall = False
		PuPlayer.LabelSet pBackglass,"high1name","",1,""
		PuPlayer.LabelSet pBackglass,"high1score","",1,""
		PuPlayer.LabelSet pBackglass,"high2name","",1,""
		PuPlayer.LabelSet pBackglass,"high2score","",1,""
		PuPlayer.LabelSet pBackglass,"high3name","",1,""
		PuPlayer.LabelSet pBackglass,"high3score","",1,""
		PuPlayer.LabelSet pBackglass,"high4name","",1,""
		PuPlayer.LabelSet pBackglass,"high4score","",1,""
		PuPlayer.LabelShowPage pBackglass,1,0,""
		pUpdateScores
		PuPlayer.playlistplayex pBackglass,"scene","base.mov",0,1
		PuPlayer.SetBackground pBackglass,1
		PuPlayer.LabelSet pBackglass,"Play1","PLAYER 1",1,"{'mt':2,'color':16777215}"
		PuPlayer.LabelSet pBackglass,"notetitle","",1,""
		PuPlayer.LabelSet pBackglass,"notecopy","",1,""
		ruleshelperon
		resetbackglass
	End Sub

	Sub StopEndOfBallMode()              'this sub is called after the last ball is drained
		ResetSkillShotTimer_Timer
	End Sub

	Sub ResetNewBallVariables()
		PuPlayer.playlistplayex pMusic,"audiobg","",99,1
		PuPlayer.SetLoop 4,1
	End Sub

	Sub TurnOffPlayfieldLights()
		Dim a
		For each a in alights
			a.State = 0
		Next
	End Sub

	Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released
		TurnOffPlayfieldLights()
		currentplayerbackglass
	End Sub

	Sub startamultiball
		Dim waittime
		waittime = 1000
		vpmtimer.addtimer waittime, "closeupshop'"
	End Sub
	
	Sub closeupshop

	End Sub

	Sub endamultiball

	End Sub


	Sub UpdateSkillShot() 'Updates the skillshot light
		l7.State = 2
	End Sub

	Sub SkillshotOff_Hit 'trigger to stop the skillshot due to a weak plunger shot
		If bSkillShotReady Then

			ResetSkillShotTimer_Timer
		End If
	End Sub

	Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
		ResetSkillShotTimer.Enabled = 0
		bSkillShotReady = False
			If l7.State = 2 then l7.State = 0
	End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  RULES HELPER FOR YOUR BACKGLASS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 


	Sub ruleshelperon
		rulestime.enabled = 1
	End Sub

	Sub ruleshelperoff
		rulestime.enabled = 0
	End Sub

	Dim rulesposition
	rulesposition = 0

	Sub rulestime_timer
		If turnoffrules = 1 then exit sub end if
		rulesposition = rulesposition + 1
		Select Case rulesposition
		Case 1
			PuPlayer.LabelSet pBackglass,"notetitle","Find Barb",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","Hit the B A R B lane targets to light Barb Lock.",1,""
		Case 8
			PuPlayer.LabelSet pBackglass,"notetitle","Save Will",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","Take down the keep out barricade, then bash the targets 15 times.",1,""
		Case 15
			PuPlayer.LabelSet pBackglass,"notetitle","Escape the Bad Men",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","Hit the Bad Men targets on the left and lock balls to escape the Bad Men.",1,""
		Case 24
			PuPlayer.LabelSet pBackglass,"notetitle","Gather the Adventuring Party",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","Every shot on the table corresponds to a party member in the middle.",1,""
		Case 32
			PuPlayer.LabelSet pBackglass,"notetitle","Visit the A.V. Club",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","The A.V. Club Is where you'll find fun modes. Shoot the ramps to power up the battery.",1,""
		Case 40
			PuPlayer.LabelSet pBackglass,"notetitle","Collect Eggos for Mystery Roll",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","The mystery is super valuable. Collect waffles to light a mystery roll in castle byers.",1,""
		Case 48
			PuPlayer.LabelSet pBackglass,"notetitle","Need Another Kickback?",1,""
			PuPlayer.LabelSet pBackglass,"notecopy","If your kickback is empty then fill up your lane light to reactivate it!",1,""
		Case 56
			rulesposition = 0
	End Select
	End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  SECONDARY HIT EVENTS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 

	Dim RStep, Lstep

	Sub RightSlingShot1_slingshot
		PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors), RightInlane
	End Sub

	Sub LeftSlingShot1_slingshot
		PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors), LeftInlane
	End Sub

	Sub RightSlingShot_Slingshot
		PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors), RightInlane
		PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
		RSling.Visible = 0
		RSling1.Visible = 1
		sling1.TransZ = -20
		RStep = 0
		RightSlingShot.TimerEnabled = 1

	End Sub

	Sub RightSlingShot_Timer
		Select Case RStep
			Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
			Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
		End Select
		RStep = RStep + 1
	End Sub

	Sub LeftSlingShot_Slingshot
		PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors), LeftInlane
		PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
		LSling.Visible = 0
		LSling1.Visible = 1
		sling2.TransZ = -20
		LStep = 0
		LeftSlingShot.TimerEnabled = 1

	End Sub

	Sub LeftSlingShot_Timer
		Select Case LStep
			Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
			Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
		End Select
		LStep = LStep + 1
	End Sub


	Sub RotateLaneLightsLeft
		Dim TempState
		TempState = l6.State
		l6.State = l7.State
		l7.State = l8.State
		l8.State = l9.State
		l9.State = TempState
	End Sub

	Sub RotateLaneLightsRight
		Dim TempState
		TempState = l9.State
		l9.State = l8.State
		l8.State = l7.State
		l7.State = l6.State
		l6.State = TempState
	End Sub





'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  MAIN SHOTS - PRIMARY HIT EVENTS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 


'Poison Ivy Targets

Dim TargetHitCnt
Sub sw003_hit
TargetHitCnt = TargetHitCnt + 1
Select Case TargetHitCnt
Case 1
PL1.State = 1
Addscore 10000
PlaySound "fx_droptarget"
PuPlayer.playlistplayex pTopper,"ivytop","",100,1
'chilloutthemusic2
Case 2
PL2.State = 1
PlaySound "fx_droptarget"
Addscore 20000
PuPlayer.playlistplayex pTopper,"ivytop","",100,1
'chilloutthemusic2
Case 3
PL3.State = 1
PlaySound "fx_droptarget"
Addscore 30000
PuPlayer.playlistplayex pTopper,"ivytop","",100,1
'chilloutthemusic2
Case 4
PL4.State = 1
PlaySound "fx_droptarget"
Addscore 40000
PuPlayer.playlistplayex pTopper,"ivytop","",100,1
'chilloutthemusic2
Case 5
PL5.State = 1
PlaySound "fx_droptarget"
chilloutthemusic4
addscore 60000 ' award 60,000 points
PuPlayer.playlistplayex pTopper,"ivytopjail","",100,2
PuPlayer.playlistplayex pDMD,"ivydmd","@ivy.mp4",0,2
PuPlayer.playlistplayex pBackglass,"ivycap","",100,2
PL5.TimerInterval = 2000
PL5.TimerEnabled = True
End Select
End Sub

Sub sw003Reset
TargetHitCnt = 0
PL1.State = LightStateBlinking
PL2.State = LightStateBlinking
PL3.State = LightStateBlinking
PL4.State = LightStateBlinking
PL5.State = LightStateBlinking
poisongl.state = 1
EndGameCheck
End Sub

Sub PL5_Timer
sw003Reset
PL5.TimerEnabled = 0
End Sub

'Mr Freeze Targets

Dim TargetHitCnt1
Sub sw004_hit
TargetHitCnt1 = TargetHitCnt1 + 1
Select Case TargetHitCnt1
Case 1
FL1.State = 1
Addscore 10000
PlaySound "ice"
PuPlayer.playlistplayex pTopper,"fretop","",100,1
chilloutthemusic2
Case 2
FL2.State = 1
Addscore 20000
PlaySound "ice"
PuPlayer.playlistplayex pTopper,"fretop","",100,1
chilloutthemusic2
Case 3
FL3.State = 1
Addscore 30000
PlaySound "ice"
PuPlayer.playlistplayex pTopper,"fretop","",100,1
chilloutthemusic2
Case 4
FL4.State = 1
Addscore 40000
PlaySound "ice"
PuPlayer.playlistplayex pTopper,"fretop","",100,1
chilloutthemusic2
Case 5
FL5.State = 1
PlaySound "ice"
chilloutthemusic4
addscore 60000 ' award 60,000 points
PuPlayer.playlistplayex pTopper,"fretopjail","",100,2
PuPlayer.playlistplayex pDMD,"fredmd","fre.mp4",0,2
PuPlayer.playlistplayex pBackglass,"frecap","",100,2
FL5.TimerInterval = 2000
FL5.TimerEnabled = True
End Select
End Sub

Sub sw004Reset
TargetHitCnt1 = 0
FL1.State = LightStateBlinking
FL2.State = LightStateBlinking
FL3.State = LightStateBlinking
FL4.State = LightStateBlinking
FL5.State = LightStateBlinking
freezegl.state = 1
EndGameCheck
End Sub

Sub FL5_Timer
sw004Reset
FL5.TimerEnabled = 0
End Sub

'Bane Targets

Dim TargetHitCnt2
Sub sw4_hit
TargetHitCnt2 = TargetHitCnt2 + 1
Select Case TargetHitCnt2
Case 1
BL1.State = 1
Addscore 20000
PlaySound "fx_droptarget"
PuPlayer.playlistplayex pDMD,"rogdmd","hatter.mp4",100,1
'chilloutthemusic
Case 2
BL2.State = 1
PlaySound "fx_droptarget"
Addscore 40000
PuPlayer.playlistplayex pDMD,"rogdmd","clayface.mp4",100,1
'chilloutthemusic
Case 3
BL3.State = 1
PlaySound "fx_droptarget"
Addscore 60000
PuPlayer.playlistplayex pDMD,"rogdmd","face.mp4",100,1
'chilloutthemusic
Case 4
BL4.State = 1
BL5.State = 1
PlaySound "fx_droptarget"
chilloutthemusic5
Addscore 80000
PuPlayer.playlistplayex pDMD,"rogdmd","rogcap.mp4",100,1
PuPlayer.playlistplayex pBackglass,"rogcap","",100,1
PuPlayer.playlistplayex pTopper,"rogtop","",100,1
Case 5
BL6.State = 1
PlaySound "fx_droptarget"
Addscore 100000
PuPlayer.playlistplayex pDMD,"bandmd","bane.mp4",100,1
PuPlayer.playlistplayex pTopper,"bantop","",100,1
'chilloutthemusic
Case 6
BL7.State = 1
PlaySound "fx_droptarget"
Addscore 150000
PuPlayer.playlistplayex pDMD,"bandmd","bane.mp4",100,1
PuPlayer.playlistplayex pTopper,"bantop","",100,1
'chilloutthemusic
Case 7
BL8.State = 1
PlaySound "fx_droptarget"
Addscore 200000
PuPlayer.playlistplayex pDMD,"bandmd","bane.mp4",100,1
PuPlayer.playlistplayex pTopper,"bantop","",100,1
'chilloutthemusic
Case 8
BL9.State = 1
PlaySound "fx_droptarget"
Addscore 250000
PuPlayer.playlistplayex pDMD,"bandmd","bane.mp4",100,1
PuPlayer.playlistplayex pTopper,"bantop","",100,1
'chilloutthemusic
Case 9
BL10.State = 1
PlaySound "fx_droptarget"
Addscore 350000
chilloutthemusic5
PuPlayer.playlistplayex pDMD,"bandmd","bancap.mp4",100,2
PuPlayer.playlistplayex pTopper,"bantopjail","",100,2
PuPlayer.playlistplayex pBackglass,"bancap","",100,2

BL10.TimerInterval = 2000
BL10.TimerEnabled = True
End Select
End Sub

Sub sw4Reset
TargetHitCnt2 = 0
BL1.State = LightStateBlinking
BL2.State = LightStateBlinking
BL3.State = LightStateBlinking
BL4.State = LightStateBlinking
BL5.State = LightStateBlinking
BL6.State = LightStateBlinking
BL7.State = LightStateBlinking
BL8.State = LightStateBlinking
BL9.State = LightStateBlinking
BL10.State = LightStateBlinking
banegl.state = 1
EndGameCheck
End Sub

Sub BL10_Timer
sw4Reset
BL10.TimerEnabled = 0
End Sub


'RIDDLER TARGETS

Sub sw222_hit
	AddScore 50000
	sw222.IsDropped = True
	PlaySound "riddlemethis"
	PuPlayer.playlistplayex pTopper,"ridtop","",100,1
	PuPlayer.playlistplayex pDMD,"riddmd","rid.mp4",0,1
	Light013.State = LightStateOn
End Sub

Sub sw22_hit
	AddScore 100000
	sw22.IsDropped = True
	PlaySound "riddlemethis"
	PuPlayer.playlistplayex pTopper,"ridtop","",100,1
	PuPlayer.playlistplayex pDMD,"riddmd","rid.mp4",0,1
	Light011.State = LightStateOn
End Sub

Sub sw2_hit
	AddScore 200000
	sw2.IsDropped = True
	PlaySound "riddlemethis"
	PuPlayer.playlistplayex pTopper,"ridtop","",100,1
	PuPlayer.playlistplayex pDMD,"riddmd","rid.mp4",0,1
	Light012.State = LightStateOn
End Sub

Sub riddlerreset
	Light013.State = LightStateOff
	Light011.State = LightStateOff
	Light012.State = LightStateOff
End Sub


'JOKER TARGETS

Sub sw777_hit
	AddScore 10000
	PuPlayer.playlistplayex pTopper,"jokdmd","jok.mp4",100,1
	PuPlayer.playlistplayex pDMD,"jokdmd","jok.mp4",0,1
	sw777.IsDropped = True
	PlaySound "fx_droptarget"
	PlaySound "joker"
End Sub

Sub sw77_hit
	AddScore 10000
	PuPlayer.playlistplayex pTopper,"jokdmd","jok.mp4",100,1
	PuPlayer.playlistplayex pDMD,"jokdmd","jok.mp4",0,1
	sw77.IsDropped = True
	PlaySound "fx_droptarget"
	PlaySound "joker"
End Sub

Sub sw7_hit
	AddScore 10000
	PuPlayer.playlistplayex pTopper,"jokdmd","jok.mp4",100,1
	PuPlayer.playlistplayex pDMD,"jokdmd","jok.mp4",0,1
	sw7.IsDropped = True
	PlaySound "fx_droptarget"
	PlaySound "joker"
End Sub



'Question Marks

Dim TargetHitCnt5
Sub questionmarks_hit
TargetHitCnt5 = TargetHitCnt5 + 1
Select Case TargetHitCnt5
Case 1
Q1.State = 1
Addscore 100000
'PlaySound "fx_droptarget"
chilloutthemusic4
PuPlayer.playlistplayex pTopper,"ridtop2","",100,3
PuPlayer.playlistplayex pBackglass,"ridback","",100,3
PuPlayer.playlistplayex pDMD,"riddmd","ridcap2.mp4",100,3
Case 2
Q2.State = 1
Addscore 200000
'PlaySound "fx_droptarget"
chilloutthemusic4
PuPlayer.playlistplayex pTopper,"ridtop2","",100,3
PuPlayer.playlistplayex pBackglass,"ridback","",100,3
PuPlayer.playlistplayex pDMD,"riddmd","ridcap2.mp4",100,3
Case 3
Q3.State = 1
Addscore 500000
'PlaySound "fx_droptarget"
chilloutthemusic4
PuPlayer.playlistplayex pTopper,"ridtop2","",100,3
PuPlayer.playlistplayex pBackglass,"ridback","",100,3
PuPlayer.playlistplayex pDMD,"riddmd","ridcap2.mp4",100,3
Case 4
Q4.State = 1
Addscore 1000000
'PlaySound "fx_droptarget"
chilloutthemusic4
PuPlayer.playlistplayex pTopper,"ridtop2","",100,3
PuPlayer.playlistplayex pBackglass,"ridback","",100,3
PuPlayer.playlistplayex pDMD,"riddmd","ridcap2.mp4",100,3
Case 5
Q5.State = 1
Addscore 1000000
'PlaySound "fx_droptarget"
chilloutthemusic4
PuPlayer.playlistplayex pTopper,"ridtop2","",100,3
PuPlayer.playlistplayex pBackglass,"ridback","",100,3
PuPlayer.playlistplayex pDMD,"riddmd","ridcap2.mp4",100,3
Case 6
Q6.State = 1
Addscore 5000000
'PlaySound "fx_droptarget"
chilloutthemusic4
PuPlayer.playlistplayex pTopper,"ridtopjail","",100,4
PuPlayer.playlistplayex pBackglass,"ridcap","",100,4
PuPlayer.playlistplayex pDMD,"riddmd","ridcap2.mp4",100,4
Q6.TimerInterval = 1000
Q6.TimerEnabled = True
End Select
End Sub

Sub questionmarksreset
TargetHitCnt5 = 0
Q1.State = LightStateBlinking
Q2.State = LightStateBlinking
Q3.State = LightStateBlinking
Q4.State = LightStateBlinking
Q5.State = LightStateBlinking
Q6.State = LightStateBlinking
riddlergl.state = 1
EndGameCheck
End Sub

Sub Q6_Timer
questionmarksreset
Q6.TimerEnabled = 0
End Sub

'Penguin Targets

Sub TargetA_Hit() ' Target A was hit
PlaySound "quack"
'PuPlayer.playlistplayex pCallouts,"audiocallouts","quack.wav",100,1
'chilloutthemusic3
PuPlayer.playlistplayex pTopper,"pentop","",100,1
LightA.state = LightStateOn ' Light target A light
addscore 10000 ' add 1000 points to the score
TargetABCCheck() ' check to see if all targets have been hit
'chilloutthemusic4
'PuPlayer.playlistplayex pBackglass,"pencap","",100,1
'PuPlayer.playlistplayex pDMD,"pendmd","@pen.mp4",0,1
'PuPlayer.playlistplayex pCallouts,"popapill","pillsall.mp3",100,1
End Sub

Sub TargetB_Hit() ' Target B was hit
PlaySound "quack"
'PuPlayer.playlistplayex pCallouts,"audiocallouts","quack.wav",100,1
'chilloutthemusic3
PuPlayer.playlistplayex pTopper,"pentop","",100,1
LightB.state = LightStateOn ' Light target B light
addscore 10000 ' add 1000 points to the score
TargetABCCheck() ' check to see if all targets have been hit
'chilloutthemusic4
'PuPlayer.playlistplayex pBackglass,"pencap","",100,1
'PuPlayer.playlistplayex pDMD,"pendmd","@pen.mp4",0,1
'PuPlayer.playlistplayex pCallouts,"popapill","pillsall.mp3",100,1
End Sub


Sub TargetC_Hit() ' Target C was hit
PlaySound "quack"
'PuPlayer.playlistplayex pCallouts,"audiocallouts","quack.wav",100,1
'chilloutthemusic3
PuPlayer.playlistplayex pTopper,"pentop","",100,1
LightC.state = LightStateOn ' Light target C light
addscore 10000 ' add 1000 points to the score
TargetABCCheck() ' check to see if all targets have been hit
'chilloutthemusic4
'PuPlayer.playlistplayex pBackglass,"pencap","",100,1
'PuPlayer.playlistplayex pDMD,"pendmd","@pen.mp4",0,1
'PuPlayer.playlistplayex pCallouts,"popapill","pillsall.mp3",100,1
End Sub

Rem *** Reset Section ***
Rem This code does all the resetting of the lights and the targets on the table.
Rem It also resets the variables associated with those items.

Sub ResetTargets()
LightA.state = LightStateOff 'Turn off TargetA Light
LightB.state = LightStateOff 'Turn off TargetB Light
LightC.state = LightStateOff 'Turn off TargetC Light
End Sub

Rem *** Scoring and Checking Section ***
Rem This section handles the adding up of the players score as well as checking
Rem for the completion of events (such as dropping all targets in a target bank).

Sub TargetABCCheck() ' Check if all targets have been hit
if LightA.state = 1 and LightB.state = 1 and LightC.state = 1 Then ' If they have then
TargetsABCTimer.Enabled = 1 '1 = true so start timer to allow ball to clear area
end if
end Sub

Sub TargetsABCTimer_Timer()
penguingl.state = 1
EndGameCheck
ResetTargets() ' reset the targets and
chilloutthemusic4
'PlaySound "Targetup"
addscore 30000 ' award 30,000 points
PuPlayer.playlistplayex pTopper,"pentopjail","",100,2
PuPlayer.playlistplayex pDMD,"pendmd","pen.mp4",0,2
PuPlayer.playlistplayex pBackglass,"pencap","",100,2
'PuPlayer.playlistplayex pBackglass,"score","fifty.mov",0,1
TargetsABCTimer.Enabled = 0 '0 = False so the timer is switched off

End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  MODES
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 

' Kicker





Sub Kicker1_Hit
	'playsound "kicker_enter_center",0,0.1,0.12,0.05
'	PlaySound "kicker_enter_center", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'	AddScore 10000
	Kicker1.timerenabled = False
	Kicker1.timerenabled = True
End Sub	

Sub Kicker1_Timer
	Kicker1.timerenabled = False
	Kicker1.kick 270,30
	sw2.IsDropped = False
	sw22.IsDropped = False
	sw222.IsDropped = False 
	riddlerreset
End Sub


'LEFT OUTLANE KICKER

Sub Kicker001_Hit
	Kicker001.timerenabled = False
	Kicker001.timerenabled = True


'	pupDMDDisplay "balllock","Ball^is^Locked", "@balllocked.mp4",3, 1,10 
End Sub

Sub Kicker001_Timer
	Kicker001.timerenabled = False
	If LightTrig1.State = LightStateOn Then
	PlaySound "batman"
	PuPlayer.playlistplayex pDMD,"Topper","@ballsave.mp4",100,1
	PuPlayer.playlistplayex pTopper,"Topper","@ballsave.mp4",100,1
	AddScore 50000
	Kicker001.kick 0,30
	End If
	
	If LightTrig1.State = LightStateOff Then
	Kicker001.kick 180,30
	End If
	
End Sub


'RIGHT OUTLANE KICKER

Sub Kicker002_Hit


	Kicker002.timerenabled = False
	Kicker002.timerenabled = True
End Sub

Sub Kicker002_Timer
	Kicker002.timerenabled = False
	If LightTrig3.State = LightStateOn Then
	PlaySound "batman"
	PuPlayer.playlistplayex pDMD,"Topper","@ballsave.mp4",100,1
	PuPlayer.playlistplayex pTopper,"Topper","@ballsave.mp4",100,1
	AddScore 50000
	Kicker002.kick 0,30
	End If
	
	If LightTrig3.State = LightStateOff Then
	Kicker002.kick 180,30
	End If
	
End Sub


'BatCave Ramp 

Dim TargetHitCnt6
Sub batramp_hit
TargetHitCnt6 = TargetHitCnt6 + 1
Select Case TargetHitCnt6
Case 1
C1.State = 1
Addscore 25000
PlaySound "engine"
PuPlayer.playlistplayex pTopper,"cartop","",100,1
PuPlayer.playlistplayex pDMD,"cardmd","@carshort.mp4",0,1
Case 2
C2.State = 1
Addscore 40000
PlaySound "engine"
PuPlayer.playlistplayex pTopper,"cartop","",100,1
PuPlayer.playlistplayex pDMD,"cardmd","carshort.mp4",0,1
Case 3
C3.State = 1
Addscore 50000
PlaySound "engine"
PuPlayer.playlistplayex pTopper,"cartopjail","",100,1
PuPlayer.playlistplayex pDMD,"cardmd","car.mp4",0,1
PuPlayer.playlistplayex pBackglass,"carcap","",100,1
C3.TimerInterval = 2000
C3.TimerEnabled = True
End Select
End Sub

Sub batrampreset
TargetHitCnt6 = 0
C1.State = LightStateBlinking
C2.State = LightStateBlinking
C3.State = LightStateBlinking
batgl.state = 1
EndGameCheck
End Sub

Sub C3_Timer
batrampreset
C3.TimerEnabled = 0
End Sub

'Joker Ramp 

Dim TargetHitCnt7
Sub jokerramp_hit
TargetHitCnt7 = TargetHitCnt7 + 1
Select Case TargetHitCnt7
Case 1
J1.State = 1
Addscore 100000
PlaySound "laugh"
sw777.IsDropped = False
sw77.IsDropped = False
sw7.IsDropped = False
	PuPlayer.playlistplayex pTopper,"joktop","",100,1
	PuPlayer.playlistplayex pDMD,"jokdmd","@jok.mp4",0,1
Case 2
J2.State = 1
PlaySound "laugh"
Addscore 100000
sw777.IsDropped = False
sw77.IsDropped = False
sw7.IsDropped = False
	PuPlayer.playlistplayex pTopper,"joktop","",100,1
	PuPlayer.playlistplayex pDMD,"jokdmd","@jok.mp4",0,1
Case 3
J3.State = 1
'PlaySound "fx_droptarget"
Addscore 1000000
sw777.IsDropped = False
sw77.IsDropped = False
sw7.IsDropped = False
chilloutthemusic4
	PuPlayer.playlistplayex pTopper,"joktopjail","",100,2
	PuPlayer.playlistplayex pDMD,"jokdmd","jokcap.mp4",0,2
	PuPlayer.playlistplayex pBackglass,"jokcap","",100,2
J3.TimerInterval = 2000
J3.TimerEnabled = True
End Select
End Sub

Sub jokerrampreset
TargetHitCnt7 = 0
J1.State = LightStateBlinking
J2.State = LightStateBlinking
J3.State = LightStateBlinking
jokergl.state = 1
EndGameCheck
End Sub

Sub J3_Timer
jokerrampreset
J3.TimerEnabled = 0
End Sub


'XXXXXXXXXXXX                     SKILLSHOT   XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  BUMPERS
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 


	Sub Bumper1_Hit
		LightEffect 5
		PuPlayer.playlistplayex pTopper,"batarang","",100,1
		PuPlayer.playlistplayex pDMD,"batarangdmd","batarang.mp4",0,1
		If NOT Tilted Then
			PlaySoundAt SoundFXDOF("fx_bumper", 107, DOFPulse, DOFContactors), ActiveBall
			DOF 110, DOFPulse
			DOF 302, DOFPulse   'DOF MX - Bumper 1
			AddScore 1000
		End If
	End Sub

	Sub Bumper2_Hit
		LightEffect 5
		PuPlayer.playlistplayex pTopper,"batarang","",100,1
		PuPlayer.playlistplayex pDMD,"batarangdmd","batarang.mp4",0,1
		If NOT Tilted Then
			PlaySoundAt SoundFXDOF("fx_bumper", 109, DOFPulse, DOFContactors), ActiveBall
			DOF 111, DOFPulse
			DOF 303, DOFPulse   'DOF MX - Bumper 2
			AddScore 1000
		End If
	End Sub

	Sub Bumper3_Hit
		LightEffect 5
		PuPlayer.playlistplayex pTopper,"batarang","",100,1
		PuPlayer.playlistplayex pDMD,"batarangdmd","batarang.mp4",0,1
		If NOT Tilted Then
			PlaySoundAt SoundFXDOF("fx_bumper", 108, DOFPulse, DOFContactors), ActiveBall
			DOF 112, DOFPulse
			DOF 304, DOFPulse   'DOF MX - Bumper 3
			AddScore 1000
		End If
	End Sub

	Sub Bumper4_Hit
		LightEffect 5
		If NOT Tilted Then
			PlaySoundAt SoundFXDOF("fx_bumper", 108, DOFPulse, DOFContactors), ActiveBall
			DOF 112, DOFPulse
			DOF 304, DOFPulse   'DOF MX - Bumper 3
			AddScore 1000
		End If
	End Sub

	Sub Bumper5_Hit
		LightEffect 5
		If NOT Tilted Then
			PlaySoundAt SoundFXDOF("fx_bumper", 108, DOFPulse, DOFContactors), ActiveBall
			DOF 112, DOFPulse
			DOF 304, DOFPulse   'DOF MX - Bumper 3
			AddScore 1000
		End If
	End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  FLIPPER LANES
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 

Dim LightTrig ' for rollover lanes
Dim TempState ' Temp varaible
Dim TempState2 'Temp variable
Dim TempState3
Dim TempState4
Dim LightTrigs(8) 'set the amount of lights. Don't forget there's a zero!
Set LightTrigs(0)=LightTrig5 'LeftInlane1
Set LightTrigs(1)=LightTrig6 'LeftInlane1
Set LightTrigs(2)=LightTrig7 'LeftInlane
Set LightTrigs(3)=LightTrig8 'RightInlane
Set LightTrigs(4)=LightTrig1 'LeftInlane1
Set LightTrigs(5)=LightTrig2 'LeftInlane
Set LightTrigs(6)=LightTrig3 'RightInlane
Set LightTrigs(7)=LightTrig4 'RightInlane1
Set LightTrigs(8)=LightTrig9 'RightInLane1
'Sub NewBall()
'lanelitesout ' Turns off 1-2-3-4-5 lane lights.
'End Sub


'*** 4 Lane Control Subroutines ***

Sub CheckLightTrig6789()
LightTrig6.State = 1
LightTrig7.State = 1
LightTrig8.State = 1
LightTrig9.State = 1
addscore 100000

end Sub
'********************************************


'********************************************

Sub ShiftLanesLeft() 'LightTrigs move right with flipper press

TempState = LightTrig6.State
LightTrig6.State = LightTrig7.State
LightTrig7.State = LightTrig8.State
LightTrig8.State = LightTrig9.State
LightTrig9.State = TempState

TempState3 = LightTrig1.State
LightTrig1.State = LightTrig2.State
LightTrig2.State = LightTrig3.State
LightTrig3.State = LightTrig4.State
LightTrig4.State = LightTrig5.State
LightTrig5.State = TempState3

End Sub ' Bullet6 end of Sub LightTrigs


Sub ShiftLanesRight() 'LightTrigs move left with flipper press

TempState2 = LightTrig9.State
LightTrig9.State = LightTrig8.State
LightTrig8.State = LightTrig7.State
LightTrig7.State = LightTrig6.State
LightTrig6.State = TempState2

TempState4 = LightTrig5.State
LightTrig5.State = LightTrig4.State
LightTrig4.State = LightTrig3.State
LightTrig3.State = LightTrig2.State
LightTrig2.State = LightTrig1.State
LightTrig1.State = TempState4

End Sub ' End of Sub ShiftLanesRight

'********************************************

Sub lanelitesout() 'LightTrigs turn off


LightTrig6.State = 0
LightTrig7.State = 0
LightTrig8.State = 0
LightTrig9.State = 0
PuPlayer.playlistplayex pCallouts,"robincallouts","",100,1
End Sub ' Bullet6 end of Sub laneliteout


'*** Trigger Hit event code ***

Sub LeftInlane_Hit()
'checks for 4 light in rollover with lane change

if tilt = false Then addscore 2500 'If not tilted then

'here's the 1-2-3-4-5 rollovers check
LightTrig7.State = 1' Then 'check if light is off
PlaySound "lane"
'chilloutthemusic2


If LightTrig6.state = 1 and LightTrig7.State = 1 and LightTrig8.state = 1 and LightTrig9.state = 1 then addscore 50000
If LightTrig6.state = 1 and LightTrig7.State = 1 and LightTrig8.state = 1 and LightTrig9.state = 1 then lanelitesout()

'PuPlayer.playlistplayex pCallouts,"robincallouts","",100,1
'LightRollover(1) 'go to light control sub
'End if

End Sub
'End of Sub LeftInlane_Hit

'********************************************


Sub LeftInLane1_Hit()
'checks for 5 light in rollover with lane change

If tilt = False Then addscore 5000 'if not tilted then

'here's the 1-2-3-4-5 rollovers check
LightTrig6.State = 1' Then 'check if light is off
PlaySound "lane"
'chilloutthemusic2
'PuPlayer.playlistplayex pCallouts,"catwomancallouts","",100,1

If LightTrig6.state = 1 and LightTrig7.State = 1 and LightTrig8.state = 1 and LightTrig9.state = 1 then addscore 50000
If LightTrig6.state = 1 and LightTrig7.State = 1 and LightTrig8.state = 1 and LightTrig9.state = 1 then lanelitesout()


'PuPlayer.playlistplayex pCallouts,"robincallouts","",100,1
'LightRollover(2) 'go to light control sub
'End If

End Sub
'Bullet6 end of Sub RightInlane_Hit

'********************************************

Sub RightInLane_Hit()
'checks for 1 light in rollover with lane change

if tilt = false Then addscore 2500 'If not tilted then
LightTrig8.State = 1' Then 'check if light is off
PlaySound "lane"
'PuPlayer.playlistplayex pCallouts,"harleycallouts","",100,1

If LightTrig6.state = 1 and LightTrig7.State = 1 and LightTrig8.state = 1 and LightTrig9.state = 1 then addscore 50000
If LightTrig6.state = 1 and LightTrig7.State = 1 and LightTrig8.state = 1 and LightTrig9.state = 1 then lanelitesout() 



'PuPlayer.playlistplayex pCallouts,"robincallouts","",100,1
'LightRollover(3) 'go to light control sub
'End if
End Sub
'End of Sub Trig1_Hit

'********************************************


Sub RightInLane1_Hit()
'checks for 2 light in rollover with lane change

If tilt = False Then addscore 5000 'if not tilted then

LightTrig9.State = 1'Then 'check if light is off
PlaySound "lane"
'chilloutthemusic2
'PuPlayer.playlistplayex pCallouts,"harleycallouts","",100,1

If LightTrig6.state = 1 and LightTrig7.State = 1 and LightTrig8.state = 1 and LightTrig9.state = 1 then addscore 50000
If LightTrig6.state = 1 and LightTrig7.State = 1 and LightTrig8.state = 1 and LightTrig9.state = 1 then lanelitesout() 

'PuPlayer.playlistplayex pCallouts,"robincallouts","",100,1
'LightRollover(2) 'go to light control sub
'End If


End Sub


'**************************************************************
'END GAME GOALS
'**************************************************************

Sub EndGameCheck() ' Check if all targets have been hit
if penguingl.state = 1 and freezegl.state = 1 and poisongl.state = 1 and banegl.state = 1 and jokergl.state = 1 and batgl.state = 1 and riddlergl.state = 1 Then ' If they have then
symbolgl.state = 1
goalwall.IsDropped = 1
end if
end Sub

Sub goalwallraise
if symbolgl.state = 1 Then
goalwall.IsDropped = 0
End If
End Sub


Sub goal_hit
	If symbolgl.State = LightStateOn Then
	chilloutthemusic4
	PuPlayer.playlistplayex pBackglass,"20mil","20mil.mp4",100,5
	PuPlayer.playlistplayex pDMD,"20mildmd","20mil.mp4",0,5
	PuPlayer.playlistplayex pTopper,"20mildmd","20mil.mp4",0,5
	AddScore 20000000

symbolgl.TimerInterval = 21000
symbolgl.TimerEnabled = True





	End If

	If symbolgl.State = LightStateOff Then
	chilloutthemusic4
	PuPlayer.playlistplayex pCallouts,"catwomancallouts","worthy.mp3",100,5
	PuPlayer.playlistplayex pBackglass,"20mil","worthy.mp4",0,5
	PuPlayer.playlistplayex pTopper,"20mil","worthy.mp4",0,5
	End If
end Sub

Sub goalReset

	penguingl.state = lightstateblinking
	freezegl.state = lightstateblinking
	poisongl.state = lightstateblinking
	banegl.state = lightstateblinking
	jokergl.state = lightstateblinking
	batgl.state = lightstateblinking
	riddlergl.state = lightstateblinking
	symbolgl.state = 0
	goalwall.IsDropped = 0
End Sub

Sub symbolgl_Timer
goalReset
symbolgl.TimerEnabled = 0
End Sub

Sub Kicker003_Hit
	Kicker003.timerenabled = False
	Kicker003.timerenabled = True

'	pupDMDDisplay "balllock","Ball^is^Locked", "@balllocked.mp4",3, 1,10 
End Sub

Sub goalwall_Timer
	goalwall.timerenabled = True
end Sub

Sub Kicker003_Timer
	Kicker003.timerenabled = False
'	If PL1.State = LightStateOn Then
	Kicker003.kick 180,20
'	End If
	
'	If symbolgl.State = LightStateOff Then
'	Kicker003.kick 180,20
'	End If
	
End Sub




'********************************************


'	Sub LeftInlane_Hit
'		PlaySoundAt "fx_sensor", ActiveBall
'		If bMultiBallMode = False Then
'			PlaySound "lane"
'			DOF 144, DOFPulse
'			DOF 313, DOFPulse   'DOF MX - Left Outer Lane
'		End If
'		If Tilted Then Exit Sub
'		LaneBonus = LaneBonus + 1
'		'll1.State = 1
'		'll8.State = 1
'		AddScore 50050
'		' Do some sound or light effect
'		LastSwitchHit = "lane1"
'	End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  WIZARD MODE
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
' 


'*************  starts PUP system,  must be called AFTER b2s/controller running so put in last line of table1_init
Sub PuPInit

Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")   
PuPlayer.B2SInit "", cGameName

if (PuPDMDDriverType=pDMDTypeReal) and (useRealDMDScale=1) Then 
       PuPlayer.setScreenEx pDMD,0,0,128,32,0  'if hardware set the dmd to 128,32
End if

PuPlayer.LabelInit pDMD


if PuPDMDDriverType=pDMDTypeReal then
Set PUPDMDObject = CreateObject("PUPDMDControl.DMD") 
PUPDMDObject.DMDOpen
PUPDMDObject.DMDPuPMirror
PUPDMDObject.DMDPuPTextMirror
PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 1, ""FN"":33 }"             'set pupdmd for mirror and hide behind other pups
PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 1, ""FN"":32, ""FQ"":3 }"   'set no antialias on font render if real
END IF


pSetPageLayouts

pDMDSetPage(pDMDBlank)   'set blank text overlay page.
pAttractStart


End Sub 'end PUPINIT



'PinUP Player DMD Helper Functions

Sub pDMDLabelHide(labName)
PuPlayer.LabelSet pDMD,labName,"",0,""   
end sub




Sub pDMDScrollBig(msgText,timeSec,mColor)
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':" & (timeSec*1000000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
end sub

Sub pDMDScrollBigV(msgText,timeSec,mColor)
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':2,'yps':1,'ype':-1,'len':" & (timeSec*1000000) & ",'mlen':" & (timeSec*1000) & ",'tt':0,'fc':" & mColor & "}"
end sub


Sub pDMDSplashScore(msgText,timeSec,mColor)
PuPlayer.LabelSet pDMD,"MsgScore",msgText,0,"{'mt':1,'at':1,'fq':250,'len':"& (timeSec*1000) &",'fc':" & mColor & "}"
end Sub

Sub pDMDSplashScoreScroll(msgText,timeSec,mColor)
PuPlayer.LabelSet pDMD,"MsgScore",msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':"& (timeSec*1000) &", 'mlen':"& (timeSec*1000) &",'tt':0, 'fc':" & mColor & "}"
end Sub

Sub pDMDZoomBig(msgText,timeSec,mColor)  'new Zoom
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':3,'hstart':5,'hend':80,'len':" & (timeSec*1000) & ",'mlen':" & (timeSec*500) & ",'tt':5,'fc':" & mColor & "}"
end sub

Sub pDMDTargetLettersInfo(msgText,msgInfo, timeSec)  'msgInfo = '0211'  0= layer 1, 1=layer 2, 2=top layer3.
'this function is when you want to hilite spelled words.  Like B O N U S but have O S hilited as already hit markers... see example.
PuPlayer.LabelShowPage pDMD,5,timeSec,""  'show page 5
Dim backText
Dim middleText
Dim flashText
Dim curChar
Dim i
Dim offchars:offchars=0
Dim spaces:spaces=" "  'set this to 1 or more depends on font space width.  only works with certain fonts
                          'if using a fixed font width then set spaces to just one space.

For i=1 To Len(msgInfo)
    curChar="" & Mid(msgInfo,i,1)
    if curChar="0" Then
            backText=backText & Mid(msgText,i,1)
            middleText=middleText & spaces
            flashText=flashText & spaces          
            offchars=offchars+1
    End If
    if curChar="1" Then
            backText=backText & spaces
            middleText=middleText & Mid(msgText,i,1)
            flashText=flashText & spaces
    End If
    if curChar="2" Then
            backText=backText & spaces
            middleText=middleText & spaces
            flashText=flashText & Mid(msgText,i,1)
    End If   
Next 

if offchars=0 Then 'all litup!... flash entire string
   backText=""
   middleText=""
   FlashText=msgText
end if  

PuPlayer.LabelSet pDMD,"Back5"  ,backText  ,1,""
PuPlayer.LabelSet pDMD,"Middle5",middleText,1,""
PuPlayer.LabelSet pDMD,"Flash5" ,flashText ,0,"{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) & "}"   
end Sub


Sub pDMDSetPage(pagenum)    
    PuPlayer.LabelShowPage pDMD,pagenum,0,""   'set page to blank 0 page if want off
    PDMDCurPage=pagenum
end Sub

Sub pHideOverlayText(pDisp)
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "& pDisp &", ""FN"": 34 }"             'hideoverlay text during next videoplay on DMD auto return
end Sub



Sub pDMDShowLines3(msgText,msgText2,msgText3,timeSec)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,3,timeSec,""
PuPlayer.LabelSet pDMD,"Splash3a",msgText,vis,pLine1Ani
PuPlayer.LabelSet pDMD,"Splash3b",msgText2,vis,pLine2Ani
PuPlayer.LabelSet pDMD,"Splash3c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDShowLines2(msgText,msgText2,timeSec)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,4,timeSec,""
PuPlayer.LabelSet pDMD,"Splash4a",msgText,vis,pLine1Ani
PuPlayer.LabelSet pDMD,"Splash4b",msgText2,vis,pLine2Ani
end Sub

Sub pDMDShowCounter(msgText,msgText2,msgText3,timeSec)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,6,timeSec,""
PuPlayer.LabelSet pDMD,"Splash6a",msgText,vis, pLine1Ani
PuPlayer.LabelSet pDMD,"Splash6b",msgText2,vis,pLine2Ani
PuPlayer.LabelSet pDMD,"Splash6c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDShowBig(msgText,timeSec, mColor)
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,vis,pLine1Ani
end sub


Sub pDMDShowHS(msgText,msgText2,msgText3,timeSec) 'High Score
Dim vis:vis=1
if pLine1Ani<>"" Then vis=0
PuPlayer.LabelShowPage pDMD,7,timeSec,""
PuPlayer.LabelSet pDMD,"Splash7a",msgText,vis,pLine1Ani
PuPlayer.LabelSet pDMD,"Splash7b",msgText2,vis,pLine2Ani
PuPlayer.LabelSet pDMD,"Splash7c",msgText3,vis,pLine3Ani
end Sub


Sub pDMDSetBackFrame(fname)
  PuPlayer.playlistplayex pDMD,"PUPFrames",fname,0,1    
end Sub

Sub pDMDStartBackLoop(fPlayList,fname)
  PuPlayer.playlistplayex pDMD,fPlayList,fname,0,1
  PuPlayer.SetBackGround pDMD,1
end Sub

Sub pDMDStopBackLoop
  PuPlayer.SetBackGround pDMD,0
  PuPlayer.playstop pDMD
end Sub


Dim pNumLines

'Theme Colors for Text (not used currenlty,  use the |<colornum> in text labels for colouring.
Dim SpecialInfo
Dim pLine1Color : pLine1Color=8454143  
Dim pLine2Color : pLine2Color=8454143
Dim pLine3Color :  pLine3Color=8454143
Dim curLine1Color: curLine1Color=pLine1Color  'can change later
Dim curLine2Color: curLine2Color=pLine2Color  'can change later
Dim curLine3Color: curLine3Color=pLine3Color  'can change later


Dim pDMDCurPriority: pDMDCurPriority =-1
Dim pDMDDefVolume: pDMDDefVolume = 0   'default no audio on pDMD

Dim pLine1
Dim pLine2
Dim pLine3
Dim pLine1Ani
Dim pLine2Ani
Dim pLine3Ani

Dim PriorityReset:PriorityReset=-1
DIM pAttractReset:pAttractReset=-1
DIM pAttractBetween: pAttractBetween=2000 '1 second between calls to next attract page
DIM pDMDVideoPlaying: pDMDVideoPlaying=False


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X� X  X  X� X  X  X  X� X  X  X  X  X� X  X  X  X  X  X� X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'� Orbital Scoreboard Code
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\
' X� X  X  X� X  X  X  X� X  X  X  X  X� X  X  X  X  X  X� X  X  X  X  

	'****************************
	' POST SCORES
	'****************************
	Dim osbtemp:osbtemp = osbdefinit
	Dim osbtempscore:osbtempscore = 0
	Sub SubmitOSBScore
		On Error Resume Next
		if osbactive = 1 or osbactive = 2 Then
		Dim objXmlHttpMain, Url, strJSONToSend 

		Url = "https://hook.integromat.com/82bu988v9grj31vxjklh2e4s6h97rnu0"

		strJSONToSend = "{""auth"":""" & osbkey &""",""player id"": """ & osbid & """,""player initials"": """ & osbtemp &""",""score"": " & CStr(osbtempscore) & ",""table"":"""& TableName & """,""version"":""" & myVersion & """}"

		Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
		objXmlHttpMain.open "PUT",Url, False
		objXmlHttpMain.setRequestHeader "Content-Type", "application/json"
		objXmlHttpMain.setRequestHeader "application", "application/json"

		objXmlHttpMain.send strJSONToSend
		end if
	End Sub	



	'****************************
	' GET SCORES
	'****************************
	dim worldscores

	Sub GetScores()
		On Error Resume Next
		Dim objXmlHttpMain, Url, strJSONToSend 

		Url = "https://hook.integromat.com/97fgba49iff2ucr2lqy72om9j0cha2wj"

		strJSONToSend = "{""auth"":"""& osbkey &""", ""table"":"""& TableName & """, ""range"":""combo"", ""format"":""csv"", ""number"":""10""}"

		Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
		objXmlHttpMain.open "PUT",Url, False
		objXmlHttpMain.setRequestHeader "Content-Type", "application/json"
		objXmlHttpMain.setRequestHeader "application", "application/json"

		objXmlHttpMain.send strJSONToSend

		worldscores = objXmlHttpMain.responseText
		vpmtimer.addtimer 3000, "showsuccess '"
		debug.print "got the scores"
		splitscores
	End Sub	

	Dim scorevar(40)
	sub emptyscores
		dim i
		For i = 0 to 40
			scorevar(i) = 0
		Next
	End Sub
	emptyscores

	Sub splitscores
		On Error Resume Next
		dim a,x
		a = Split(worldscores,",")
		Dim myNum:myNum = 0
		for each x in a
			myNum = MyNum + 1
			scorevar(MyNum) = x
			'debug.print x
		Next
	end Sub

	sub showsuccess
		pNote "OSB High Scores","Downloaded Successfully"
	end sub


	GetScores


'************************ where all the MAGIC goes,  pretty much call this everywhere  ****************************************
'*************************                see docs for examples                ************************************************
'****************************************   DONT TOUCH THIS CODE   ************************************************************

Sub pupDMDDisplay(pEventID, pText, VideoName,TimeSec, pAni,pPriority)
' pEventID = reference if application,  
' pText = "text to show" separate lines by ^ in same string
' VideoName "gameover.mp4" will play in background  "@gameover.mp4" will play and disable text during gameplay.
' also global variable useDMDVideos=true/false if user wishes only TEXT
' TimeSec how long to display msg in Seconds
' animation if any 0=none 1=Flasher
' also,  now can specify color of each line (when no animation).  "sometext|12345"  will set label to "sometext" and set color to 12345

DIM curPos
if pDMDCurPriority>pPriority then Exit Sub  'if something is being displayed that we don't want interrupted.  same level will interrupt.
pDMDCurPriority=pPriority
if timeSec=0 then timeSec=1 'don't allow page default page by accident


pLine1=""
pLine2=""
pLine3=""
pLine1Ani=""
pLine2Ani=""
pLine3Ani=""


if pAni=1 Then  'we flashy now aren't we
pLine1Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"  
pLine2Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"  
pLine3Ani="{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) &  "}"  
end If

curPos=InStr(pText,"^")   'Lets break apart the string if needed
if curPos>0 Then 
   pLine1=Left(pText,curPos-1) 
   pText=Right(pText,Len(pText) - curPos)
   
   curPos=InStr(pText,"^")   'Lets break apart the string
   if curPOS>0 Then
      pLine2=Left(pText,curPos-1) 
      pText=Right(pText,Len(pText) - curPos)

      curPos=InStr("^",pText)   'Lets break apart the string   
      if curPos>0 Then
         pline3=Left(pText,curPos-1) 
      Else 
        if pText<>"" Then pline3=pText 
      End if 
   Else 
      if pText<>"" Then pLine2=pText
   End if    
Else 
  pLine1=pText  'just one line with no break 
End if


'lets see how many lines to Show
pNumLines=0
if pLine1<>"" then pNumLines=pNumlines+1
if pLine2<>"" then pNumLines=pNumlines+1
if pLine3<>"" then pNumLines=pNumlines+1

if pDMDVideoPlaying Then 
			PuPlayer.playstop pDMD
			pDMDVideoPlaying=False
End if


if (VideoName<>"") and (useDMDVideos) Then  'we are showing a splash video instead of the text.
    
    PuPlayer.playlistplayex pDMD,"DMDSplash",VideoName,pDMDDefVolume,pPriority  'should be an attract background (no text is displayed)
    pDMDVideoPlaying=true
end if 'if showing a splash video with no text




if StrComp(pEventID,"shownum",1)=0 Then              'check eventIDs
    pDMDShowCounter pLine1,pLine2,pLine3,timeSec
Elseif StrComp(pEventID,"target",1)=0 Then              'check eventIDs
    pDMDTargetLettersInfo pLine1,pLine2,timeSec
Elseif StrComp(pEventID,"highscore",1)=0 Then              'check eventIDs
    pDMDShowHS pLine1,pLine2,pline3,timeSec
Elseif (pNumLines=3) Then                'depends on # of lines which one to use.  pAni=1 will flash.
    pDMDShowLines3 pLine1,pLine2,pLine3,TimeSec
Elseif (pNumLines=2) Then
    pDMDShowLines2 pLine1,pLine2,TimeSec
Elseif (pNumLines=1) Then
    pDMDShowBig pLine1,timeSec, curLine1Color
Else
    pDMDShowBig pLine1,timeSec, curLine1Color
End if

PriorityReset=TimeSec*1000
End Sub 'pupDMDDisplay message



Sub pupDMDupdate_Timer()
	pUpdateScores    

    if PriorityReset>0 Then  'for splashes we need to reset current prioirty on timer
       PriorityReset=PriorityReset-pupDMDUpdate.interval
       if PriorityReset<=0 Then 
            pDMDCurPriority=-1            
            if pInAttract then pAttractReset=pAttractBetween ' pAttractNext  call attract next after 1 second
			pDMDVideoPlaying=false			
			End if
    End if

    if pAttractReset>0 Then  'for splashes we need to reset current prioirty on timer
       pAttractReset=pAttractReset-pupDMDUpdate.interval
       if pAttractReset<=0 Then 
            pAttractReset=-1            
            if pInAttract then pAttractNext
			End if
    end if 
End Sub

Sub PuPEvent(EventNum)
if hasPUP=false then Exit Sub
PuPlayer.B2SData "E"&EventNum,1  'send event to puppack driver  
End Sub


'********************* END OF PUPDMD FRAMEWORK v1.0 *************************
'******************** DO NOT MODIFY STUFF ABOVE THIS LINE!!!! ***************
'****************************************************************************

'*****************************************************************
'   **********  PUPDMD  MODIFY THIS SECTION!!!  ***************
'PUPDMD Layout for each Table1
'Setup Pages.  Note if you use fonts they must be in FONTS folder of the pupVideos\tablename\FONTS  "case sensitive exact naming fonts!"
'*****************************************************************

Sub pSetPageLayouts

DIM dmddef
DIM dmdalt
DIM dmdscr
DIM dmdfixed

'labelNew <screen#>, <Labelname>, <fontName>,<size%>,<colour>,<rotation>,<xalign>,<yalign>,<xpos>,<ypos>,<PageNum>,<visible>
'***********************************************************************'
'<screen#>, in standard we�d set this to pDMD ( or 1)
'<Labelname>, your name of the label. keep it short no spaces (like 8 chars) although you can call it anything really. When setting the label you will use this labelname to access the label.
'<fontName> Windows font name, this must be exact match of OS front name. if you are using custom TTF fonts then double check the name of font names.
'<size%>, Height as a percent of display height. 20=20% of screen height.
'<colour>, integer value of windows color.
'<rotation>, degrees in tenths   (900=90 degrees)
'<xAlign>, 0= horizontal left align, 1 = center horizontal, 2= right horizontal
'<yAlign>, 0 = top, 1 = center, 2=bottom vertical alignment
'<xpos>, this should be 0, but if you want to �force� a position you can set this. it is a % of horizontal width. 20=20% of screen width.
'<ypos> same as xpos.
'<PageNum> IMPORTANT� this will assign this label to this �page� or group.
'<visible> initial state of label. visible=1 show, 0 = off.



if PuPDMDDriverType=pDMDTypeReal Then 'using RealDMD Mirroring.  **********  128x32 Real Color DMD  
	dmdalt="PKMN Pinball"
    dmdfixed="Instruction"
    dmdscr="Impact"    'main scorefont
	dmddef="Zig"

	'Page 1 (default score display)
  		 PuPlayer.LabelNew pDMD,"Credits" ,dmddef,20,33023   ,0,2,2,95,0,1,0
		 PuPlayer.LabelNew pDMD,"Play1"   ,dmdalt,21,59624   ,1,0,0,79,0,1,0
		 PuPlayer.LabelNew pDMD,"Ball"    ,dmdalt,21,59624  ,1,2,1,100,0,1,0
		 PuPlayer.LabelNew pDMD,"MsgScore",dmddef,45,33023   ,0,1,0, 0,40,1,0
		 PuPlayer.LabelNew pDMD,"CurScore",dmdscr,60,59624   ,0,1,2, 0,0,1,0


	'Page 2 (default Text Splash 1 Big Line)
		 PuPlayer.LabelNew pDMD,"Splash"  ,dmdalt,40,33023,0,1,1,0,0,2,0

	'Page 3 (default Text Splash 2 and 3 Lines)
		 PuPlayer.LabelNew pDMD,"Splash3a",dmddef,30,8454143,0,1,0,0,2,3,0
		 PuPlayer.LabelNew pDMD,"Splash3b",dmdalt,30,33023,0,1,0,0,30,3,0
	     PuPlayer.LabelNew pDMD,"Splash3c",dmdalt,25,33023,0,1,0,0,55,3,0


	'Page 4 (2 Line Gameplay DMD)
		 PuPlayer.LabelNew pDMD,"Splash4a",dmddef,40,8454143,0,1,0,0,0,4,0
	     PuPlayer.LabelNew pDMD,"Splash4b",dmddef,30,33023,0,1,2,0,75,4,0

	'Page 5 (3 layer large text for overlay targets function,  must you fixed width font!
		PuPlayer.LabelNew pDMD,"Back5"    ,dmdfixed,80,8421504,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Middle5"  ,dmdfixed,80,65535  ,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Flash5"   ,dmdfixed,80,65535  ,0,1,1,0,0,5,0

	'Page 6 (3 Lines for big # with two lines,  "19^Orbits^Count")
		PuPlayer.LabelNew pDMD,"Splash6a",dmddef,90,65280,0,0,0,15,1,6,0
		PuPlayer.LabelNew pDMD,"Splash6b",dmddef,50,33023,0,1,0,60,0,6,0
		PuPlayer.LabelNew pDMD,"Splash6c",dmddef,40,33023,0,1,0,60,50,6,0

 	'Page 7 (Show High Scores Fixed Fonts)
		PuPlayer.LabelNew pDMD,"Splash7a",dmddef,20,8454143,0,1,0,0,2,7,0
		PuPlayer.LabelNew pDMD,"Splash7b",dmdfixed,40,33023,0,1,0,0,20,7,0
		PuPlayer.LabelNew pDMD,"Splash7c",dmdfixed,40,33023,0,1,0,0,50,7,0


END IF  ' use PuPDMDDriver

if PuPDMDDriverType=pDMDTypeLCD THEN  'Using 4:1 Standard ratio LCD PuPDMD  ************ lcd **************

	'dmddef="Impact"
	dmdalt="PKMN Pinball"    
    dmdfixed="Instruction"
	dmdscr="Impact"  'main score font
	dmddef="Impact"

	'Page 1 (default score display)
		PuPlayer.LabelNew pDMD,"Credits" ,dmddef,20,33023   ,0,2,2,70,0,1,0
		PuPlayer.LabelNew pDMD,"Play1"   ,dmdalt,20,59624   ,1,0,0,0,0,1,0
		PuPlayer.LabelNew pDMD,"Ball"    ,dmdalt,20,59624   ,1,2,0,100,0,1,0
		PuPlayer.LabelNew pDMD,"MsgScore",dmddef,45,33023   ,0,1,0, 0,40,1,0
		PuPlayer.LabelNew pDMD,"CurScore",dmdscr,60,59624   ,0,1,1, 0,0,1,0


	'Page 2 (default Text Splash 1 Big Line)
		PuPlayer.LabelNew pDMD,"Splash"  ,dmdalt,40,33023,0,1,1,0,0,2,0

	'Page 3 (default Text 3 Lines)
		PuPlayer.LabelNew pDMD,"Splash3a",dmddef,30,8454143,0,1,0,0,2,3,0
		PuPlayer.LabelNew pDMD,"Splash3b",dmdalt,30,33023,0,1,0,0,30,3,0
		PuPlayer.LabelNew pDMD,"Splash3c",dmdalt,25,33023,0,1,0,0,57,3,0


	'Page 4 (default Text 2 Line)
		PuPlayer.LabelNew pDMD,"Splash4a",dmddef,40,8454143,0,1,0,0,0,4,0
		PuPlayer.LabelNew pDMD,"Splash4b",dmddef,30,33023,0,1,2,0,75,4,0

	'Page 5 (3 layer large text for overlay targets function,  must you fixed width font!
		PuPlayer.LabelNew pDMD,"Back5"    ,dmdfixed,80,8421504,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Middle5"  ,dmdfixed,80,65535  ,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Flash5"   ,dmdfixed,80,65535  ,0,1,1,0,0,5,0

	'Page 6 (3 Lines for big # with two lines,  "19^Orbits^Count")
		PuPlayer.LabelNew pDMD,"Splash6a",dmddef,90,65280,0,0,0,15,1,6,0
		PuPlayer.LabelNew pDMD,"Splash6b",dmddef,50,33023,0,1,0,60,0,6,0
		PuPlayer.LabelNew pDMD,"Splash6c",dmddef,40,33023,0,1,0,60,50,6,0

	'Page 7 (Show High Scores Fixed Fonts)
		PuPlayer.LabelNew pDMD,"Splash7a",dmddef,20,8454143,0,1,0,0,2,7,0
		PuPlayer.LabelNew pDMD,"Splash7b",dmdfixed,40,33023,0,1,0,0,20,7,0
		PuPlayer.LabelNew pDMD,"Splash7c",dmdfixed,40,33023,0,1,0,0,50,7,0


END IF  ' use PuPDMDDriver

if PuPDMDDriverType=pDMDTypeFULL THEN  'Using FULL BIG LCD PuPDMD  ************ lcd **************

	'dmddef="Impact"
	dmdalt="PKMN Pinball"    
    dmdfixed="Instruction"
	dmdscr="Impact"  'main score font
	dmddef="Impact"

	'Page 1 (default score display)
		PuPlayer.LabelNew pDMD,"Credits" ,dmddef,20,33023   ,0,2,2,70,0,1,0
		PuPlayer.LabelNew pDMD,"Play1"   ,dmdalt,20,59624   ,1,0,0,0,0,1,0
		PuPlayer.LabelNew pDMD,"Ball"    ,dmdalt,20,59624   ,1,2,0,100,0,1,0
		PuPlayer.LabelNew pDMD,"MsgScore",dmddef,45,33023   ,0,1,0, 0,40,1,0
		PuPlayer.LabelNew pDMD,"CurScore",dmdscr,50,59624   ,0,1,1, 0,0,1,0		


	'Page 2 (default Text Splash 1 Big Line)
		PuPlayer.LabelNew pDMD,"Splash"  ,dmdalt,40,33023,0,1,1,0,0,2,0

	'Page 3 (default Text 3 Lines)
		PuPlayer.LabelNew pDMD,"Splash3a",dmddef,30,8454143,0,1,0,0,2,3,0
		PuPlayer.LabelNew pDMD,"Splash3b",dmdalt,30,33023,0,1,0,0,30,3,0
		PuPlayer.LabelNew pDMD,"Splash3c",dmdalt,25,33023,0,1,0,0,57,3,0


	'Page 4 (default Text 2 Line)
		PuPlayer.LabelNew pDMD,"Splash4a",dmddef,40,8454143,0,1,0,0,0,4,0
		PuPlayer.LabelNew pDMD,"Splash4b",dmddef,30,33023,0,1,2,0,75,4,0

	'Page 5 (3 layer large text for overlay targets function,  must you fixed width font!
		PuPlayer.LabelNew pDMD,"Back5"    ,dmdfixed,80,8421504,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Middle5"  ,dmdfixed,80,65535  ,0,1,1,0,0,5,0
		PuPlayer.LabelNew pDMD,"Flash5"   ,dmdfixed,80,65535  ,0,1,1,0,0,5,0

	'Page 6 (3 Lines for big # with two lines,  "19^Orbits^Count")
		PuPlayer.LabelNew pDMD,"Splash6a",dmddef,90,65280,0,0,0,15,1,6,0
		PuPlayer.LabelNew pDMD,"Splash6b",dmddef,50,33023,0,1,0,60,0,6,0
		PuPlayer.LabelNew pDMD,"Splash6c",dmddef,40,33023,0,1,0,60,50,6,0

	'Page 7 (Show High Scores Fixed Fonts)
		PuPlayer.LabelNew pDMD,"Splash7a",dmddef,20,8454143,0,1,0,0,2,7,0
		PuPlayer.LabelNew pDMD,"Splash7b",dmdfixed,40,33023,0,1,0,0,20,7,0
		PuPlayer.LabelNew pDMD,"Splash7c",dmdfixed,40,33023,0,1,0,0,50,7,0



	'Full-screen DMD (1920x1200 panel, PUPDMDAspect = 0 on the cabinet): the sizes above already fit the panel.
	'Only a score in the billions runs wider. 'autow':94 makes PuP shrink any caption wider
	'than 94% of the screen until it fits, instead of running off both edges (same key on Windows PuP).
	'`u` leaves each caption as it is; the vis value repeats the label's LabelNew visibility.
	Dim fitLabel
	For Each fitLabel in Array("Credits","Play1","Ball","MsgScore","CurScore","Splash","Splash3a","Splash3b","Splash3c","Splash4a","Splash4b","Back5","Middle5","Flash5","Splash6a","Splash6b","Splash6c","Splash7a","Splash7b","Splash7c")
		PuPlayer.LabelSet pDMD,fitLabel,"`u`",0,"{'mt':2,'autow':94}"
	Next

END IF  ' use PuPDMDDriver




end Sub 'page Layouts


'*****************************************************************
'        PUPDMD Custom SUBS/Events for each Table1
'     **********    MODIFY THIS SECTION!!!  ***************
'*****************************************************************
'
'
'  we need to somewhere in code if applicable
'
'   call pDMDStartGame,pDMDStartBall,pGameOver,pAttractStart
'
'
'
'
'


Sub pDMDStartGame
pInAttract=False
pDMDSetPage(pScores)   'set blank text overlay page.

end Sub


Sub pDMDStartBall
end Sub

Sub pDMDGameOver
pAttractStart
end Sub

Sub pAttractStart

'pDMDSetPage(pDMDBlank)   'set blank text overlay page.

'pupDMDDisplay "attract","Welcome^To Goonies","@welcome.mp4",5,1,10
pInAttract=False 'Startup in AttractMode
end Sub

DIM pCurAttractPos: pCurAttractPos=0


'********************** gets called auto each page next and timed already in DMD_Timer.  make sure you use pupDMDDisplay or it wont advance auto.
Sub pAttractNext
pCurAttractPos=pCurAttractPos+1

  Select Case pCurAttractPos

'  Case 1 pupDMDDisplay "attract","Game Over","",3,0,10
'  Case 2 pupDMDDisplay "attract","Dunri^Wikan45","",3,0,10
'  Case 3 pupDMDDisplay "attract","Ricky^Dunigan","",2,0,10
'  Case 4 pupDMDDisplay "attract","Testing By^John Memeo","",3,1,10
'  Case 5 pupDMDDisplay "attract","Welcome^To Goonies","@welcome.mp4",3,1,10
  'Case 6 pupDMDDisplay "attract","Attract^6","",3,1,10
  'Case 7 pupDMDDisplay "attract","Attract^7","",2,0,10
  'Case 8 pupDMDDisplay "attract","Attract^8","",1,0,10
  'Case 9 pupDMDDisplay "attract","Attract^9","",1,1,10
  'Case 10 pupDMDDisplay "attract","Attract^10","",3,1,10
  Case Else
    pCurAttractPos=0
   pAttractNext 'reset to beginning
  end Select

end Sub


'************************ called during gameplay to update Scores ***************************
Dim CurTestScore:CurTestScore=100000
Sub pUpdateScores  'call this ONLY on timer 300ms is good enough
if pDMDCurPage <> pScores then Exit Sub
'puPlayer.LabelSet pDMD,"Credits","CREDITS " & ""& Credits ,1,""
'puPlayer.LabelSet pDMD,"Play1","Player 1",1,""
'puPlayer.LabelSet pDMD,"Ball"," "&pDMDCurPriority ,1,""

puPlayer.LabelSet pDMD,"CurScore","" & FormatNumber(Score(CurrentPlayer),0),1,""
puPlayer.LabelSet pDMD,"Play1"," player " & CurrentPlayer,1,""
puPlayer.LabelSet pDMD,"Ball"," ball "  & bpgcurrent - BallsRemaining(CurrentPlayer) + 1,1,""
end Sub



'BallsRemaining(CurrentPlayer) + 1

PUPInit  'this should be called in table1_init at bottom after all else b2s/controller running.


'********************  pretty much only use pupDMDDisplay all over ************************   
' Sub pupDMDDisplay(pEventID, pText, VideoName,TimeSec, pAni,pPriority)
' pEventID = reference if application,  
' pText = "text to show" separate lines by ^ in same string
' VideoName "gameover.mp4" will play in background  "@gameover.mp4" will play and disable text during gameplay.
' also global variable useDMDVideos=true/false if user wishes only TEXT
' TimeSec how long to display msg in Seconds
' animation if any 0=none 1=Flasher
' also,  now can specify color of each line (when no animation).  "sometext|12345"  will set label to "sometext" and set color to 12345
'Samples
'pupDMDDisplay "shoot", "SHOOT AGAIN!", ", 3, 1, 10 
'pupDMDDisplay "default", "DATA GADGET LIT", "@DataGadgetLit.mp4", 3, 1, 10
'pupDMDDisplay "shoot", "SHOOT AGAIN!", "@shootagain.mp4", 3, 1, 10   
'pupDMDDisplay "balllock", "Ball^Locked|16744448", "", 5, 1, 10             '  5 seconds,  1=flash, 10=priority, ball is first line, locked on second and locked has custom color |
'pupDMDDisplay "balllock","Ball 2^is^Locked", "balllocked2.mp4",3, 1,10     '  3 seconds,  1=flash, play balllocked2.mp4 from dmdsplash folder, 
'pupDMDDisplay "balllock","Ball^is^Locked", "@balllocked.mp4",3, 1,10       '  3 seconds,  1=flash, play @balllocked.mp4 from dmdsplash folder, because @ text by default is hidden unless useDmDvideos is disabled.


'pupDMDDisplay "shownum", "3^More To|616744448^GOOOO", "", 5, 1, 10         ' "shownum" is special.  layout is line1=BIG NUMBER and line2,line3 are side two lines.  "4^Ramps^Left"

'pupDMDDisplay "target", "POTTER^110120", "blank.mp4", 10, 0, 10            ' 'target'...  first string is line,  second is 0=off,1=already on, 2=flash on for each character in line (count must match)

pupDMDDisplay "highscore", "High Score^AAA  2451654^BBB   2342342", "", 5, 0, 10            ' highscore is special  line1=text title like highscore, line2, line3 are fixed fonts to show AAA 123,123,123
'pupDMDDisplay "highscore", "High Score^AAA  2451654|616744448^BBB   2342342", "", 5, 0, 10  ' sames as above but notice how we use a custom color for text |
