' ****************************************************************
'                         VISUAL PINBALL X
'                          Zandys Arcade
'                       THE BLIZZARD OF OZZ
'		plain VPX script using core.vbs for supporting functions
'                         Version 1.0
' ****************************************************************
' DOF Triggers (MerlinRTP)
' E101		Left Flipper
' E102		Right Flipper
' E105		Left Slingshot
' E106		Right Slingshot
' E112		Left Target
' E113		Target 5 & 7 (Center Near Blizzard Drop Target009)
' E114		Target 8 & 10 (Right Ramp)
' E115		Target 9 (Ozzy Target)
' E116		Target 13 (Ball Lock)
' E117		Target 1
' E118		GI ON
' E120		Target Blizzard Drop & Spinner
' E121		Extraball, Special, Autofire & highscore
' E124		Spinner 2 (R)
' E125		Credits Exist or Freeplay Mode (ON/OFF)
' E126		Check WinBattle & Mr.Crowley Van
' E127		Skill Shot
' E128		SW 1 Hit (Top Lane Left)
' E129		SW 6 Hit (Top Lane Right)
' E130		SW 8 (Left Orbital)
' E131		SW 7 (Right Orbital)
' E132		SW 2 Hit (Left Outlane) - NOT LIT
' E133		SW 4 Hit (Left Inlane) - NOT LIT
' E134		SW 3 Hit (Right Inlane) - NOT LIT
' E135		SW 5 Hit (Right Outlane) - NOT LIT
' E136		Spinner 1 (L)
' E137		Bumper #3 (Center)
' E138		Bumper #1 (Upper Left)
' E139		Album Collection
' E140		Bumper #2 (Upper Right)
' E141		Auto Plunger
' E142		Release Dropped Ball
' E143		Multiball (ON/OFF)
' E144		Ball in Plunger Lane (ON/OFF)
' E323		Attract Mode (ON/OFF)
' E146		Ball Drain
' E157		Jackpot
' E158		Super Jackpot
' E159		TILT!
' E160		Game Over (ON/OFF)

' *******************************
Option Explicit
Randomize
SetLocale 1033 ' Force US format so math works out

Dim UseFlexDMD
UseFlexDMD = False		' Can Also run with PupDMD if you want


'***************** This is where you turn on SCORBIT *******
Const ScorbitActive	= 0 	' Is Scorbit Active, Change to 1	
'***********************************************************
dim Scorbit : Set Scorbit = New ScorbitIF
Const     ScorbitShowClaimQR	= 0 	' If Scorbit is active this will show a QR Code  on ball 1 that allows player to claim the active player from the app
Const     ScorbitUploadLog		= 0 	' Store local log and upload after the game is over 
Const     ScorbitAlternateUUID  = 0 	' Force Alternate UUID from Windows Machine and saves it in VPX Users directory (C:\Visual Pinball\User\ScorbitUUID.dat)
Dim GameModeStrTmp
Dim bOnTheFirstBallScorbit
Dim bUnpaired : bUnpaired = True

Dim Tracker_Ramps
Dim Tracker_Orbits
Dim Tracker_Targets
Dim Tracker_Bumpers
Dim Tracker_Slings
Dim Tracker_Spinners
Dim Tracker_BonusPoints
Dim Tracker_ComboHits
Dim Tracker_Burritos
Dim Tracker_Lanes
Dim Tracker_ComboValue

Const FontSizeEvents = 10

Const BallSize = 50  ' 50 is the normal size used in the core.vbs, VP kicker routines uses this value divided by 2
Const BallMass = 1.2   ' normal ball mass
Const IntroSongVolume = 0.5 ' 1 is full volume. Value is from 0 to 1

'FlexDMD in high or normal quality
'change it to True if you have an LCD screen, 256x64
'or keep it False if you have a real DMD at 128x32 in size
Const FlexDMDHighQuality = True

' Define any Constants
Const cGameName = "BlizzardOfOzz-Data"
Const TableName = "THE_BLIZZARD_OF_OZZ"
Const myVersion = "1.0"
Const MaxPlayers = 4     ' from 1 to 4
Const BallSaverTime = 20 ' in seconds
Const MaxMultiplier = 5  ' limit to 5x in this game, both bonus multiplier and playfield multiplier

Const MaxMultiballs = 9  ' max number of balls during multiballs
Dim fMusicVolume : fMusicVolume = .4
Dim bAttract	' boolean for attract mode running
Const nTrainSpeed = 25	' Determines how fast the ball comes out of the train -- higher is faster
Dim nOzzyDuration	' Dynamically sets the speech interval for Ozzy animation
Dim nDevilDuration	' Dynamically sets the speech interval for Devil animation
Dim nDemDuration	' Dynamically sets the speech interval for Devil animation
'----- General Sound Options -----
Const VolumeDial = 0.9			  'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 1.0		  'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 1.0		  'Level of ramp rolling volume. Value between 0 and 

'****** PuP Variables ******

Dim cPuPPack: Dim PUPStatus: PUPStatus=false ' dont edit this line!!!

Dim pupPackScreenFile
Dim ObjFso
Dim ObjFile
Dim PupType
Dim VRChoice
Dim bGameReady
Dim bEOB
Dim bLFHeld
Dim bRFHeld
Dim nSlowPC 
Dim bShotClear

Dim DMDType : DMDType = 99		' Set to 99 in case using FlexDMD
Dim nRulecards
Dim nRulecardPF
Const EOBFontSize = 12

Dim bSupressModeMessages
Dim bSupress3Ca
Dim bSupressModeProgress

Dim BallHandlingQueue : Set BallHandlingQueue = New vpwQueueManager
Dim EOBQueue : Set EOBQueue = New vpwQueueManager
Dim DMDQueue : Set DMDQueue = New vpwQueueManager
Dim HSQueue : Set HSQueue = New vpwQueueManager
Dim AudioQueue : Set AudioQueue = New vpwQueueManager
Dim GeneralPupQueue: Set GeneralPupQueue = New vpwQueueManager
Dim LightQueue: Set LightQueue = New vpwQueueManager
' Used to track progress for mode 5
dim Mode5Lights(9)
Dim TrackName
Dim TrackFontSize
Dim nWizardModeMultiplier


' StandAlone Variations
'Dim nEventFontSize
'Dim nInstantInfoFontSize
'Dim nAttractFontSize
'Dim nMessageFontSize
Const nDMDFontSize = 10

'Const nOffsetY = 0
Dim nOffsetX
Const nDTOffsetX = 0




'playevent  SNUM, playlist, filename, volume, priority, playtype, seconds, special""
'   PuPlayer.playevent pDMDFull,"RandomScoring","Fire Missile 1.mp4",100,1,0,0,""
'
'playevent allows you in table script to do almost everything you can like pupevent
'playtype is whats similar... maybe confusing for some new people, but pup-pack gurus will understand the following playtypes
'//playtype for triggers
'ptNormal=0;
'ptLoop=1;
'ptSplashReset=2;
'ptSplashResume=3;
'ptStopScreen=4;
'ptStopFile=5;
'ptSetBG=6;
'ptPlaySSF=7;
'ptSkipSameP=8;
'ptCustomFunc=9;
'ptForcePlay=10;
'ptQueueSameP=11;
'ptQueueAlways=12;

'*************************** PuP Settings for this table ********************************

cPuPPack = "BlizzardOfOzz"    ' name of the PuP-Pack / PuPVideos folder for this table

'//////////////////// PINUP PLAYER: STARTUP & CONTROL SECTION //////////////////////////


' Load the core.vbs for supporting Subs and functions
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
    On Error Goto 0
End Sub

Const LiveViewVRSim = 0  ' 0 = Default, 1 = View table in VR mode in "Live View Editor" 

' VR uses FlexDMD to display DMD.
Dim VR_Obj, VRMode

' Define Global Variables
Dim RemoveTrustPost
Dim bBLIZZARDMode
Dim bBlizzardPrepMode
Dim BlizzardLetters(8)
Dim nBlizzTargetCount
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim BonusPoints(4)
Dim BonusHeldPoints(4)
Dim BonusMultiplier(4)
Dim PlayfieldMultiplier(4)
Dim bBonusHeld
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Score(4)
Dim HighScore(4)
Dim HighScoreName(4)
Dim Jackpot(4)
Dim SuperJackpot
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim TotalGamesPlayed
Dim mBalls2Eject
Dim SkillshotValue(4)
Dim bAutoPlunger
Dim bInstantInfo
Dim bAttractMode
Dim bplasticsActive
Dim bWorldPFActive
Dim LaneSwitch(4) 

' Define Game Control Variables
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsInLock(4)
Dim BallsInHole
Dim nSong : nSong = RndNbr(55)
Dim bWizMode1Active
Dim bWizMode2Active
Dim bWizMode3Active
Dim bFinalWizModeActive
Dim bMainMultiballMode

' Define Game Flags
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
Dim bJackpot
Dim bSuper
Dim bFlamingBalls

' core.vbs variables
Dim plungerIM 'used mostly as an autofire plunger during multiballs
Dim cbRight   'captive ball
Dim bsJackal
Dim x
Dim ttSpinDisk

Dim bEnablePuP
bEnablePuP = True		' Leave this set to True


'*******************************************
'*****       SCORING       *****************
'*******************************************

Const SCORE_LANES = 1000
Const SCORE_RAMPS = 10000
Const SCORE_ORBITS = 7500
Const SCORE_TARGETS = 2000
'Const SCORE_BUMPERS = 1500
' Uses BumperValue(CurrentPlayer)
'Const SCORE_SPINNERS = 1000
' Uses spinnervalue(CurrentPlayer)
Const SCORE_MYSTERY = 25000
Const SCORE_COMBOHITS = 100000
Const SCORE_MODESCOMPLETED = 5000000


Const BONUS_SPINNERS = 3000
Const BONUS_TARGETS = 666
Const BONUS_BURRITOS = 8000
Const BONUS_RAMPS = 3333
Const BONUS_MODESCOMPLETED = 250000

'*******************************************
'*****Timers for animations*****************

dim countr1
countr1 = 0
'OZZtimer.enabled = 1
OZZtimer.enabled = False
OZZTimerstop.enabled = False

dim countr2
countr2 = 0
'DEVtimer.enabled = 1
DEVtimer.enabled = False
DEVTimerstop.enabled = False

dim countr3
countr3 = 0
'DEV2timer.enabled = 1
DEV2timer.enabled = False
DEV2Timerstop.enabled = False

dim countr4
countr4 = 0
DEMtimer.enabled = False
DEMTimerstop.enabled = False

dim countr
countr = 0
speakertimer.enabled = 1
dim countra
countra = 0
speaker2timer.enabled = 1

' *********************************************************************
'                Visual Pinball Defined Script Events
' *********************************************************************

Sub Table1_Init()
    LoadEM
    Dim i
    Randomize
    FlameActive = False
    InitIce
	DbgTracker "Game START "
    InitBatFlaps
'    Controller.Games("THE BLIZZARD OF OZ").Settings.Value("sound") = 1
    'Impulse Plunger as autoplunger
    Const IMPowerSetting = 66 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime

        .Random 1.5
        .InitExitSnd SoundFXDOF("fx_kicker", 123, DOFPulse, DOFContactors), SoundFXDOF("fx_solenoid", 123, DOFPulse, DOFContactors)
        .CreateEvents "plungerIM"
    End With

    Set cbRight = New cvpmCaptiveBall
    With cbRight
        .InitCaptive CapTrigger1, CapWall1, Array(CapKicker1, CapKicker1a), 0
        .NailedBalls = 1
        .ForceTrans = .9
        .MinForce = 3.5
        '.CreateEvents "cbRight"
        .Start
    End With
    CapKicker1.CreateSizedBallWithMass BallSize / 2, BallMass

    ' Jackal hole
    Set bsJackal = New cvpmTrough
    With bsJackal
        .size = 5
        .Initexit JackalHole, 160, nTrainSpeed
        '.InitExitVariance 2, 2
        .MaxBallsPerKick = 1
    End With

    ' Turn Table - Spinner disk
    Set ttSpinDisk = New cvpmTurnTable
    With ttSpinDisk
        .InitTurnTable SpinningDiskTrigger1, 35
        .spinCW = True
        .SpinUp = 5
        .SpinDown = 3
        .MotorOn = False
        .CreateEvents "ttSpinDisk"
    End With

    ' Misc. VP table objects Initialisation, droptargets, animations...
    VPObjects_Init

    ' load saved values, highscore, names, jackpot
    Loadhs

    ' Initalise the DMD display
    DMD_Init

	' Turn off Bumper Lights
	FlBumperFadeTarget(1) = 0
	FlBumperFadeTarget(2) = 0
	FlBumperFadeTarget(3) = 0

    ' freeplay or coins
    bFreePlay = False 'we want coins

    if bFreePlay Then DOF 125, DOFOn

     ' Init main variables and any other flags
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
    BallsInLock(1) = 0
    BallsInLock(2) = 0
    BallsInLock(3) = 0
    BallsInLock(4) = 0
    BallsInHole = 0
    LastSwitchHit = ""
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bBonusHeld = False
    bJustStarted = True
    bJackpot = False
    bInstantInfo = False
    stopfire
    InitBliz
    InitSnow
    InitFlameFollowers
    FlameActive = False
    tFlameFollow.Enabled = False


    ' set any lights for the attract mode
    GiOff
	pupinit


	if Not ScorbitActive Then
		if DMDType = 2 Then 
			GeneralPupQueue.Add "LoadAttract","LoadAttract",24,3000,0,0,0,False
			DMDQueue.Add "StartAttractMode","StartAttractMode",95,3100,0,0,0,False		
		Else
			DMDQueue.Add "StartAttractMode","StartAttractMode",95,1500,0,0,0,False		
		End If

	End If

	PlaySong "Mu_End"

	DMDQueue.Add "DisplayTrack","DisplayTrack",95,1000,0,0,0,False


	' Prevent User From Starting Game Before Pup Loads Properly
	'DMDQueue.Add "bGameReady = True","bGameReady = True",95,2000,0,0,0,False

    ' Start the RealTime timer
    RealTime.Enabled = 1

    ' Load table color
    LoadLut
	vpmtimer.addtimer 100, "HideLUT '"
    Fire1.Visible = False
    Fire2.Visible = False
    Fire3.Visible = False



End Sub


sub DisplayTrack
	TrackName = "Mr.Crowley Instrumental"
	PuPlayer.LabelSet pDMD,"JukeBox2a"," JUKEBOX MODE ",1,"{'mt':2,'color': " & cWhite &", 'xpos': 50, 'ypos': 82.5, 'fonth':10}"
	PuPlayer.LabelSet pDMD,"JukeBox2b",TrackName,1,"{'mt':2,'color': " & cOrange &", 'xpos': 50, 'ypos': 91, 'fonth':10 }"
	pDMDLabelSetColorGradient "JukeBox2a",  cYellow, cRed
	pDMDLabelSetColorGradientPercent "JukeBox2b",  cBlue, cPurple, 40
End Sub


'==================================================================================================================================
' Called when options are tweaked by the player. 
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts

' Table1.Option arguments are: 
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings
Dim RailChoice: RailChoice = True
Dim LeftOutlaneDifficulty,RightOutlaneDifficulty,nBallsPerGame
Sub Table1_OptionEvent(ByVal eventId)

	'Balls Per Game
	nBallsPerGame = Table1.Option("Balls Per Game", 0, 2, 1, 0, 0, Array("3 (Default)", "4", "5"))
	if bGameInPlay = False Then SetBallsPerGame nBallsPerGame

	'Difficulty
	RemoveTrustPost = Table1.Option("Remove Trust Post", 0, 1, 1, 0, 0, Array("False", "True (Default)"))
	CheckTrustPost 

	'VR 
	VRChoice = Table1.Option("VR Room", 0, 1, 1, 0, 0, Array("OFF (Default)", "Mega"))
	LoadVR

	' Display rulecards
	nSlowPC = Table1.Option("Slow PC", 0, 1, 1, 0, 0, Array("Off (Default)", "ON"))
	SetPCSpeed

	'Flaming Balls
	bFlamingBalls = Table1.Option("Flaming Balls", 0, 1, 1, 1, 0, Array("OFF", "On (Default)"))
    
    'Rails on/offset
    RailChoice = Table1.Option("Rails Visible", 0, 1, 1, 1, 0, Array("Cabinet", "True (Default)"))
	SetRails RailChoice

	' Display rulecards
	nRulecardPF = Table1.Option("Rules on Playfield", 0, 1, 1, 0, 0, Array("Off (Default)", "ON"))
	SetRulesPF nRuleCardPF


	' Display rulecards
	nRulecards = Table1.Option("Rules on Backglass", 0, 3, 1, 0, 0, Array("Backglass", "Page 1", "Page 2", "Page 3"))
	if DMDType <> 2 Then SetRuleCards nRuleCards

	'MusicVolume
	fMusicVolume = Table1.Option("Song Volume", 0, 1, .01, .4, 1)

	'Outlane Difficulty
	LeftOutlaneDifficulty = Table1.Option("Left Outlane Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Medium (Default)", "Hard"))
	UpdateLeftOutlanePosts LeftOutlaneDifficulty

	RightOutlaneDifficulty = Table1.Option("Right Outlane Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Medium (Default)", "Hard"))
	UpdateRightOutlanePosts RightOutlaneDifficulty

	' LUT Controls
	LUTImage = Table1.Option("LUT Image", 0, 22, 1, LUTImage, 0)
	NextLUT

	if EventID = 3 Then 
		bLutActive = False
		HideLUT
	End If
End Sub

Sub SetRails(Opt)
	Select Case Opt
		Case 0:
			lrail.Visible = 0
			rrail.Visible = 0
            Cabinetmode.Visible = 1		
		Case 1:
			lrail.Visible = 1
			rrail.Visible = 1
            Cabinetmode.Visible = 0				
	End Select
End Sub

Sub SetRulesPF(Opt)
	if Opt = 1 Then 
		Rules.Visible = True
	Else
		Rules.Visible = False
	End If
End Sub

Sub SetRuleCards(Opt)
	if DMDType = 2 or PlatformOS <> "windows" Then exit sub

	PuPlayer.playlistplayex pBackglass,"PuPOverlays","card"&Opt&".png",0,1
	if renderingmode = 2 then PinCab_Backglass.image = "card"&Opt
End Sub

Sub SetPCSpeed
	if nSlowPC Then
		pSetLowQualityPc	' sets performnance to lower speed
		bFlamingBalls = False
	End If
End Sub

Sub LoadVR
	If renderingmode = 2 or LiveViewVRSim = 1 Then
		VRMode = True

		lrail.Visible = 0
		rrail.Visible = 0 
		For Each VR_Obj in VRCabinet:VR_Obj.Visible = 1:Next
		For Each VR_Obj in VRWorld:VR_Obj.Visible = 1:Next

'		UseFlexDMD = False
			digitgrid.visible = 0
			digit041.visible = 0

			Dim VrObj
			For Each VrObj in DMDUpper
				VrObj.visible = 0

			Next
			For Each VrObj in DMDLower
				VrObj.visible = 0
			Next


	Else
		VRMode = False
		For Each VR_Obj in VRCabinet:VR_Obj.Visible = 0:Next
		For Each VR_Obj in VRWorld:VR_Obj.Visible = 0:Next
		lrail.Visible = 1
		rrail.Visible = 1
	End If
End Sub

Sub TestRTP2
	updatemodeprogress True
End Sub

'******************
' Captive Ball Subs
'******************
Sub CapTrigger1_Hit:cbRight.TrigHit ActiveBall:End Sub
Sub CapTrigger1_UnHit:cbRight.TrigHit 0:End Sub
Sub CapWall1_Hit:cbRight.BallHit ActiveBall:PlaySoundAtBall "Ball_Collide_1":End Sub
Sub CapKicker1a_Hit:cbRight.BallReturn Me:End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

	if keycode = "8" then testRTP
'    If keycode = LeftMagnaSave And bAttract = False Then MusicDown()
'	If keycode = RightMagnaSave And bAttract = False Then MusicUp()

    If keycode = LeftMagnaSave Then MusicDown()
	If keycode = RightMagnaSave Then MusicUp()


    If Keycode = AddCreditKey Then
        Credits = Credits + 1
        if bFreePlay = False Then DOF 125, DOFOn
        If(Tilted = False)Then
			'if bAttract Then StopAttractMode

                Select Case Int(rnd*3)
                        Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
                        Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
                        Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
                End Select

			if NOT bGameInPlay Then
				pDMDSplashTwoLines "CREDITS ", Credits , 2000, cRed
				DMDQueue.Add "ClearTwoLines","ClearTwoLines",45,2100,0,0,0,False

				DMDFlush
				DMD "_", CL("CREDITS: " & Credits), "", eNone, eNone, eNone, 500, True, ""
				ShowTableInfo
			End If
        End If
    End If


    If keycode = PlungerKey Then
        Plunger.Pullback
		SoundPlungerPull()
		If VRMode = True Then
			TimerVRPlunger.Enabled = True
			TimerVRPlunger2.Enabled = False
		End If
    End If

	If keycode = LeftFlipperKey Then
		bLFHeld = True
		if bEOB AND bRFHeld Then ExpediteEOB
		'if bEOB Then ExpediteEOB
		If VRMode = True Then PinCab_Flipper_Button_Left.X = PinCab_Flipper_Button_Left.X + 10
		FlipperActivate LeftFlipper, LFPress
	End If

	If keycode = RightFlipperKey Then
		bRFHeld = True
		if bEOB AND bLFHeld Then ExpediteEOB
		'if bEOB Then ExpediteEOB
		If VRMode = True Then PinCab_Flipper_Button_Right.X = PinCab_Flipper_Button_Right.X - 10
		FlipperActivate RightFlipper, RFPress
	End If

    If hsbModeActive Then
        EnterHighScoreKey(keycode)
        Exit Sub
    End If


    ' Table specific

    ' Normal flipper action

    If bGameInPlay AND NOT Tilted Then

        If keycode = LeftTiltKey Then Nudge 90, 8:SoundNudgeLeft:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 8:SoundNudgeRight:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 9:SoundNudgeCenter:CheckTilt

        If keycode = LeftFlipperKey Then

			FlipperActivate LeftFlipper, LFPress
			SolLFlipper 1
			InstantInfoTimer.Enabled = True
			RotateLaneSwitchLeft
		End If

        If keycode = RightFlipperKey Then
			FlipperActivate RightFlipper, RFPress
			SolRFlipper 1
			InstantInfoTimer.Enabled = True
			RotateLaneSwitchRight
		End If

        If keycode = StartGameKey AND bGameReady Then
            If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = True))Then

                If(bFreePlay = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
					PlayPlayersInGame
                    TotalGamesPlayed = TotalGamesPlayed + 1
					DMDUpdatePlayerName
                    DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 500, True, "so_fanfare1"
                Else
                    If(Credits > 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
						PlayPlayersInGame
                        TotalGamesPlayed = TotalGamesPlayed + 1
                        Credits = Credits - 1
						DMDUpdatePlayerName
                        DMD "_", CL(PlayersPlayingGame & " PLAYERS"), "", eNone, eBlink, eNone, 500, True, "so_fanfare1"
                        If Credits < 1 And bFreePlay = False Then DOF 125, DOFOff
                        Else
							if bAttract Then StopAttractMode
							NeedCoins
							pDMDSplashTwoLines "INSERT COINS", "", 2000, cRed
							DMDQueue.Add "ClearTwoLines","ClearTwoLines",45,2100,0,0,0,False
                            ' Not Enough Credits to start a game.
                            DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 500, True, "so_nocredits"
                    End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey AND bGameReady Then
                If(bFreePlay = True)Then
                    If(BallsOnPlayfield = 0)Then
                        ResetForNewGame()
                    End If
                Else
                    If(Credits > 0)Then
                        If(BallsOnPlayfield = 0)Then
                            Credits = Credits - 1
                            If Credits < 1 And bFreePlay = False Then DOF 125, DOFOff
                            ResetForNewGame()
                        End If
                    Else
						if bAttract Then StopAttractMode
						NeedCoins
						pDMDSplashTwoLines "INSERT COINS", "", 2000, cRed
						DMDQueue.Add "ClearTwoLines","ClearTwoLines",45,2100,0,0,0,False
                        ' Not Enough Credits to start a game.
                        DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 500, True, "so_nocredits"
                        ShowTableInfo
                    End If
                End If
            End If
    End If ' If (GameInPlay)

'test keys
End Sub

Sub Table1_KeyUp(ByVal keycode)
    'If keycode = LeftMagnaSave Then bLutActive = False: HideLUT

     If keycode = PlungerKey Then
        Plunger.Fire
        PlaySoundAt "fx_plunger", plunger
        If bBallInPlungerLane Then PlaySoundAt "fx_fire", plunger
		If VRMode = True Then
			TimerVRPlunger.Enabled = False
			TimerVRPlunger2.Enabled = True
			Pincab_Shooter.Y = 0
		End If
    End If

	If keycode = LeftFlipperKey Then
		bLFHeld = False
		If VRMode = True Then PinCab_Flipper_Button_Left.X = PinCab_Flipper_Button_Left.X - 10
		FlipperDeActivate LeftFlipper, LFPress
	End If

	If keycode = RightFlipperKey Then
		bRFHeld = False
		If VRMode = True Then PinCab_Flipper_Button_Right.X = PinCab_Flipper_Button_Right.X + 10
		FlipperDeActivate RightFlipper, RFPress
	End If

    If hsbModeActive Then
        Exit Sub
    End If

    ' Table specific

    If bGameInPLay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then
			FlipperDeActivate LeftFlipper, LFPress
            SolLFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
				StopDMDInstantInfo
            End If
        End If
        If keycode = RightFlipperKey Then
			FlipperDeActivate RightFlipper, RFPress
            SolRFlipper 0
            InstantInfoTimer.Enabled = False
            If bInstantInfo Then
                DMDScoreNow
                bInstantInfo = False
				StopDMDInstantInfo
            End If
        End If
    End If
End Sub

Sub InstantInfoTimer_Timer

    InstantInfoTimer.Enabled = False
    If NOT hsbModeActive Then
        bInstantInfo = True
        DMDFlush
        InstantInfo
		StartDMDInstanInfo
    End If
End Sub
Sub InstantInfo
    DMD CL("INSTANT INFO"), "", "", eNone, eNone, eNone, 800, False, ""
    DMD CL("JACKPOT VALUE"), CL(Jackpot(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("SPINNER VALUE"), CL(spinnervalue(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("BUMPER VALUE"), CL( bumpervalue(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("BONUS X"), CL(BonusMultiplier(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("PLAYFIELD X"), CL(PlayfieldMultiplier(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("LOCKED BALLS"), CL(BallsInLock(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("LANE BONUS"), CL(LaneBonus), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("TARGET BONUS"), CL(TargetBonus), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("RAMP BONUS"), CL(RampBonus), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("BURRITOS"), CL(MonstersKilled(CurrentPlayer)), "", eNone, eNone, eNone, 800, False, ""
    DMD CL("HIGHEST SCORE"), CL(HighScoreName(0) & " " & HighScore(0)), "", eNone, eNone, eNone, 800, False, ""
End Sub


Sub StartDMDInstanInfo
	pDMDLabelSetColorGradient "InstantInfo2a", cOrange, cRed
	pDMDLabelSetColorGradient "InstantInfo2b", cOrange, cRed
	bSupressModeMessages = True

	IIDMD_Count = 0
	'IIDMDTimer.Enabled = 1
	IIDMDTimer_Timer
End Sub

Sub StopDMDInstantInfo
	IIDMDTimer.Enabled = 0
	bSupressModeMessages = False
	PuPlayer.LabelSet pDMD,"InstantInfo2a","",0,"{'mt':2,'color': " & cWhite &" }"
	PuPlayer.LabelSet pDMD,"InstantInfo2b","",0,"{'mt':2,'color': " & cWhite &" }"  	
	DMDQueue.Add "UpdateModeProgress False","UpdateModeProgress False",45,750,0,0,0,False
End Sub

Dim IIDMD_Count

Sub IIDMDTimer_Timer
	IIDMDTimer.Enabled = 1
	Select Case IIDMD_Count
		Case 0
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," INSTANT INFO ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b","",1,""
		Case 1
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," JACKPOT VALUE ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",FormatScoreDMD(Jackpot(CurrentPlayer)),1,""  
		Case 2
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," SPINNER VALUE ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",FormatScoreDMD(spinnervalue(CurrentPlayer)),1,""  
		Case 3
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," BUMPER VALUE ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",FormatScoreDMD(bumpervalue(CurrentPlayer)),1,""  
		Case 4
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," BONUS MULTIPLIER " ,1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",BonusMultiplier(CurrentPlayer)&"X",1,""  
		Case 5
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," PLAYFIELD MULTIPLIER " ,1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",PlayfieldMultiplier(CurrentPlayer) &"X",1,""  
		Case 6
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," LOCKED BALLS ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",BallsInLock(CurrentPlayer),1,""  
		Case 7
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," LANE BONUS ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",FormatScoreDMD(LaneBonus),1,""
		Case 8
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," TARGET BONUS ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",FormatScoreDMD(TargetBonus),1,""
		Case 9
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," RAMP BONUS ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",FormatScoreDMD(RampBonus),1,""
		Case 10
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," BURRITOS ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",MonstersKilled(CurrentPlayer),1,""
		Case 11
			PuPlayer.LabelSet pDMD,"InstantInfo2a"," HIGHEST SCORE ",1,""
			PuPlayer.LabelSet pDMD,"InstantInfo2b",HighScoreName(0) & " " & FormatScoreDMD(HighScore(0)),1,""

	End Select

	IIDMD_Count = IIDMD_Count + 1
	if IIDMD_Count > 11 Then IIDMD_Count = 0:StopDMDInstantInfo
End Sub


'*************
' Pause Table
'*************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub Table1_Exit
    Savehs
    If UseFlexDMD Then FlexDMD.Run = False
    If B2SOn = true Then Controller.Stop
'    Controller.Games("THE BLIZZARD OF OZ").Settings.Value("sound") = 1
End Sub

'********************
'     Flippers
'********************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
		LF.Fire
		If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then 
			RandomSoundReflipUpLeft LeftFlipper
			DOF 101, DOFOn
		Else 
			SoundFlipperUpAttackLeft LeftFlipper
			RandomSoundFlipperUpLeft LeftFlipper
			DOF 101, DOFOn
		End If		
	Else
		LeftFlipper.RotateToStart
		If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
			RandomSoundFlipperDownLeft LeftFlipper
			DOF 101, DOFOff
		End If
		FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub


Sub SolRFlipper(Enabled)
    If Enabled Then
		RF.Fire
		If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
			RandomSoundReflipUpRight RightFlipper
			DOF 102, DOFOn
		Else 
			SoundFlipperUpAttackRight RightFlipper
			RandomSoundFlipperUpRight RightFlipper
			DOF 102, DOFOn
		End If
	Else
		RightFlipper.RotateToStart
		If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
			RandomSoundFlipperDownRight RightFlipper
			DOF 102, DOFOff
		End If	
		FlipperRightHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub LeftFlipper_Collide(parm)
	CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
	LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
	CheckLiveCatch Activeball, RightFlipper, RFCount, parm
	RightFlipperCollide parm
End Sub


'*********
' TILT
'*********

'NOTE: The TiltDecreaseTimer Subtracts .01 from the "Tilt" variable every round

Sub CheckTilt                                    'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt > TiltSensitivity)AND(Tilt < 15)Then 'show a warning
        DMD "_", CL("CAREFUL"), "_", eNone, eBlinkFast, eNone, 500, True, ""
    PlayCareful
    End if
    If Tilt > 15 Then 'If more that 15 then TILT the table
        Tilted = True
        'display Tilt
		DOF 159, DOFPulse
        DMDFlush
        DMD "", CL("TILT"), "", eNone, eNone, eNone, 200, False, ""
        PuPlayer.playevent pDMDVideo,"GameOver","Tilt.mp4",nPupVideoVolume,67,3,0,""
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'start the Tilt delay to check for all the balls to be drained
        PlayYouTilted
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt > 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        'turn off GI and turn off all the lights
        GiOff
        LightSeqTilt.Play SeqAllOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        'Bumper1.Force = 0

        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'turn back on GI and the lights
        GiOn
        LightSeqTilt.StopPlay
        'Bumper1.Force = 6
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        'clean up the buffer display
        DMDFlush
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' if all the balls have been drained then..
    If(BallsOnPlayfield = 0)Then
        ' do the normal end of ball thing (this doesn't give a bonus if the table is tilted)
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
' else retry (checks again in another second or so)
End Sub

'*********************************************
'*     Random Music Mod    *
'*********************************************
sub MusicUp
'debug.print "WTF"

   nSong = nSong + 1 
   If nSong > 56 then nSong = 0
	MusicOn
End Sub

sub MusicDown
   nSong = nSong - 1 
   If nSong < 0 then nSong = 56
	MusicOn
End Sub

Sub MusicOn
    if bAttract Then StopSound "Mu_End"
   'Dim musicEnd
if nSong < 0 or nSong > 56 then nSong = 0

	Select Case nSong
		Case 0:PlayMusic "BLIZZARD/01 - I Don_t Wanna Stop.mp3", fMusicVolume : SongTitle.ImageA = "S1"
		Case 1:PlayMusic "BLIZZARD/02 - Perry Mason.mp3", fMusicVolume :SongTitle.ImageA = "S2"
		Case 2:PlayMusic "BLIZZARD/03 - I Can_t Save You.mp3", fMusicVolume :SongTitle.ImageA = "S3"
		Case 3:PlayMusic "BLIZZARD/04 - Dreamer.mp3", fMusicVolume :SongTitle.ImageA = "S4"
		Case 4:PlayMusic "BLIZZARD/05 - Thunder Underground.mp3", fMusicVolume :SongTitle.ImageA = "S5"
		Case 5:PlayMusic "BLIZZARD/06 - Not Going Away.mp3", fMusicVolume :SongTitle.ImageA = "S6"
		Case 6:PlayMusic "BLIZZARD/07 - Mama I_m Coming Home.mp3", fMusicVolume :SongTitle.ImageA = "S7"
		Case 7:PlayMusic "BLIZZARD/08 - I Just Want You.mp3", fMusicVolume :SongTitle.ImageA = "S8"
		Case 8:PlayMusic "BLIZZARD/09 - No Easy Way Out.mp3", fMusicVolume :SongTitle.ImageA = "S9"
		Case 9:PlayMusic "BLIZZARD/10 - No More Tears.mp3" , fMusicVolume :SongTitle.ImageA = "S10"
		Case 10:PlayMusic "BLIZZARD/11 - Back On Earth.mp3", fMusicVolume :SongTitle.ImageA = "S11"
		Case 11:PlayMusic "BLIZZARD/12 - 21St Century Schizoid Man.mp3", fMusicVolume :SongTitle.ImageA = "S12"
		Case 12:PlayMusic "BLIZZARD/13 - Walk On Water.mp3", fMusicVolume :SongTitle.ImageA = "S13"
		Case 13:PlayMusic "BLIZZARD/14 - In My Life.mp3", fMusicVolume :SongTitle.ImageA = "S14"
		Case 14:PlayMusic "BLIZZARD/15 - Ghost Behind My Eyes.mp3", fMusicVolume :SongTitle.ImageA = "S15"
		Case 15:PlayMusic "BLIZZARD/16 - I Don_t Want To Change The World.mp3", fMusicVolume :SongTitle.ImageA = "S16"
		Case 16:PlayMusic "BLIZZARD/17 - Time After Time.mp3", fMusicVolume :SongTitle.ImageA = "S17"
		Case 17:PlayMusic "BLIZZARD/18 - Nightmare.mp3", fMusicVolume :SongTitle.ImageA = "S18"
		Case 18:PlayMusic "BLIZZARD/19 - Gets Me Through.mp3", fMusicVolume :SongTitle.ImageA = "S19"
		Case 19:PlayMusic "BLIZZARD/20 - See You On The Other Side.mp3", fMusicVolume :SongTitle.ImageA = "S20"
		Case 20:PlayMusic "BLIZZARD/21 - Mississippi Queen.mp3", fMusicVolume :SongTitle.ImageA = "S21"
		Case 21:PlayMusic "BLIZZARD/22 - Mr Tinkertrain.mp3", fMusicVolume :SongTitle.ImageA = "S22"
		Case 22:PlayMusic "BLIZZARD/23 - Shot In The Dark.mp3", fMusicVolume :SongTitle.ImageA = "S23"
		Case 23:PlayMusic "BLIZZARD/24 - Breaking All The Rules.mp3", fMusicVolume :SongTitle.ImageA = "S24"
		Case 24:PlayMusic "BLIZZARD/25 - I Don_t Know.mp3", fMusicVolume :SongTitle.ImageA = "S25"
		Case 25:PlayMusic "BLIZZARD/26 - Bark At The Moon.mp3", fMusicVolume :SongTitle.ImageA = "S26"
		Case 26:PlayMusic "BLIZZARD/27 - Over The Mountain.mp3", fMusicVolume :SongTitle.ImageA = "S27"
		Case 27:PlayMusic "BLIZZARD/28 - Crazy Train.mp3", fMusicVolume :SongTitle.ImageA = "S28"
		Case 28:PlayMusic "BLIZZARD/29 - Miracle Man.mp3", fMusicVolume :SongTitle.ImageA = "S29"
		Case 29:PlayMusic "BLIZZARD/30 - Flying High Again.mp3", fMusicVolume :SongTitle.ImageA = "S30"
		Case 30:PlayMusic "BLIZZARD/31 - Mr Crowley.mp3", fMusicVolume :SongTitle.ImageA = "S31"
		Case 31:PlayMusic "BLIZZARD/32 - Crazy Babies.mp3", fMusicVolume :SongTitle.ImageA = "S32"
		Case 32:PlayMusic "BLIZZARD/33 - Diary Of A Madman.mp3", fMusicVolume :SongTitle.ImageA = "S33"
		Case 33:PlayMusic "BLIZZARD/34 - Am I Going Insane.mp3", fMusicVolume :SongTitle.ImageA = "S34"
		Case 34:PlayMusic "BLIZZARD/35 - Black Sabbath.mp3", fMusicVolume :SongTitle.ImageA = "S35"
		Case 35:PlayMusic "BLIZZARD/36 - Changes.mp3", fMusicVolume :SongTitle.ImageA = "S36"
		Case 36:PlayMusic "BLIZZARD/37 - Children of the Grave.mp3", fMusicVolume :SongTitle.ImageA = "S37"
		Case 37:PlayMusic "BLIZZARD/38 - Electric Funeral.mp3", fMusicVolume :SongTitle.ImageA = "S38"
		Case 38:PlayMusic "BLIZZARD/39 - Fairies Wear Boots.mp3", fMusicVolume :SongTitle.ImageA = "S39"
		Case 39:PlayMusic "BLIZZARD/40 - Hand of Doom.mp3", fMusicVolume :SongTitle.ImageA = "S40"
		Case 40:PlayMusic "BLIZZARD/41 - Iron Man.mp3", fMusicVolume :SongTitle.ImageA = "S41"
		Case 41:PlayMusic "BLIZZARD/42 - Megalomania.mp3", fMusicVolume :SongTitle.ImageA = "S42"
		Case 42:PlayMusic "BLIZZARD/43 - N I B.mp3", fMusicVolume :SongTitle.ImageA = "S43"
		Case 43:PlayMusic "BLIZZARD/44 - Paranoid.mp3", fMusicVolume :SongTitle.ImageA = "S44"
		Case 44:PlayMusic "BLIZZARD/45 - Rat Salad.mp3", fMusicVolume :SongTitle.ImageA = "S45"
		Case 45:PlayMusic "BLIZZARD/46 - Sabbath Bloody Sabbath.mp3", fMusicVolume :SongTitle.ImageA = "S46"
		Case 46:PlayMusic "BLIZZARD/47 - Snowblind.mp3", fMusicVolume :SongTitle.ImageA = "S47"
		Case 47:PlayMusic "BLIZZARD/48 - Sweet Leaf.mp3", fMusicVolume:SongTitle.ImageA = "S48"
		Case 48:PlayMusic "BLIZZARD/49 - The Wizard.mp3", fMusicVolume :SongTitle.ImageA = "S49"
		Case 49:PlayMusic "BLIZZARD/50 - Tomorrow_s Dream.mp3", fMusicVolume :SongTitle.ImageA = "S50"
		Case 50:PlayMusic "BLIZZARD/51 - War Pigs.mp3", fMusicVolume :SongTitle.ImageA = "S51"
		Case 51:PlayMusic "BLIZZARD/52 - Under The Graveyard.mp3", fMusicVolume :SongTitle.ImageA = "S52"
		Case 52:PlayMusic "BLIZZARD/53 - Take What You Want From Me.mp3", fMusicVolume :SongTitle.ImageA = "S53"
		Case 53:PlayMusic "BLIZZARD/54 - PatientNo9.mp3", fMusicVolume :SongTitle.ImageA = "S54"
		Case 54:PlayMusic "BLIZZARD/55 - Immortal.mp3", fMusicVolume :SongTitle.ImageA = "S55"
		Case 55:PlayMusic "BLIZZARD/56 - Degredation Rules.mp3", fMusicVolume :SongTitle.ImageA = "S56"
		Case 56:PlayMusic "BLIZZARD/57 - PlanetCaravan.mp3", fMusicVolume :SongTitle.ImageA = "S57"
	End Select	

	if bAttract Then
		Select Case nSong
			Case 0: TrackName = "01 - I Don_t Wanna Stop.mp3" : TrackFontSize = 9
			Case 1: TrackName = "02 - Perry Mason.mp3" : TrackFontSize = 10
			Case 2: TrackName = "03 - I Can_t Save You.mp3" : TrackFontSize = 9
			Case 3: TrackName = "04 - Dreamer.mp3" : TrackFontSize = 10
			Case 4: TrackName = "05 - Thunder Underground.mp3" : TrackFontSize = 8
			Case 5: TrackName = "06 - Not Going Away.mp3" : TrackFontSize = 10
			Case 6: TrackName = "07 - Mama I_m Coming Home.mp3" : TrackFontSize = 8
			Case 7: TrackName = "08 - I Just Want You.mp3" : TrackFontSize = 10
			Case 8: TrackName = "09 - No Easy Way Out.mp3" : TrackFontSize = 10
			Case 9: TrackName = "10 - No More Tears.mp3" : TrackFontSize = 10 
			Case 10: TrackName = "11 - Back On Earth.mp3" : TrackFontSize = 10
			Case 11: TrackName = "12 - 21St Century Schizoid Man.mp3" : TrackFontSize = 7
			Case 12: TrackName = "13 - Walk On Water.mp3" : TrackFontSize = 10
			Case 13: TrackName = "14 - In My Life.mp3" : TrackFontSize = 10
			Case 14: TrackName = "15 - Ghost Behind My Eyes.mp3" : TrackFontSize = 8
			Case 15: TrackName = "16 - I Don_t Want To Change The World.mp3" : TrackFontSize = 6
			Case 16: TrackName = "17 - Time After Time.mp3" : TrackFontSize = 9
			Case 17: TrackName = "18 - Nightmare.mp3" : TrackFontSize = 10
			Case 18: TrackName = "19 - Gets Me Through.mp3" : TrackFontSize = 9
			Case 19: TrackName = "20 - See You On The Other Side.mp3" : TrackFontSize = 7
			Case 20: TrackName = "21 - Mississippi Queen.mp3" : TrackFontSize = 8.5
			Case 21: TrackName = "22 - Mr Tinkertrain.mp3" : TrackFontSize = 9
			Case 22: TrackName = "23 - Shot In The Dark.mp3" : TrackFontSize = 9
			Case 23: TrackName = "24 - Breaking All The Rules.mp3" : TrackFontSize = 8
			Case 24: TrackName = "25 - I Don_t Know.mp3" : TrackFontSize = 10
			Case 25: TrackName = "26 - Bark At The Moon.mp3" : TrackFontSize = 9
			Case 26: TrackName = "27 - Over The Mountain.mp3" : TrackFontSize = 8.5
			Case 27: TrackName = "28 - Crazy Train.mp3" : TrackFontSize = 10
			Case 28: TrackName = "29 - Miracle Man.mp3" : TrackFontSize = 10
			Case 29: TrackName = "30 - Flying High Again.mp3" : TrackFontSize = 8.5
			Case 30: TrackName = "31 - Mr Crowley.mp3" : TrackFontSize = 10
			Case 31: TrackName = "32 - Crazy Babies.mp3" : TrackFontSize = 10
			Case 32: TrackName = "33 - Diary Of A Madman.mp3" : TrackFontSize = 8.5
			Case 33: TrackName = "34 - Am I Going Insane.mp3" : TrackFontSize = 8.5
			Case 34: TrackName = "35 - Black Sabbath.mp3" : TrackFontSize = 10
			Case 35: TrackName = "36 - Changes.mp3" : TrackFontSize = 10
			Case 36: TrackName = "37 - Children of the Grave.mp3" : TrackFontSize = 7.5
			Case 37: TrackName = "38 - Electric Funeral.mp3" : TrackFontSize = 10
			Case 38: TrackName = "39 - Fairies Wear Boots.mp3" : TrackFontSize = 8
			Case 39: TrackName = "40 - Hand of Doom.mp3" : TrackFontSize = 10
			Case 40: TrackName = "41 - Iron Man.mp3" : TrackFontSize = 10
			Case 41: TrackName = "42 - Megalomania.mp3" : TrackFontSize = 10
			Case 42: TrackName = "43 - N I B.mp3" : TrackFontSize = 10
			Case 43: TrackName = "44 - Paranoid.mp3" : TrackFontSize = 10
			Case 44: TrackName = "45 - Rat Salad.mp3" : TrackFontSize = 10
			Case 45: TrackName = "46 - Sabbath Bloody Sabbath.mp3" : TrackFontSize = 7.5
			Case 46: TrackName = "47 - Snowblind.mp3" : TrackFontSize = 10
			Case 47: TrackName = "48 - Sweet Leaf.mp3" : TrackFontSize = 10
			Case 48: TrackName = "49 - The Wizard.mp3" : TrackFontSize = 10
			Case 49: TrackName = "50 - Tomorrow_s Dream.mp3" : TrackFontSize = 8
			Case 50: TrackName = "51 - War Pigs.mp3" : TrackFontSize = 10
			Case 51: TrackName = "52 - Under The Graveyard.mp3" : TrackFontSize = 8
			Case 52: TrackName = "53 - Take What You Want From Me.mp3" : TrackFontSize = 7
			Case 53: TrackName = "54 - PatientNo9.mp3" : TrackFontSize = 10
			Case 54: TrackName = "55 - Immortal.mp3" : TrackFontSize = 10
			Case 55: TrackName = "56 - Degredation Rules.mp3" : TrackFontSize = 9
			Case 56: TrackName = "57 - PlanetCaravan.mp3" : TrackFontSize = 10
		End Select	


		PuPlayer.LabelSet pDMD,"JukeBox2a"," JUKEBOX MODE ",1,"{'mt':2,'color': " & cWhite &", 'xpos': 50, 'ypos': 82.5, 'fonth':10}"
		PuPlayer.LabelSet pDMD,"JukeBox2b",TrackName,1,"{'mt':2,'color': " & cOrange &", 'xpos': 50, 'ypos': 91, 'fonth':" & TrackFontSize &" }"
		pDMDLabelSetColorGradient "JukeBox2a",  cYellow, cRed
		pDMDLabelSetColorGradientPercent "JukeBox2b",  cBlue, cPurple, 40
	End If

End Sub

Sub Table1_MusicDone()
	if bAttract Then Exit Sub
    MusicUp
    
End Sub



'********************
' Music as wav sounds
'********************

Dim Song
Song = ""

Sub PlaySong(name)
    If bMusicOn Then
        If Song <> name Then
            StopSound Song
            Song = name
            PlaySound Song, -1, IntroSongVolume
        End If
    End If
End Sub

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
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = 1 Then 'we have 2 captive balls on the table (-1 means no balls, 0 is the first ball, 1 is the second..)
            GiOff               ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
	if DMDType = 5 Then Exit Sub 
    DOF 118, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1	
    Next
    For each bulb in aBumperLights
        bulb.State = 1
    Next

	For each bulb in GIBulbs: bulb.blenddisablelighting = 1.25: Next

End Sub

Sub GiOff
    DOF 118, DOFOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    For each bulb in aBumperLights
        bulb.State = 0
    Next

	For each bulb in GIBulbs: bulb.blenddisablelighting = 0: Next

End Sub

' GI, light & flashers sequence effects

Sub GiEffect(n)
    Dim ii
    Select Case n
        Case 0 'all off
            LightSeqGi.Play SeqAlloff
        Case 1 'all blink
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqGi.UpdateInterval = 10
            LightSeqGi.Play SeqBlinking, , 10, 10
    End Select
End Sub

Sub LightEffect(n)
    Select Case n
        Case 0 ' all off
            LightSeqInserts.Play SeqAlloff
        Case 1 'all blink
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 15, 10
        Case 2 'random
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqInserts.UpdateInterval = 10
            LightSeqInserts.Play SeqBlinking, , 10, 10
    End Select
End Sub

Sub FlashEffect(n)
    Dim ii
    Select case n
        Case 0 ' all off
            LightSeqFlasher.Play SeqAlloff
        Case 1 'all blink
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqBlinking, , 10, 10
        Case 2 'random
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqRandom, 50, , 1000
        Case 3 'all blink fast
            LightSeqFlasher.UpdateInterval = 10
            LightSeqFlasher.Play SeqBlinking, , 5, 10
    End Select
End Sub

'Dim TableWidth, TableHeight

'TableWidth = Table1.width
'TableHeight = Table1.height

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v4.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 2     'number of locked balls
Const maxvel = 40 'max ball velocity

' *********************************************************************
'                        User Defined Script Events
' *********************************************************************

' Initialise the Table for a new Game
'
Sub ResetForNewGame()
    Dim i

    bGameInPLay = True
	ClearTwoLines
	ClearPupAttractMessages

	ResetTrackers

    'resets the score display, and turn off attract mode
    StopAttractMode
	vpmtimer.addtimer 1700, "StopAttractMode '"	' make sure it sotpped
    GiOn

	LoadBG
	Minion1wallDown.Enabled = 1
	Minion2wallDown.Enabled = 1

	DMDQueue.Add "pdmdsetpage 1","pdmdsetpage 1", 50, 100, 0,0,0,False 
	

	If DMDType = 4 Then
		PuPlayer.playevent pTopper,"Topper","Topper.mp4",0,20,6,0,""
	End IF

    TotalGamesPlayed = TotalGamesPlayed + 1
    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        BonusPoints(i) = 0
        BonusHeldPoints(i) = 0
        BonusMultiplier(i) = 1
        PlayfieldMultiplier(i) = 1
        BallsRemaining(i) = BallsPerGame
        ExtraBallsAwards(i) = 0
    Next

	bWizMode1Active = False
	bWizMode2Active = False
	bWizMode3Active = False
	bFinalWizModeActive = False

	'bGameReady = False
	bEOB = False


    ' initialise any other flags
    Tilt = 0

    ' initialise Game variables
    Game_Init()
    MusicOn


	if ScorbitActive = 1 And (Scorbit.bNeedsPairing) = False Then 
		Scorbit.StartSession()
		Dbg2 "Starting Scorbit Session"
		GameModeStrTmp="NA{Yellow}:Starting Game"
		if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
	End If
	bOnTheFirstBallScorbit = True

	DMDUpdateAll
    ' you may wish to start some music, play a sound, do whatever at this point
	bBallSaverReady = True
    vpmtimer.addtimer 1500, "FirstBall '"


End Sub

' This is used to delay the start of a game to allow any attract sequence to
' complete.  When it expires it creates a ball for the player to start playing with

Sub FirstBall
    ' reset the table for a new ball
    ResetForNewPlayerBall()
    ' create a new ball in the shooters lane
    CreateNewBall()
End Sub

' (Re-)Initialise the Table for a new ball (either a new ball after the player has
' lost one or we have moved onto the next player (if multiple are playing))

Sub ResetForNewPlayerBall()
    ' make sure the correct display is upto date
    AddScore 0
    'This is a new ball, so activate the ballsaver
    bBallSaverReady = True

	' Reset EOB flag
	bEOB = False

    ' set the current players bonus multiplier back down to 1X
    SetBonusMultiplier 1


	IconStep = 0
	IconTimer.Enabled = True

	DbgTracker ""	
	DbgTracker "PLAYER " & CurrentPlayer & " NEXT BALL: " &balls

    ' reduce the playfield multiplier
    SetPlayfieldMultiplier 1

    ' reset any drop targets, lights, game Mode etc..

    BonusPoints(CurrentPlayer) = 0
    bBonusHeld = False
    bExtraBallWonThisBall = False
	ComboCount = 0

	UpdateAlbumCount
	UpdateComboCount

    'Reset any table specific
    ResetNewBallVariables
    ResetNewBallLights()

	nWizardModeMultiplier = 1



    'and the skillshot
    bSkillShotReady = True

	DMDUpdateAll

'Change the music ?
End Sub

' Create a new ball on the Playfield

Sub CreateNewBall()
    ' create a ball in the plunger lane kicker.
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    ' There is a (or another) ball on the playfield
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' kick it out..
    PlaySoundAt SoundFXDOF("fx_Ballrel", 123, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4

' if there is 2 or more balls then set the multibal flag (remember to check for locked balls and other balls used for animations)
' set the bAutoPlunger flag to kick the ball in play automatically
    If BallsOnPlayfield > 1 Then
        DOF 143, DOFPulse
        bMultiBallMode = True
        bAutoPlunger = True
        ChangeGi 5
    End If

	if Scorbit.bSessionActive then
		GameModeStrTmp="NA{Red}:Ball: " &balls
		if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
	End If
End Sub

' Add extra balls to the table with autoplunger
' Use it as AddMultiball 4 to add 4 extra balls to the table

Sub AddMultiball(nballs)
    mBalls2Eject = mBalls2Eject + nballs
    CreateMultiballTimer.Enabled = True
    'and eject the first ball
    CreateMultiballTimer_Timer
End Sub

' Eject the ball after the delay, AddMultiballDelay
Sub CreateMultiballTimer_Timer()
    ' wait if there is a ball in the plunger lane
    If bBallInPlungerLane Then
        Exit Sub
    Else
        If BallsOnPlayfield < MaxMultiballs Then
            CreateNewBall()
            mBalls2Eject = mBalls2Eject -1
            If mBalls2Eject = 0 Then 'if there are no more balls to eject then stop the timer
                CreateMultiballTimer.Enabled = False
            End If
        Else 'the max number of multiballs is reached, so stop the timer
            mBalls2Eject = 0
            CreateMultiballTimer.Enabled = False
        End If
    End If
End Sub


Sub ResetTrackers
	Tracker_Lanes = 0
	Tracker_Ramps = 0
	Tracker_Orbits = 0
	Tracker_Targets = 0
	Tracker_Bumpers = 0
	Tracker_Spinners = 0
	Tracker_Burritos = 0
	Tracker_ComboHits = 0
	Tracker_BonusPoints = 0
	Tracker_ComboValue = 0
End Sub

' The Player has lost his ball (there are no more balls on the playfield).
' Handle any bonus points awarded
	Dim Msg1
Sub DisplayEOB1
	Dim AP
	AP = SpinCount * BONUS_SPINNERS
	if DMDType = 5 Then
		Msg1 = " SPINNER BONUS = " & FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5A",msg1,1,"{'mt':2,'fonth': "&EOBFontSize&"}"	
	Else
		Msg1 = " SPINNER BONUS = " & FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5A",msg1,1,""	
	End If
 
End Sub

	Dim Msg2
Sub DisplayEOB2
	Dim AP
	AP = TargetBonus * BONUS_TARGETS
	if DMDType = 5 Then
		Msg2 = " TARGET BONUS = " & FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5B",msg2,1,"{'mt':2,'fonth': "&EOBFontSize&"}"
	Else
		Msg2 = " TARGET BONUS = " & FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5B",msg2,1,"" 
	End If

End Sub

	Dim Msg3
Sub DisplayEOB3
	Dim AP
	AP = MonstersKilled(CurrentPlayer) * BONUS_BURRITOS
	if DMDType = 3 Then
		Msg3 = " BURRITOS = " & FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5C",msg3,1,"{'mt':2,'fonth': "&EOBFontSize&"}" 
	Else
		Msg3 = " BURRITOS = " & FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5C",msg3,1,"" 
	End If

End Sub

	Dim Msg4
Sub DisplayEOB4
	Dim AP
	AP = RampBonus * BONUS_RAMPS
	if DMDType = 5 Then
		Msg4 = " RAMP BONUS = " & FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5D",msg4,1,"{'mt':2,'fonth': "&EOBFontSize&"}" 
	Else
		Msg4 = " RAMP BONUS = " & FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5D",msg4,1,"" 
	End If

End Sub



	Dim Msg5
Sub DisplayEOB5
	Dim AP
	AP = BattlesWon(CurrentPlayer) * BONUS_MODESCOMPLETED
	if DMDType = 5 Then
		Msg5 = " ALBUMS COLLECTED = " &FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5E",msg5,1,"{'mt':2,'fonth': "&EOBFontSize&"}" 
	Else
		Msg5 = " ALBUMS COLLECTED = " &FormatScoreDMD(AP)
		PuPlayer.LabelSet pDMD,"Event5E",msg5,1,"" 
	End If

End Sub


	Dim Msg6
Sub DisplayEOB6
	Dim TB
	TB = TotalBonus 
	if DMDType = 5 Then
		Msg6 = "TOTAL BONUS = " &FormatScoreDMD(TB)
		PuPlayer.LabelSet pDMD,"Event5F",msg6,1,"{'mt':2,'fonth': "&EOBFontSize&"}"
	Else
		Msg6 = "TOTAL BONUS = " &FormatScoreDMD(TB)
		PuPlayer.LabelSet pDMD,"Event5F",msg6,1,""
	End If

End Sub


Sub ClearEOBDMD

	PuPlayer.LabelSet pDMD,"Event5A","",0,""
	PuPlayer.LabelSet pDMD,"Event5B","",0,""
	PuPlayer.LabelSet pDMD,"Event5C","",0,""
	PuPlayer.LabelSet pDMD,"Event5D","",0,""
	PuPlayer.LabelSet pDMD,"Event5E","",0,""
	PuPlayer.LabelSet pDMD,"Event5F","",0,""
End Sub

Dim TotalBonus
Sub EndOfBall()

	DOF 146, DOFPulse

    Dim AwardPoints, ii
    AwardPoints = 0
    TotalBonus = 0
    StopSpinner
    CloseDoor
    PlayBallLost
	ballhandlingQueue.Add "hideblizztimer","hideblizztimer",65,250,0,0,0,False	
    ResetCountdown
    StopFire
    StopMinionMode
	if Scorbit.bSessionActive then
		GameModeStrTmp="NA{Red}:Ball Lost: " &balls
		if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
	End If

    ' the first ball has been lost. From this point on no new players can join in
    bOnTheFirstBall = False

    ' only process any of this if the table is not tilted.  (the tilt recovery
    ' mechanism will handle any extra balls or end of game)

        ' handle the bonus held
        ' reset the bonus held value since it has been already added to the bonus
        BonusHeldPoints(CurrentPlayer) = 0
	bEOB = True

    If NOT Tilted Then

' TRACKER CODE FOR SCORING



        ' calculate the totalbonus
		TotalBonus = (SpinCount*BONUS_SPINNERS) + (TargetBonus*BONUS_TARGETS) + (MonstersKilled(CurrentPlayer)*BONUS_BURRITOS)+ (RampBonus*BONUS_RAMPS) + (BattlesWon(CurrentPlayer)*BONUS_MODESCOMPLETED)


        TotalBonus = (TotalBonus * BonusMultiplier(CurrentPlayer)) + BonusHeldPoints(CurrentPlayer)



        ' the player has won the bonus held award so do something with it :)
        If bBonusHeld Then
            If Balls = BallsPerGame Then ' this is the last ball, so if bonus held has been awarded then double the bonus
                TotalBonus = TotalBonus * 2
            End If
        Else ' this is not the last ball so save the bonus for the next ball
            BonusHeldPoints(CurrentPlayer) = TotalBonus
        End If
        bBonusHeld = False






		EOBQueue.Add "Line1", "DisplayEOB1", 50, 10, 0,0,0,False 
'		EOBQueue.Add "Line6a", "DisplayEOB6 1 ", 150, 400, 0,0,0,False 

		EOBQueue.Add "Line2", "DisplayEOB2", 50, 800, 0,0,0,False 
'		EOBQueue.Add "Line6b", "DisplayEOB6 2", 50, 1400, 0,0,0,False 

		EOBQueue.Add "Line3", "DisplayEOB3", 50, 1600, 0,0,0,False 
'		EOBQueue.Add "Line6c", "DisplayEOB6 3", 50, 2400, 0,0,0,False 


		EOBQueue.Add "Line4", "DisplayEOB4", 50, 2400, 0,0,0,False 
'		EOBQueue.Add "Line6d", "DisplayEOB6 4", 50, 3400, 0,0,0,False 

		EOBQueue.Add "Line5", "DisplayEOB5", 50, 3200, 0,0,0,False 
'		EOBQueue.Add "Line6e", "DisplayEOB6 5", 50, 4400, 0,0,0,False 

		EOBQueue.Add "Line6", "DisplayEOB6", 50, 4200, 0,0,0,False 
'		EOBQueue.Add "Line6f", "DisplayEOB6 6", 50, 5400, 0,0,0,False 


		' add a bit of a delay to allow for the bonus points to be shown & added up
		EOBQueue.Add "Line7","ClearEOBDMD", 50, 5600, 0,0,0,False 
		'EOBQueue.Add "Line8","EndOfBall2", 50, 5100, 0,0,0,False 

'add in any bonus points (multipled by the bonus multiplier)
'AwardPoints = BonusPoints(CurrentPlayer) * BonusMultiplier(CurrentPlayer)
'AddScore AwardPoints
'debug.print "Bonus Points = " & AwardPoints
'DMD "", CL("BONUS: " & BonusPoints(CurrentPlayer) & " X" & BonusMultiplier(CurrentPlayer) ), "", eNone, eBlink, eNone, 1000, True, ""

'Count the bonus. This table uses several bonus

		'Number of Spinners hit
        AwardPoints = SpinCount * BONUS_SPINNERS
        'TotalBonus = AwardPoints
        DMD CL(FormatScore(AwardPoints)), CL("SPINNER BONUS " & LaneBonus), "", eBlink, eNone, eNone, 800, False, ""

        'Number of Target hits
        AwardPoints = TargetBonus * BONUS_TARGETS
        'TotalBonus = TotalBonus + AwardPoints
        DMD CL(FormatScore(AwardPoints)), CL("TARGET BONUS " & TargetBonus), "", eBlink, eNone, eNone, 800, False, ""
       
        'Number of Burritos Collected
        AwardPoints = MonstersKilled(CurrentPlayer) * BONUS_BURRITOS
        'TotalBonus = TotalBonus + AwardPoints
        DMD CL(FormatScore(AwardPoints)), CL("BURRITOS " & MonstersKilled(CurrentPlayer)), "", eBlink, eNone, eNone, 800, False, ""

        'Number of Ramps completed
        AwardPoints = RampBonus * BONUS_RAMPS
        'TotalBonus = TotalBonus + AwardPoints
        DMD CL(FormatScore(AwardPoints)), CL("RAMP BONUS " & RampBonus), "", eBlink, eNone, eNone, 800, False, ""



        'Modes X 250.000
        AwardPoints = BattlesWon(CurrentPlayer) * BONUS_MODESCOMPLETED
        'TotalBonus = TotalBonus + AwardPoints
        DMD CL("ALBUMS COLLECTED"), CL(FormatScore(AwardPoints) ), "", eNone, eNone, eNone, 800, True, ""

        ' Add the bonus to the score
        DMD CL(FormatScore(TotalBonus)), CL("TOTAL BONUS " & " X" & BonusMultiplier(CurrentPlayer)), "", eBlinkFast, eNone, eNone, 1500, True, ""

        'AddScore TotalBonus

''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
''               ############################    TRACKING INFO   #########################################################
''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'
'		Tracker_Lanes = Tracker_Lanes + LaneBonus
'		Tracker_Ramps = Tracker_Ramps + RampBonus
'		Tracker_Orbits = Tracker_Orbits + OrbitHits
'		Tracker_Targets = Tracker_Targets + TargetBonus
'		Tracker_Bumpers = Tracker_Bumpers + BumperHits
'		Tracker_Spinners = Tracker_Spinners + SpinCount
'		Tracker_Burritos = Tracker_Burritos + MonstersKilled(CurrentPlayer)
'		Tracker_ComboHits = Tracker_ComboHits + ComboHits(CurrentPlayer)
'		Tracker_BonusPoints = Tracker_BonusPoints + BonusPoints(CurrentPlayer)
'
'		DbgTracker "*********************************"
'		DbgTracker "End Of Ball " & balls
'		DbgTracker "---------------------------------"
'
'		DbgTracker "LANES: " & LaneBonus &" * " &SCORE_LANES &" = " &(LaneBonus*SCORE_LANES)
'		DbgTracker "RAMPS: " & RampBonus &" * " & BONUS_RAMPS &" = "&(RampBonus*BONUS_RAMPS)
'		DbgTracker "ORBITS: " & OrbitHits &" * " &SCORE_ORBITS &" = "&(OrbitHits*SCORE_ORBITS)
'		DbgTracker "TARGETS: " & TargetBonus &" * " & BONUS_TARGETS &" = "&(TargetBonus*BONUS_TARGETS)
'		DbgTracker "BUMPERS: " & BumperHits &" * " & BumperValue(CurrentPlayer) &" = "& (BumperHits*BumperValue(CurrentPlayer))
'		DbgTracker "SPINNERS: " & SpinCount &" * " & BONUS_SPINNERS &" = "&(SpinCount*BONUS_SPINNERS)
'		DbgTracker "BURRITOS: " & MonstersKilled(CurrentPlayer) &" * " & BONUS_BURRITOS &" = "&(MonstersKilled(CurrentPlayer)*BONUS_BURRITOS)
'		DbgTracker "COMBO HITS: " & ComboHits(CurrentPlayer) &" * " &SCORE_COMBOHITS &" = "&(Combohits(CurrentPlayer)*SCORE_COMBOHITS)
'		DbgTracker "MODES: " & BattlesWon(CurrentPlayer) &" * " & BONUS_MODESCOMPLETED &" = "&(BattlesWon(CurrentPlayer)*BONUS_MODESCOMPLETED)
'		DbgTracker "BONUS: " & BonusPoints(CurrentPlayer)
'		DbgTracker "BONUS HELD: " & BonusHeldPoints(CurrentPlayer)
'
''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
''               ############################    END OF SCORING    #########################################################
''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&



        ' add a bit of a delay to allow for the bonus points to be shown & added up
'        vpmtimer.addtimer 6000, "EndOfBall2 '"
		EOBQueue.Add "EndOfBall2", "EndOfBall2", 50, 6000, 0,0,0,False 
		
    Else 'if tilted then only add a short delay
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub


Sub ExpediteEOB
	bEOB=False
	DMDFlush
	EOBQueue.RemoveAll(True)
	ClearEOBDMD
	EndofBall2
End Sub

' The Timer which delays the machine to allow any bonus points to be added up
' has expired.  Check to see if there are any extra balls for this player.
' if not, then check to see if this was the last ball (of the CurrentPlayer)
'
Sub EndOfBall2()
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'               ############################    TRACKING INFO   #########################################################
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

		Tracker_Lanes = Tracker_Lanes + LaneBonus
		Tracker_Ramps = Tracker_Ramps + RampBonus
		Tracker_Orbits = Tracker_Orbits + OrbitHits
		Tracker_Targets = Tracker_Targets + TargetBonus
		Tracker_Bumpers = Tracker_Bumpers + BumperHits
		Tracker_Spinners = Tracker_Spinners + SpinCount
		Tracker_Burritos = Tracker_Burritos + MonstersKilled(CurrentPlayer)
		Tracker_ComboHits = Tracker_ComboHits + ComboHits(CurrentPlayer)
		'Tracker_ComboValue = Tracker_ComboValue + ComboHits(CurrentPlayer)
		Tracker_BonusPoints = Tracker_BonusPoints + BonusPoints(CurrentPlayer)

		DbgTracker "*********************************"
		DbgTracker "End Of Ball " & balls
		DbgTracker "---------------------------------"

		DbgTracker "LANES: " & LaneBonus &" * " &SCORE_LANES &" = " &(LaneBonus*SCORE_LANES)
		DbgTracker "RAMPS: " & RampBonus &" * " & BONUS_RAMPS &" = "&(RampBonus*BONUS_RAMPS)
		DbgTracker "ORBITS: " & OrbitHits &" * " &SCORE_ORBITS &" = "&(OrbitHits*SCORE_ORBITS)
		DbgTracker "TARGETS: " & TargetBonus &" * " & BONUS_TARGETS &" = "&(TargetBonus*BONUS_TARGETS)
		DbgTracker "BUMPERS: " & BumperHits &" * " & BumperValue(CurrentPlayer) &" = "& (BumperHits*BumperValue(CurrentPlayer))
		DbgTracker "SPINNERS: " & SpinCount &" * " & BONUS_SPINNERS &" = "&(SpinCount*BONUS_SPINNERS)
		DbgTracker "COMBO VALUE: " & Tracker_ComboValue
		DbgTracker "BURRITOS: " & MonstersKilled(CurrentPlayer) &" * " & BONUS_BURRITOS &" = "&(MonstersKilled(CurrentPlayer)*BONUS_BURRITOS)
		DbgTracker "COMBO HITS: " & ComboHits(CurrentPlayer) &" * " &SCORE_COMBOHITS &" = "&(Combohits(CurrentPlayer)*SCORE_COMBOHITS)
		DbgTracker "MODES: " & BattlesWon(CurrentPlayer) &" * " & BONUS_MODESCOMPLETED &" = "&(BattlesWon(CurrentPlayer)*BONUS_MODESCOMPLETED)
		DbgTracker "BONUS: " & BonusPoints(CurrentPlayer)
		DbgTracker "BONUS HELD: " & BonusHeldPoints(CurrentPlayer)

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'               ############################    END OF SCORING    #########################################################
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&


    ' if were tilted, reset the internal tilted flag (this will also
    ' set TiltWarnings back to zero) which is useful if we are changing player LOL
    Tilted = False
    Tilt = 0
    DisableTable False 'enable again bumpers and slingshots

	AddScore TotalBonus

    ' has the player won an extra-ball ? (might be multiple outstanding)
    If(ExtraBallsAwards(CurrentPlayer) <> 0)Then
        'debug.print "Extra Ball"

        ' yep got to give it to them
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' if no more EB's then turn off any shoot again light
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            LightShootAgain.State = 0
        End If

        ' You may wish to do a bit of a song AND dance at this point
        DMD CL("EXTRA BALL"), CL("SHOOT AGAIN"), "", eNone, eNone, eBlink, 1000, True, ""
        PlayShootAgainsng
        ' In this table an extra ball will have the skillshot and ball saver, so we reset the playfield for the new ball
        ResetForNewPlayerBall()

        ' Create a new ball in the shooters lane
        CreateNewBall()
    Else ' no extra balls

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

        ' was that the last ball ?
        If(BallsRemaining(CurrentPlayer) <= 0)Then
            'debug.print "No More Balls, High Score Entry"

            ' Submit the CurrentPlayers score to the High Score system
            CheckHighScore()
        ' you may wish to play some music at this point

        Else

            ' not the last ball (for that player)
            ' if multiple players are playing then move onto the next one
            EndOfBallComplete()
        End If
    End If
End Sub

' This function is called when the end of bonus display
' (or high score entry finished) AND it either end the game or
' move onto the next player (or the next ball of the same player)
'
Sub EndOfBallComplete()
    Dim NextPlayer
    
    'debug.print "EndOfBall - Complete"

    ' are there multiple players playing this game ?
    If(PlayersPlayingGame > 1)Then
        ' then move to the next player
        NextPlayer = CurrentPlayer + 1
        ' are we going from the last player back to the first
        ' (ie say from player 4 back to player 1)
        If(NextPlayer > PlayersPlayingGame)Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' is it the end of the game ? (all balls been lost for all players)
    If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then
        ' you may wish to do some sort of Point Match free game award here
        ' generally only done when not in free play mode

        ' set the machine into game over mode
        EndOfGame()

    ' you may wish to put a Game Over message on the desktop/backglass

    Else
        ' set the next player
        CurrentPlayer = NextPlayer

        ' make sure the correct display is up to date
        AddScore 0

        ' reset the playfield for the new player (or new ball)
        ResetForNewPlayerBall()
        ' AND create a new ball
        CreateNewBall()

        ' play a sound if more than 1 player
        If PlayersPlayingGame > 1 Then
            DMD "_", CL("PLAYER " &CurrentPlayer), "_", eNone, eNone, eNone, 800, True, ""
            PlayPlayerUp
        End If
    End If
End Sub

' This function is called at the End of the Game, it should reset all
' Drop targets, AND eject any 'held' balls, start any attract sequences etc..

Sub EndOfGame()
	EndMusic
    Playsound "mu_end"
    'debug.print "End Of Game"
    bGameInPLay = False
    ' just ended your game then play the end of game tune
    If NOT bJustStarted Then
    End If

	PupGameOver

    bJustStarted = False
    ' ensure that the flippers are down
    SolLFlipper 0
    SolRFlipper 0

	if Scorbit.bSessionActive then
		GameModeStrTmp="NA{Red}:Game Over"
		if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
		StopScorbit
	End If

    ' terminate all Mode - eject locked balls
    ' most of the Mode/timers terminate at the end of the ball

    ' set any lights for the attract mode
    GiOff
	GeneralPupQueue.Add "StartAttractMode","StartAttractMode",60,6000,0,0,0,False


		DbgTracker "---------------------------------"
		DbgTracker "$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"
		DbgTracker "^^^^		GAME OVER        ^^^^"
		DbgTracker "            PLAYER " & CurrentPlayer
		DbgTracker "$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"
		DbgTracker "---------------------------------"

		DbgTracker "LANES: " &Tracker_Lanes  &" POINTS: " & Tracker_Lanes*1000
		DbgTracker "RAMPS: " &Tracker_Ramps  &" POINTS: " & Tracker_Ramps*10000
		DbgTracker "ORBITS: " &Tracker_Orbits &" POINTS: " & Tracker_Orbits'*SCORE_ORBITS
		DbgTracker "TARGETS: " &Tracker_Targets &" POINTS: " & Tracker_Targets*2000
		DbgTracker "BUMPERS: " &Tracker_Bumpers &" POINTS: " & Tracker_Bumpers'*SCORE_BUMPER
		DbgTracker "SPINNERS: " &Tracker_Spinners &" POINTS: " & Tracker_Spinners'*SCORE_SPINNER
		DbgTracker "BURRITOS: " &Tracker_Burritos &" POINTS: " & Tracker_Spinners*25000
		DbgTracker "COMBOHITS: " &Tracker_combohits &" POINTS: " & Tracker_combohits*150000
		DbgTracker "BONUS: " &Tracker_BonusPoints 
		DbgTracker "#################################"

    
' you may wish to light any Game Over Light you may have
End Sub

Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' lost a ball ;-( check to see how many balls are on the playfield.
' if only one then decrement the remaining count AND test for End of game
' if more than 1 ball (multi-ball) then kill of the ball but don't create
' a new one
'
Sub Drain_Hit()
	debug.print "DRAIN:" &gametime
	debug.print "BS:" &bBallSaverActive

    ' Destroy the ball
    Drain.DestroyBall
    ' Exit Sub ' only for debugging - this way you can add balls from the debug window

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' pretend to knock the ball into the ball storage mech
    RandomSoundDrain Drain
    'if Tilted the end Ball Mode
    If Tilted Then
        StopEndOfBallMode
    End If

    ' if there is a game in progress AND it is not Tilted
    If(bGameInPLay = True)AND(Tilted = False)Then

        ' is the ball saver active,
        If(bBallSaverActive = True)Then

            ' yep, create a new ball in the shooters lane
            ' we use the Addmultiball in case the multiballs are being ejected
            AddMultiball 1
            ' we kick the ball with the autoplunger
            bAutoPlunger = True
            ' you may wish to put something on a display or play a sound at this point
            DMD "_", CL("BALL SAVED"), "_", eNone, eBlinkfast, eNone, 800, True, ""
            PlayBALLSAVE

			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Ball Saved !"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If

        Else
            ' cancel any multiball if on last ball (ie. lost all other balls)
            If(BallsOnPlayfield = 1)Then
                ' AND in a multi-ball??
                If(bMultiBallMode = True)then
                    ' not in multiball mode any more
                    bMultiBallMode = False
					bMainMultiballMode = False
					bAutoPlunger = False


					if bBlizzardPrepMode Then
						bBlizzardPrepMode = False
						PuPlayer.playevent pDMDVideo,"Misc","BlizzardCollect.mp4",0,60,5,0,""
					Else
						if Target009.isDropped Then ResetBLIZZARDLights : ResetCountdown
					End If

					StopFire
                    StopMinionMode
                    StopSpots
                    StopSpinner
					RiseTarget	

                    ' you may wish to change any music over at this point and
                    ' turn off any multiball specific lights
 

                    ResetJackpotLights
                    Select Case Battle(CurrentPlayer, 0)
                        Case 13:WinBattle
                    End Select
                    ChangeGi white
                    'ChangeSong

                    if bBlizzardMode Then StopBlizzardMode
					if bWizMode1Active Then StopWizMode1
					if bWizMode2Active Then StopWizMode2
					if bWizMode3Active Then StopWizMode3
					if bFinalWizModeActive Then StopFinalWizMode

                End If
            End If

            ' was that the last ball on the playfield
            If(BallsOnPlayfield = 0)Then
				bBlizzardPrepMode = False
				PuPlayer.LabelSet pDMD,"Event3A","",0,""
				PuPlayer.LabelSet pDMD,"Event3B","",0,""
				PuPlayer.LabelSet pDMD,"Event3Ca","",0,""
				PuPlayer.LabelSet pDMD,"Event3C","",0,""
	

                ' End Mode and timers
                'ChangeSong
                ChangeGi white
                ' Show the end of ball animation
                ' and continue with the end of ball
                ' DMD something?
                StopEndOfBallMode
                vpmtimer.addtimer 200, "EndOfBall '" 'the delay is depending of the animation of the end of ball, since there is no animation then move to the end of ball
            End If
        End If
    End If
End Sub

' The Ball has rolled out of the Plunger Lane and it is pressing down the trigger in the shooters lane
' Check to see if a ball saver mechanism is needed and if so fire it up.

Sub swPlungerRest_Hit()
    'debug.print "ball in plunger lane"
    ' some sound according to the ball position
    PlaySoundAt "fx_sensor", swPlungerRest
    bBallInPlungerLane = True
	DOF 142, DOFOn
    ' turn on Launch light is there is one

    'be sure to update the Scoreboard after the animations, if any

    ' kick the ball in play if the bAutoPlunger flag is on
    If bAutoPlunger Then
        'debug.print "autofire the ball"

'activeball.vely = -30
        PlungerIM.AutoFire
        DOF 121, DOFPulse
		DOF 141, DOFPulse
        PlaySoundAt "fx_fire", swPlungerRest
        bAutoPlunger = False
    End If
    ' if there is a need for a ball saver, then start off a timer
    ' only start if it is ready, and it is currently not running, else it will reset the time period
    If(bBallSaverReady = True)AND(BallSaverTime <> 0)And(bBallSaverActive = False)Then
'		if Not bBLIZZARDMode Then
'			EnableBallSaver BallSaverTime
'        End if
	Else	
	'debug.print "WTH"
        ' show the message to shoot the ball in case the player has fallen sleep
        swPlungerRest.TimerEnabled = 1
    End If
    'Start the Selection of the skillshot if ready
    If bSkillShotReady Then
        UpdateSkillshot()
    End If
    ' remember last trigger hit by the ball.
    LastSwitchHit = "swPlungerRest"
End Sub

' The ball is released from the plunger turn off some flags and check for skillshot

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
	DOF 142, DOFOff

    If(bBallSaverReady = True)AND(BallSaverTime <> 0)And(bBallSaverActive = False)Then
		if Not bBLIZZARDMode Then
			EnableBallSaver BallSaverTime
        End if
	End If
	
    swPlungerRest.TimerEnabled = 0 'stop the launch ball timer if active
    If bSkillShotReady Then
        ResetSkillShotTimer.Enabled = 1
    End If
    If bMultiballMode Then
		bAutoPlunger = True
        If BallsOnPlayfield = 2 Then
            'ChangeSong
        End If
    Else
        'ChangeSong
    End If
End Sub

' swPlungerRest timer to show the "launch ball" if the player has not shot the ball during 6 seconds

Sub swPlungerRest_Timer
    DMD "_", CL("SHOOT THE BALL"), "_", eNone, eNone, eNone, 800, True, ""
    swPlungerRest.TimerEnabled = 0
End Sub

Sub EnableBallSaver(seconds)
    'debug.print "Ballsaver started"
    ' set our game flag
	debug.print "Enable Ball Saver Called"
    bBallSaverActive = True
    bBallSaverReady = False
    ' start the timer
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
debug.print "LSH:" &LastSwitchHit
	if LastSwitchHit = "RightOutlane" or LastSwitchHit = "LeftOutlane" Then
debug.print "HERE"
		BallHandlingQueue.Add "bBallSaverActive = False","bBallSaverActive = False",100,2000,0,0,0,True
	Else
		BallHandlingQueue.Add "bBallSaverActive = False","bBallSaverActive = False",100,1250,0,0,0,True
	End If
    ' if you have a ball saver light then turn it off at this point
    LightShootAgain.State = 0
End Sub

Sub BallSaverSpeedUpTimer_Timer()
    'debug.print "Ballsaver Speed Up Light"
    BallSaverSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    LightShootAgain.BlinkInterval = 80
    LightShootAgain.State = 2
End Sub

' *********************************************************************
'                      Supporting Score Functions
' *********************************************************************

' Add points to the score AND update the score board
' In this table we use SecondRound variable to double the score points in the second round after killing Malthael
Sub AddScore(points)
    If(Tilted = False)Then
        ' add the points to the current players score variable
        Score(CurrentPlayer) = Score(CurrentPlayer) + points * PlayfieldMultiplier(CurrentPlayer)
    End if

	if PlayersPlayingGame = 1 Then
		PuPlayer.LabelSet pDMD,"CurrScore"," " & FormatScoreDMD(Score(CurrentPlayer)) & " ",1,"{'mt':2,'fonth':17,'xpos':50,'ypos':87.5}"	
	Else
		PuPlayer.LabelSet pDMD,"CurrScore"," " & FormatScoreDMD(Score(CurrentPlayer)) & " ",1,"{'mt':2,'fonth':13,'xpos':50,'ypos':87.5}"
	End If
End Sub

' Add bonus to the bonuspoints AND update the score board
Sub AddBonus(points) 'not used in this table, since there are many different bonus items.
    If(Tilted = False)Then
        ' add the bonus to the current players bonus variable
        BonusPoints(CurrentPlayer) = BonusPoints(CurrentPlayer) + points
    End if
End Sub

' Add some points to the current Jackpot.
'
Sub AddJackpot(points)
    ' Jackpots only generally increment in multiball mode AND not tilted
    ' but this doesn't have to be the case
    If(Tilted = False)Then

        ' If(bMultiBallMode = True) Then
        Jackpot(CurrentPlayer) = Jackpot(CurrentPlayer) + points
        DMD "_", CL("INCREASED JACKPOT"), "_", eNone, eNone, eNone, 800, True, ""
        PlayJackpotIncr
    ' you may wish to limit the jackpot to a upper limit, ie..
    '	If (Jackpot >= 6000) Then
    '		Jackpot = 6000
    ' 	End if
    'End if
    End if
End Sub

Sub AddSuperJackpot(points) 'not used in this table
    If(Tilted = False)Then
    End if
End Sub

Sub AddBonusMultiplier(n)
    Dim NewBonusLevel
    ' if not at the maximum bonus level
    if(BonusMultiplier(CurrentPlayer) + n <= MaxMultiplier)then
        ' then add and set the lights
        NewBonusLevel = BonusMultiplier(CurrentPlayer) + n
        SetBonusMultiplier(NewBonusLevel)
        DMD "_", CL("BONUS X " &NewBonusLevel), "_", eNone, eNone, eNone, 2000, True, ""
        PlayBonusmp
		PlayMultipliers
        LeftFlash
        RightFlash
        BackLeftFlash
        BackRightFlash

	if Scorbit.bSessionActive then
		GameModeStrTmp="NA{Blue}:Bonus X: " &NewBonusLevel
		if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
	End If

    Else
        AddScore 50000
        DMD "_", CL("50000"), "_", eNone, eNone, eNone, 800, True, ""
    End if
End Sub

' Set the Bonus Multiplier to the specified level AND set any lights accordingly

Sub SetBonusMultiplier(Level)
    ' Set the multiplier to the specified level
    BonusMultiplier(CurrentPlayer) = Level
    UPdateBonusXLights(Level)
	DisplayBonusIcon
End Sub

Sub DisplayBonusIcon

	Select Case BonusMultiplier(CurrentPlayer)
		Case 2
			HideAllBonusIcons
			PuPlayer.LabelSet pDMD, "BonusMP2x","PupOverlays\\2x.png",1,""
		Case 3
			HideAllBonusIcons
			PuPlayer.LabelSet pDMD, "BonusMP3x","PupOverlays\\3x.png",1,""
		Case 4
			HideAllBonusIcons
			PuPlayer.LabelSet pDMD, "BonusMP4x","PupOverlays\\4x.png",1,""
		Case 5
			HideAllBonusIcons
			PuPlayer.LabelSet pDMD, "BonusMP5x","PupOverlays\\5x.png",1,""
		Case Else
			HideAllBonusIcons
	End Select
End Sub

Sub HideAllBonusIcons

	PuPlayer.LabelSet pDMD, "BonusMP2x","PupOverlays\\2x.png",0,""
	PuPlayer.LabelSet pDMD, "BonusMP3x","PupOverlays\\3x.png",0,""
	PuPlayer.LabelSet pDMD, "BonusMP4x","PupOverlays\\4x.png",0,""
	PuPlayer.LabelSet pDMD, "BonusMP5x","PupOverlays\\5x.png",0,""
End Sub

Sub UpdateBonusXLights(Level)
    ' Update the lights
    Select Case Level
        Case 1:light56.State = 0:light57.State = 0:light58.State = 0:light59.State = 0:light005.State = 0:light006.State = 0:light007.State = 0:light008.State = 0
        Case 2:light56.State = 1:light57.State = 0:light58.State = 0:light59.State = 0:light005.State = 0:light006.State = 0:light007.State = 1:light008.State = 0
        Case 3:light56.State = 0:light57.State = 1:light58.State = 0:light59.State = 0:light005.State = 0:light006.State = 1:light007.State = 0:light008.State = 0
        Case 4:light56.State = 0:light57.State = 0:light58.State = 1:light59.State = 0:light005.State = 0:light006.State = 0:light007.State = 0:light008.State = 1
        Case 5:light56.State = 0:light57.State = 0:light58.State = 0:light59.State = 1:light005.State = 1:light006.State = 0:light007.State = 0:light008.State = 0
    End Select
End Sub

Sub AddPlayfieldMultiplier(n)
    Dim NewPFLevel
    ' if not at the maximum level x
    if(PlayfieldMultiplier(CurrentPlayer) + n <= MaxMultiplier)then
        ' then add and set the lights
        NewPFLevel = PlayfieldMultiplier(CurrentPlayer) + n
        SetPlayfieldMultiplier(NewPFLevel)
        DMD "_", CL("PLAYFIELD X " &NewPFLevel), "_", eNone, eNone, eNone, 2000, True, ""
        TimerJAG.Enabled = True
        FlashForMs BLight001, 1000, 50, 0
        Playsound "WW1"
        PlayPlayfieldInc
        PlayPlayfieldMultipliers
        LeftFlash
        RightFlash
        BackLeftFlash
        BackRightFlash

	if Scorbit.bSessionActive then
		GameModeStrTmp="NA{Blue}:Playfield X: " &NewPFLevel
		if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
	End If


    Else 'if the 5x is already lit
        AddScore 50000
        DMD "_", CL("50000"), "_", eNone, eNone, eNone, 2000, True, ""
    End if
End Sub

' Set the Playfield Multiplier to the specified level AND set any lights accordingly

Sub SetPlayfieldMultiplier(Level)
    ' Set the multiplier to the specified level
    PlayfieldMultiplier(CurrentPlayer) = Level
    UpdatePFXLights(Level)
End Sub

Sub UpdatePFXLights(Level)
' Update the lights
Select Case Level
    Case 1:LP3.State = 0:LP2.State = 0:LP1.State = 0:LP4.State = 0
    Case 2:LP3.State = 1:LP2.State = 0:LP1.State = 0:LP4.State = 0
    Case 3:LP3.State = 0:LP2.State = 1:LP1.State = 0:LP4.State = 0
    Case 4:LP3.State = 0:LP2.State = 0:LP1.State = 1:LP4.State = 0
    Case 5:LP3.State = 0:LP2.State = 0:LP1.State = 0:LP4.State = 1
End Select
' show the multiplier in the DMD
End Sub

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        PlayExtraBall
        DMD "_", CL(("EXTRA BALL WON")), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
        SupressModeMessages 2500
        PuPlayer.playevent pDMDVideo,"ExtraBall","ExtraBall.mp4",nPupVideoVolume,65,3,0,""
        DOF 121, DOFPulse
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        GiEffect 2
        LightEffect 2
    END If
End Sub

Sub AwardSpecial()
    DMD "_", CL(("EXTRA GAME WON")), "_", eNone, eBlink, eNone, 1000, True, SoundFXDOF("", 122, DOFPulse, DOFKnocker)
    DOF 121, DOFPulse
    Credits = Credits + 1
    If bFreePlay = False Then DOF 125, DOFOn
    LightEffect 2
    FlashEffect 2
    PlayExtragame
End Sub

Sub AwardJackpot() 'award a normal jackpot, double or triple jackpot
    PlayJackpotsound
    if NOT JackpotFlashTimer.Enabled Then
		JPCount = 0
		JackpotFlashTimer.Enabled = True
	End If
    Dim tmp
    DMD CL(FormatScore(Jackpot(CurrentPlayer))), CL("JACKPOT"), "d_border", eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    SupressModeMessages 2500
    PuPlayer.playevent pDMDVideo,"Jackpot","Jackpot.mp4",nPupVideoVolume,65,3,0,""
    DOF 157, DOFPulse
    tmp = INT(RND * 2)
    PlayJackpotsound
    AddScore Jackpot(CurrentPlayer)*nWizardModeMultiplier
    LightEffect 2
    FlashEffect 2
   
End Sub

Sub AwardSuperJackpot() 'this is actually 4 times a jackpot
    if NOT JackpotFlashTimer.Enabled Then
		JPCount = 0
		JackpotFlashTimer.Enabled = True
	End If
    SuperJackpot = Jackpot(CurrentPlayer) * 4
    DMD CL(FormatScore(SuperJackpot)), CL("SUPER JACKPOT"), "d_border", eBlinkFast, eBlinkFast, eNone, 1000, True, ""
    SupressModeMessages 2500
	PuPlayer.playevent pDMDVideo,"Jackpot","SuperJackpot.mp4",nPupVideoVolume,65,3,0,""
    PlaySuperJackpotsnd
    DOF 158, DOFPulse	
    AddScore SuperJackpot*nWizardModeMultiplier
    LightEffect 2
    FlashEffect 2
    'enabled jackpots again
    StartJackpots
End Sub

Sub AwardSkillshot()
    dim tmp
    tmp = INT(RND * 7) + 1
    PlaySound "fx_thunder" & tmp
    ResetSkillShotTimer_Timer
    'show dmd animation
    DMD CL(FormatScore(SkillshotValue(CurrentPlayer))), CL(("SKILLSHOT")), "d_border", eBlinkFast, eBlink, eNone, 1000, True, ""
    DOF 127, DOFPulse
    PlaySkillshot
    SupressModeMessages 2500
    PuPlayer.playevent pDMDVideo,"Skillshot","Skillshot.mp4",nPupVideoVolume,65,3,0,""
    LightningStrike()
    Addscore SkillShotValue(CurrentPlayer)
    ' increment the skillshot value with 250.000
    SkillShotValue(CurrentPlayer) = SkillShotValue(CurrentPlayer) + 250000
    'do some light show
    GiEffect 2
    LightEffect 2

	if Scorbit.bSessionActive then
		GameModeStrTmp="NA{Red}:Skillshot: "&SkillshotValue(CurrentPlayer) 
		if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
	End If

End Sub

'*****************************
'    Load / Save / Highscore
'*****************************

Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "")Then HighScore(0) = CDbl(x)Else HighScore(0) = 100000 End If
    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "")Then HighScoreName(0) = x Else HighScoreName(0) = "AAA" End If
    x = LoadValue(TableName, "HighScore2")
    If(x <> "")then HighScore(1) = CDbl(x)Else HighScore(1) = 100000 End If
    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "")then HighScoreName(1) = x Else HighScoreName(1) = "BBB" End If
    x = LoadValue(TableName, "HighScore3")
    If(x <> "")then HighScore(2) = CDbl(x)Else HighScore(2) = 100000 End If
    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "")then HighScoreName(2) = x Else HighScoreName(2) = "CCC" End If
    x = LoadValue(TableName, "HighScore4")
    If(x <> "")then HighScore(3) = CDbl(x)Else HighScore(3) = 100000 End If
    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "")then HighScoreName(3) = x Else HighScoreName(3) = "DDD" End If
    x = LoadValue(TableName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0:If bFreePlay = False Then DOF 125, DOFOff:End If
    x = LoadValue(TableName, "TotalGamesPlayed")
    If(x <> "")then TotalGamesPlayed = CInt(x)Else TotalGamesPlayed = 0 End If
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
    SaveValue TableName, "Credits", Credits
    SaveValue TableName, "TotalGamesPlayed", TotalGamesPlayed
End Sub

Sub Reseths
    HighScoreName(0) = "OZZ"
    HighScoreName(1) = "ZAN"
    HighScoreName(2) = "RTP"
    HighScoreName(3) = "JEB"
    HighScore(0) = 50000000
    HighScore(1) = 40000000
    HighScore(2) = 25000000
    HighScore(3) = 10000000
    Savehs
End Sub

' ***********************************************************
'  High Score Initals Entry Functions - based on Black's code
' ***********************************************************

Dim hsbModeActive
Dim hsEnteredName
Dim hsEnteredDigits(3)
Dim hsCurrentDigit
Dim hsValidLetters
Dim hsCurrentLetter
Dim hsLetterFlash

Sub CheckHighscore()
    Dim tmp
    tmp = Score(CurrentPlayer)

    If tmp > HighScore(0)Then 'add 1 credit for beating the highscore
        Credits = Credits + 1
        DOF 125, DOFOn
    End If

    If tmp > HighScore(3)Then
        PlaySound SoundFXDOF("fx_Knocker", 122, DOFPulse, DOFKnocker)
        DOF 121, DOFPulse
        HighScore(3) = tmp
        'enter player's name
        HighScoreEntryInit()
    Else
        EndOfBallComplete()
    End If
End Sub

Sub HighScoreEntryInit()
    hsbModeActive = True
    'PlaySound "vo_greatscore" &RndNbr(6)
    hsLetterFlash = 0

    hsEnteredDigits(0) = " "
    hsEnteredDigits(1) = " "
    hsEnteredDigits(2) = " "
    hsCurrentDigit = 0

    hsValidLetters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<" ' < is back arrow
    hsCurrentLetter = 1

    DMDFlush()
    HighScoreDisplayNameNow()

	ClearHighScoreTwoLine
	HighScoreHelper "YOUR NAME:", " <A  > ", 9999

    HighScoreFlashTimer.Interval = 250
    HighScoreFlashTimer.Enabled = True

End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        playsound "fx_Previous"
        hsCurrentLetter = hsCurrentLetter - 1
        if(hsCurrentLetter = 0)then
            hsCurrentLetter = len(hsValidLetters)
        end if
        HighScoreDisplayName()
    End If

    If keycode = RightFlipperKey Then
        playsound "fx_Next"
        hsCurrentLetter = hsCurrentLetter + 1
        if(hsCurrentLetter > len(hsValidLetters))then
            hsCurrentLetter = 1
        end if
        HighScoreDisplayName()
    End If

    If keycode = PlungerKey OR keycode = StartGameKey Then
        if(mid(hsValidLetters, hsCurrentLetter, 1) <> "<")then
            playsound "fx_Enter"
            hsEnteredDigits(hsCurrentDigit) = mid(hsValidLetters, hsCurrentLetter, 1)
            hsCurrentDigit = hsCurrentDigit + 1
            if(hsCurrentDigit = 3)then
                HighScoreCommitName()
            else
                HighScoreDisplayName()
            end if
        else
            playsound "fx_Esc"
            hsEnteredDigits(hsCurrentDigit) = " "
            if(hsCurrentDigit > 0)then
                hsCurrentDigit = hsCurrentDigit - 1
            end if
            HighScoreDisplayName()
        end if
    end if
End Sub

Sub HighScoreDisplayNameNow()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreDisplayName()
    Dim i
    Dim TempTopStr
    Dim TempBotStr

    TempTopStr = "YOUR NAME:"
    dLine(0) = ExpandLine(TempTopStr)
    DMDUpdate 0

    TempBotStr = " >"
    if(hsCurrentDigit > 0)then TempBotStr = TempBotStr & hsEnteredDigits(0)
    if(hsCurrentDigit > 1)then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit > 2)then TempBotStr = TempBotStr & hsEnteredDigits(2)

    if(hsCurrentDigit <> 3)then
        if(hsLetterFlash <> 0)then
            TempBotStr = TempBotStr & "_"
        else
            TempBotStr = TempBotStr & mid(hsValidLetters, hsCurrentLetter, 1)
        end if
    end if

    if(hsCurrentDigit < 1)then TempBotStr = TempBotStr & hsEnteredDigits(1)
    if(hsCurrentDigit < 2)then TempBotStr = TempBotStr & hsEnteredDigits(2)

    TempBotStr = TempBotStr & "< "

    dLine(1) = ExpandLine(TempBotStr)
    DMDUpdate 1

	HighScoreHelper "YOUR NAME:", Mid(TempBotStr, 2, 5), 9999
End Sub

Sub HighScoreFlashTimer_Timer()
    HighScoreFlashTimer.Enabled = False
    hsLetterFlash = hsLetterFlash + 1
    if(hsLetterFlash = 2)then hsLetterFlash = 0
    HighScoreDisplayName()
    HighScoreFlashTimer.Enabled = True
End Sub

Sub HighScoreCommitName()
    HighScoreFlashTimer.Enabled = False
    hsbModeActive = False

    hsEnteredName = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    if(hsEnteredName = "   ")then
        hsEnteredName = "YOU"
    end if

    HighScoreName(3) = hsEnteredName
    SortHighscore
	ClearHighScoreTwoLine
    EndOfBallComplete()
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 0 to 3
        For j = 0 to 2
            If HighScore(j) < HighScore(j + 1)Then
                tmp = HighScore(j + 1)
                tmp2 = HighScoreName(j + 1)
                HighScore(j + 1) = HighScore(j)
                HighScoreName(j + 1) = HighScoreName(j)
                HighScore(j) = tmp
                HighScoreName(j) = tmp2
            End If
        Next
    Next
End Sub

' *************************************************************************
'   JP's Reduced Display Driver Functions (based on script by Black)
' only 5 effects: none, scroll left, scroll right, blink and blinkfast
' 3 Lines, treats all 3 lines as text.
' 1st and 2nd lines are 20 characters long
' 3rd line is just 1 character
' Example format:
' DMD "text1","text2","backpicture", eNone, eNone, eNone, 250, True, "sound"
' Short names:
' dq = display queue
' de = display effect
' *************************************************************************

Const eNone = 0        ' Instantly displayed
Const eScrollLeft = 1  ' scroll on from the right
Const eScrollRight = 2 ' scroll on from the left
Const eBlink = 3       ' Blink (blinks for 'TimeOn')
Const eBlinkFast = 4   ' Blink (blinks for 'TimeOn') at user specified intervals (fast speed)

Const dqSize = 64

Dim dqHead
Dim dqTail
Dim deSpeed
Dim deBlinkSlowRate
Dim deBlinkFastRate

Dim dCharsPerLine(2)
Dim dLine(2)
Dim deCount(2)
Dim deCountEnd(2)
Dim deBlinkCycle(2)

Dim dqText(2, 64)
Dim dqEffect(2, 64)
Dim dqTimeOn(64)
Dim dqbFlush(64)
Dim dqSound(64)

Dim FlexDMD
Dim DMDScene

Sub DMD_Init() 'default/startup values
    If UseFlexDMD Then
        Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
        If Not FlexDMD is Nothing Then
            If FlexDMDHighQuality Then
                FlexDMD.TableFile = Table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 256
                FlexDMD.Height = 64
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                For i = 0 to 40
                    DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
                    Digits(i).Visible = False
                Next
                digitgrid.Visible = False
                For i = 0 to 19 ' Top
                    DMDScene.GetImage("Dig" & i).SetBounds 8 + i * 12, 6, 12, 22
                Next
                For i = 20 to 39 ' Bottom
                    DMDScene.GetImage("Dig" & i).SetBounds 8 + (i - 20) * 12, 34, 12, 22
                Next
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            Else
                FlexDMD.TableFile = Table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 128
                FlexDMD.Height = 32
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.d_border")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                For i = 0 to 40
                    DMDScene.AddActor FlexDMD.NewImage("Dig" & i, "VPX.d_empty&dmd=2")
                    Digits(i).Visible = False
                Next
                digitgrid.Visible = False
                For i = 0 to 19 ' Top
                    DMDScene.GetImage("Dig" & i).SetBounds 4 + i * 6, 3, 6, 11
                Next
                For i = 20 to 39 ' Bottom
                    DMDScene.GetImage("Dig" & i).SetBounds 4 + (i - 20) * 6, 17, 6, 11
                Next
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            End If
        End If
    End If

    Dim i, j
    DMDFlush()
    deSpeed = 20
    deBlinkSlowRate = 10
    deBlinkFastRate = 5
    dCharsPerLine(0) = 20 'characters lower line
    dCharsPerLine(1) = 20 'characters top line
    dCharsPerLine(2) = 1  'characters back line
    For i = 0 to 2
        dLine(i) = Space(dCharsPerLine(i))
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
        dqTimeOn(i) = 0
        dqbFlush(i) = True
        dqSound(i) = ""
    Next
    dLine(2) = " "
    For i = 0 to 2
        For j = 0 to 64
            dqText(i, j) = ""
            dqEffect(i, j) = eNone
        Next
    Next
    DMD dLine(0), dLine(1), dLine(2), eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDFlush()
    Dim i
    DMDTimer.Enabled = False
    DMDEffectTimer.Enabled = False
    dqHead = 0
    dqTail = 0
    For i = 0 to 2
        deCount(i) = 0
        deCountEnd(i) = 0
        deBlinkCycle(i) = 0
    Next
End Sub

Sub DMDScore()
    Dim tmp, tmp1, tmp2
    if(dqHead = dqTail)Then
        tmp = RL(FormatScore(Score(Currentplayer)))
        'tmp = CL(FormatScore(Score(Currentplayer) ) )
        'tmp1 = CL("PLAYER " & CurrentPlayer & " BALL " & Balls)
        'tmp1 = FormatScore(Bonuspoints(Currentplayer) ) & " X" &BonusMultiplier(Currentplayer)

        Select Case Battle(CurrentPlayer, 0)
            Case 0:tmp1 = CL("PLAYER " & CurrentPlayer & " BALL " & Balls & " X" & PlayfieldMultiplier(CurrentPlayer))
            Case 1:tmp1 = CL("SPINNERS LEFT " & 100-SpinCount)
            Case 2:tmp1 = CL("BUMPER HITS LEFT " & 25-SuperBumperHIts)
            Case 3:tmp1 = CL("RAMP HITS LEFT " & 6-ramphits3)
            Case 4:tmp1 = CL("ORBIT HITS LEFT " & 6-orbithits)
            Case 5:tmp1 = CL("HIT THE LIGHTS")
            Case 6:tmp1 = CL("HIT THE LIGHTS")
            Case 7:tmp1 = CL("HIT THE TARGETS " & 20-TargetHits7)
            Case 8:tmp1 = CL("HIT THE TARGETS " & 6-TargetHits8)
            Case 9:tmp1 = CL("HIT THE LIT LIGHT " & 8-LightHits9)
            Case 10:tmp1 = CL("HIT THE LOOPS " & 6-loopCount)
            Case 11:tmp1 = CL("HIT THE LIT LIGHT " & 8-LightHits11)
            Case 12:tmp1 = CL("HIT RAMPS ORBITS " & 6-RampHits12)
            Case 13:tmp1 = CL("HIT THE JACKPOTS")
        End Select
        tmp2 = ""
    End If
    DMD tmp, tmp1, tmp2, eNone, eNone, eNone, 25, True, ""
End Sub

Sub DMDScoreNow
    DMDFlush
    DMDScore
End Sub

Sub DMD(Text0, Text1, Text2, Effect0, Effect1, Effect2, TimeOn, bFlush, Sound)
    if(dqTail < dqSize)Then
        if(Text0 = "_")Then
            dqEffect(0, dqTail) = eNone
            dqText(0, dqTail) = "_"
        Else
            dqEffect(0, dqTail) = Effect0
            dqText(0, dqTail) = ExpandLine(Text0)
        End If

        if(Text1 = "_")Then
            dqEffect(1, dqTail) = eNone
            dqText(1, dqTail) = "_"
        Else
            dqEffect(1, dqTail) = Effect1
            dqText(1, dqTail) = ExpandLine(Text1)
        End If

        if(Text2 = "_")Then
            dqEffect(2, dqTail) = eNone
            dqText(2, dqTail) = "_"
        Else
            dqEffect(2, dqTail) = Effect2
            dqText(2, dqTail) = Text2 'it is always 1 letter in this table
        End If

        dqTimeOn(dqTail) = TimeOn
        dqbFlush(dqTail) = bFlush
        dqSound(dqTail) = Sound
        dqTail = dqTail + 1
        if(dqTail = 1)Then
            DMDHead()
        End If
    End If
End Sub

Sub DMDHead()
    Dim i
    deCount(0) = 0
    deCount(1) = 0
    deCount(2) = 0
    DMDEffectTimer.Interval = deSpeed

    For i = 0 to 2
        Select Case dqEffect(i, dqHead)
            Case eNone:deCountEnd(i) = 1
            Case eScrollLeft:deCountEnd(i) = Len(dqText(i, dqHead))
            Case eScrollRight:deCountEnd(i) = Len(dqText(i, dqHead))
            Case eBlink:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
            Case eBlinkFast:deCountEnd(i) = int(dqTimeOn(dqHead) / deSpeed)
                deBlinkCycle(i) = 0
        End Select
    Next
    if(dqSound(dqHead) <> "")Then
        PlaySound(dqSound(dqHead))
    End If
    DMDEffectTimer.Enabled = True
End Sub

Sub DMDEffectTimer_Timer()
    DMDEffectTimer.Enabled = False
    DMDProcessEffectOn()
End Sub

Sub DMDTimer_Timer()
    Dim Head
    DMDTimer.Enabled = False
    Head = dqHead
    dqHead = dqHead + 1
    if(dqHead = dqTail)Then
        if(dqbFlush(Head) = True)Then
            DMDScoreNow()
        Else
            dqHead = 0
            DMDHead()
        End If
    Else
        DMDHead()
    End If
End Sub

Sub DMDProcessEffectOn()
    Dim i
    Dim BlinkEffect
    Dim Temp

    BlinkEffect = False

    For i = 0 to 2
        if(deCount(i) <> deCountEnd(i))Then
            deCount(i) = deCount(i) + 1

            select case(dqEffect(i, dqHead))
                case eNone:
                    Temp = dqText(i, dqHead)
                case eScrollLeft:
                    Temp = Right(dLine(i), dCharsPerLine(i)- 1)
                    Temp = Temp & Mid(dqText(i, dqHead), deCount(i), 1)
                case eScrollRight:
                    Temp = Mid(dqText(i, dqHead), (dCharsPerLine(i) + 1)- deCount(i), 1)
                    Temp = Temp & Left(dLine(i), dCharsPerLine(i)- 1)
                case eBlink:
                    BlinkEffect = True
                    if((deCount(i)MOD deBlinkSlowRate) = 0)Then
                        deBlinkCycle(i) = deBlinkCycle(i)xor 1
                    End If

                    if(deBlinkCycle(i) = 0)Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i))
                    End If
                case eBlinkFast:
                    BlinkEffect = True
                    if((deCount(i)MOD deBlinkFastRate) = 0)Then
                        deBlinkCycle(i) = deBlinkCycle(i)xor 1
                    End If

                    if(deBlinkCycle(i) = 0)Then
                        Temp = dqText(i, dqHead)
                    Else
                        Temp = Space(dCharsPerLine(i))
                    End If
            End Select

            if(dqText(i, dqHead) <> "_")Then
                dLine(i) = Temp
                DMDUpdate i
            End If
        End If
    Next

    if(deCount(0) = deCountEnd(0))and(deCount(1) = deCountEnd(1))and(deCount(2) = deCountEnd(2))Then

        if(dqTimeOn(dqHead) = 0)Then
            DMDFlush()
        Else
            if(BlinkEffect = True)Then
                DMDTimer.Interval = 10
            Else
                DMDTimer.Interval = dqTimeOn(dqHead)
            End If

            DMDTimer.Enabled = True
        End If
    Else
        DMDEffectTimer.Enabled = True
    End If
End Sub

Function ExpandLine(TempStr) 'id is the number of the dmd line
    If TempStr = "" Then
        TempStr = Space(20)
    Else
        if Len(TempStr) > Space(20)Then
            TempStr = Left(TempStr, Space(20))
        Else
            if(Len(TempStr) < 20)Then
                TempStr = TempStr & Space(20 - Len(TempStr))
            End If
        End If
    End If
    ExpandLine = TempStr
End Function

Function FormatScore(ByVal Num) 'it returns a string with commas (as in Black's original font)
    dim i
    dim NumString
    NumString = CStr(abs(Num))
    For i = Len(NumString)-3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1))then
            NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1)) + 128) & right(NumString, Len(NumString)- i)
        end if
    Next
    FormatScore = NumString
End function

Function FormatScoreDMD(ByVal Num) 'it returns a string with commas (as in Black's original font)
    dim i
    dim NumString

	if Num=0 then 
		FormatScoreDMD="0"
		Exit Function
	End if 

    NumString = CStr(abs(Num))

    For i = Len(NumString)-3 to 1 step -3
        if IsNumeric(mid(NumString, i, 1))then
            NumString = left(NumString, i) & "," & right(NumString, Len(NumString)-i)   ' ANDREW
		   'NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1)) + 48) & right(NumString, Len(NumString)- i)
        end if
    Next
    FormatScoreDMD = NumString
End function

Function FL(NumString1, NumString2) 'Fill line
    Dim Temp, TempStr
    If Len(NumString1) + Len(NumString2) < 20 Then
        Temp = 20 - Len(NumString1)- Len(NumString2)
        TempStr = NumString1 & Space(Temp) & NumString2
        FL = TempStr
    End If
End Function

Function CL(NumString) 'center line
    Dim Temp, TempStr
    If Len(NumString) > 20 Then NumString = Left(NumString, 20)
    Temp = (20 - Len(NumString)) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(NumString) 'right line
    Dim Temp, TempStr
    If Len(NumString) > 20 Then NumString = Left(NumString, 20)
    Temp = 20 - Len(NumString)
    TempStr = Space(Temp) & NumString
    RL = TempStr
End Function

'**************
' Update DMD
'**************

Sub DMDUpdate(id)
    Dim digit, value
    If UseFlexDMD Then FlexDMD.LockRenderThread
    Select Case id
        Case 0 'top text line
            For digit = 0 to 19
                DMDDisplayChar mid(dLine(0), digit + 1, 1), digit
            Next
        Case 1 'bottom text line
            For digit = 20 to 39
                DMDDisplayChar mid(dLine(1), digit -19, 1), digit
            Next
        Case 2 ' back image - back animations
            If dLine(2) = "" OR dLine(2) = " " Then dLine(2) = "bkempty"
            Digits(40).ImageA = dLine(2)
            If UseFlexDMD Then DMDScene.GetImage("Back").Bitmap = FlexDMD.NewImage("", "VPX." & dLine(2) & "&dmd=2").Bitmap
    End Select
    If UseFlexDMD Then FlexDMD.UnlockRenderThread
End Sub

Sub DMDDisplayChar(achar, adigit)
    If achar = "" Then achar = " "
    achar = ASC(achar)
    Digits(adigit).ImageA = Chars(achar)

    If UseFlexDMD Then DMDScene.GetImage("Dig" & adigit).Bitmap = FlexDMD.NewImage("", "VPX." & Chars(achar) & "&dmd=2&add").Bitmap
End Sub

'****************************
' JP's new DMD using flashers
'****************************

Dim Digits, DigitsBack, Chars(255), Images(255)

DMDInit

Sub DMDInit
    Dim i
    Digits = Array(digit001, digit002, digit003, digit004, digit005, digit006, digit007, digit008, digit009, digit010, _
        digit011, digit012, digit013, digit014, digit015, digit016, digit017, digit018, digit019, digit020,            _
        digit021, digit022, digit023, digit024, digit025, digit026, digit027, digit028, digit029, digit030,            _
        digit031, digit032, digit033, digit034, digit035, digit036, digit037, digit038, digit039, digit040,            _
        digit041)
    For i = 0 to 255:Chars(i) = "d_empty":Next

    Chars(32) = "d_empty"
    Chars(33) = ""        '!
    Chars(34) = ""        '"
    Chars(35) = ""        '#
    Chars(36) = ""        '$
    Chars(37) = ""        '%
    Chars(38) = ""        '&
    Chars(39) = ""        ''
    Chars(40) = ""        '(
    Chars(41) = ""        ')
    Chars(42) = ""        '*
    Chars(43) = ""        '+
    Chars(44) = ""        '
    Chars(45) = "d_minus" '-
    Chars(46) = "d_dot"   '.
    Chars(47) = ""        '/
    Chars(48) = "d_0"     '0
    Chars(49) = "d_1"     '1
    Chars(50) = "d_2"     '2
    Chars(51) = "d_3"     '3
    Chars(52) = "d_4"     '4
    Chars(53) = "d_5"     '5
    Chars(54) = "d_6"     '6
    Chars(55) = "d_7"     '7
    Chars(56) = "d_8"     '8
    Chars(57) = "d_9"     '9
    Chars(60) = "d_less"  '<
    Chars(61) = ""        '=
    Chars(62) = "d_more"  '>
    Chars(64) = ""        '@
    Chars(65) = "d_a"     'A
    Chars(66) = "d_b"     'B
    Chars(67) = "d_c"     'C
    Chars(68) = "d_d"     'D
    Chars(69) = "d_e"     'E
    Chars(70) = "d_f"     'F
    Chars(71) = "d_g"     'G
    Chars(72) = "d_h"     'H
    Chars(73) = "d_i"     'I
    Chars(74) = "d_j"     'J
    Chars(75) = "d_k"     'K
    Chars(76) = "d_l"     'L
    Chars(77) = "d_m"     'M
    Chars(78) = "d_n"     'N
    Chars(79) = "d_o"     'O
    Chars(80) = "d_p"     'P
    Chars(81) = "d_q"     'Q
    Chars(82) = "d_r"     'R
    Chars(83) = "d_s"     'S
    Chars(84) = "d_t"     'T
    Chars(85) = "d_u"     'U
    Chars(86) = "d_v"     'V
    Chars(87) = "d_w"     'W
    Chars(88) = "d_x"     'X
    Chars(89) = "d_y"     'Y
    Chars(90) = "d_z"     'Z
    Chars(94) = "d_up"    '^
    '    Chars(95) = '_
    Chars(96) = ""
    Chars(97) = ""  'a
    Chars(98) = ""  'b
    Chars(99) = ""  'c
    Chars(100) = "" 'd
    Chars(101) = "" 'e
    Chars(102) = "" 'f
    Chars(103) = "" 'g
    Chars(104) = "" 'h
    Chars(105) = "" 'i
    Chars(106) = "" 'j
    Chars(107) = "" 'k
    Chars(108) = "" 'l
    Chars(109) = "" 'm
    Chars(110) = "" 'n
    Chars(111) = "" 'o
    Chars(112) = "" 'p
    Chars(113) = "" 'q
    Chars(114) = "" 'r
    Chars(115) = "" 's
    Chars(116) = "" 't
    Chars(117) = "" 'u
    Chars(118) = "" 'v
    Chars(119) = "" 'w
    Chars(120) = "" 'x
    Chars(121) = "" 'y
    Chars(122) = "" 'z
    Chars(123) = "" '{
    Chars(124) = "" '|
    Chars(125) = "" '}
    Chars(126) = "" '~
    'used in the FormatScore function
    Chars(176) = "d_0a" '0.
    Chars(177) = "d_1a" '1.
    Chars(178) = "d_2a" '2.
    Chars(179) = "d_3a" '3.
    Chars(180) = "d_4a" '4.
    Chars(181) = "d_5a" '5.
    Chars(182) = "d_6a" '6.
    Chars(183) = "d_7a" '7.
    Chars(184) = "d_8a" '8.
    Chars(185) = "d_9a" '9.
End Sub

'****************************************
' Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub Realtime_Timer
p49on.blenddisablelighting = Light49.getinplayintensity * 40
p49off.blenddisablelighting = Light49.getinplayintensity + 1
p50on.blenddisablelighting = Light50.getinplayintensity * 40
p50off.blenddisablelighting = Light50.getinplayintensity + 1
p19on.blenddisablelighting = Light19.getinplayintensity * 40
p19off.blenddisablelighting = Light19.getinplayintensity + 1
p20on.blenddisablelighting = Light20.getinplayintensity * 40
p20off.blenddisablelighting = Light20.getinplayintensity + 1
p44on.blenddisablelighting = Light44.getinplayintensity * 40
p44off.blenddisablelighting = Light44.getinplayintensity + 1
p43on.blenddisablelighting = Light43.getinplayintensity * 40
p43off.blenddisablelighting = Light43.getinplayintensity + 1
p40on.blenddisablelighting = Light40.getinplayintensity * 40
p40off.blenddisablelighting = Light40.getinplayintensity + 1
p39on.blenddisablelighting = Light39.getinplayintensity * 40
p39off.blenddisablelighting = Light39.getinplayintensity + 1
p38on.blenddisablelighting = Light38.getinplayintensity * 40
p38off.blenddisablelighting = Light38.getinplayintensity + 1
p41on.blenddisablelighting = Light41.getinplayintensity * 40
p41off.blenddisablelighting = Light41.getinplayintensity + 1
p42on.blenddisablelighting = Light42.getinplayintensity * 40
p42off.blenddisablelighting = Light42.getinplayintensity + 1
p45on.blenddisablelighting = Light45.getinplayintensity * 40
p45off.blenddisablelighting = Light45.getinplayintensity + 1
p46on.blenddisablelighting = Light46.getinplayintensity * 40
p46off.blenddisablelighting = Light46.getinplayintensity + 1
p47on.blenddisablelighting = Light47.getinplayintensity * 40
p47off.blenddisablelighting = Light47.getinplayintensity + 1
p51on.blenddisablelighting = Light51.getinplayintensity * 40
p51off.blenddisablelighting = Light51.getinplayintensity + 1
p48on.blenddisablelighting = Light48.getinplayintensity * 40
p48off.blenddisablelighting = Light48.getinplayintensity + 1
p52on.blenddisablelighting = Light52.getinplayintensity * 40
p52off.blenddisablelighting = Light52.getinplayintensity + 1
p53on.blenddisablelighting = Light53.getinplayintensity * 40
p53off.blenddisablelighting = Light53.getinplayintensity + 1
p54on.blenddisablelighting = Light54.getinplayintensity * 40
p54off.blenddisablelighting = Light54.getinplayintensity + 1
p56on.blenddisablelighting = Light56.getinplayintensity * 40
p56off.blenddisablelighting = Light56.getinplayintensity + 1
p57on.blenddisablelighting = Light57.getinplayintensity * 40
p57off.blenddisablelighting = Light57.getinplayintensity + 1
p58on.blenddisablelighting = Light58.getinplayintensity * 40
p58off.blenddisablelighting = Light58.getinplayintensity + 1
p59on.blenddisablelighting = Light59.getinplayintensity * 40
p59off.blenddisablelighting = Light59.getinplayintensity + 1
p13on.blenddisablelighting = Light13.getinplayintensity * 40
p13off.blenddisablelighting = Light13.getinplayintensity + 1
p14on.blenddisablelighting = Light14.getinplayintensity * 40
p14off.blenddisablelighting = Light14.getinplayintensity + 1
p15on.blenddisablelighting = Light15.getinplayintensity * 40
p15off.blenddisablelighting = Light15.getinplayintensity + 1
p16on.blenddisablelighting = Light16.getinplayintensity * 40
p16off.blenddisablelighting = Light16.getinplayintensity + 1
p17on.blenddisablelighting = Light17.getinplayintensity * 40
p17off.blenddisablelighting = Light17.getinplayintensity + 1
p18on.blenddisablelighting = Light18.getinplayintensity * 40
p18off.blenddisablelighting = Light18.getinplayintensity + 1
p28on.blenddisablelighting = Light28.getinplayintensity * 40
p28off.blenddisablelighting = Light28.getinplayintensity + 1
p27on.blenddisablelighting = Light27.getinplayintensity * 40
p27off.blenddisablelighting = Light27.getinplayintensity + 1
p26on.blenddisablelighting = Light26.getinplayintensity * 40
p26off.blenddisablelighting = Light26.getinplayintensity + 1
p25on.blenddisablelighting = Light25.getinplayintensity * 40
p25off.blenddisablelighting = Light25.getinplayintensity + 1
p24on.blenddisablelighting = Light24.getinplayintensity * 40
p24off.blenddisablelighting = Light24.getinplayintensity + 1
p23on.blenddisablelighting = Light23.getinplayintensity * 40
p23off.blenddisablelighting = Light23.getinplayintensity + 1
p29on.blenddisablelighting = Light29.getinplayintensity * 40
p29off.blenddisablelighting = Light29.getinplayintensity + 1
p31on.blenddisablelighting = Light31.getinplayintensity * 40
p31off.blenddisablelighting = Light31.getinplayintensity + 1
p35on.blenddisablelighting = Light35.getinplayintensity * 40
p35off.blenddisablelighting = Light35.getinplayintensity + 1
p34on.blenddisablelighting = Light34.getinplayintensity * 40
p34off.blenddisablelighting = Light34.getinplayintensity + 1
p36on.blenddisablelighting = Light36.getinplayintensity * 40
p36off.blenddisablelighting = Light36.getinplayintensity + 1
p32on.blenddisablelighting = Light32.getinplayintensity * 40
p32off.blenddisablelighting = Light32.getinplayintensity + 1
p33on.blenddisablelighting = Light33.getinplayintensity * 40
p33off.blenddisablelighting = Light33.getinplayintensity + 1
p30on.blenddisablelighting = Light30.getinplayintensity * 40
p30off.blenddisablelighting = Light30.getinplayintensity + 1
p37on.blenddisablelighting = Light37.getinplayintensity * 40
p37off.blenddisablelighting = Light37.getinplayintensity + 1
p001on.blenddisablelighting = li001.getinplayintensity * 40
p001off.blenddisablelighting = li001.getinplayintensity + 1
p021on.blenddisablelighting = li021.getinplayintensity * 40
p021off.blenddisablelighting = li021.getinplayintensity + 1
p022on.blenddisablelighting = li022.getinplayintensity * 40
p022off.blenddisablelighting = li022.getinplayintensity + 1
p023on.blenddisablelighting = li023.getinplayintensity * 40
p023off.blenddisablelighting = li023.getinplayintensity + 1
p024on.blenddisablelighting = li024.getinplayintensity * 40
p024off.blenddisablelighting = li024.getinplayintensity + 1
p025on.blenddisablelighting = li025.getinplayintensity * 40
p025off.blenddisablelighting = li025.getinplayintensity + 1
p026on.blenddisablelighting = li026.getinplayintensity * 40
p026off.blenddisablelighting = li026.getinplayintensity + 1
p027on.blenddisablelighting = li027.getinplayintensity * 40
p027off.blenddisablelighting = li027.getinplayintensity + 1
p028on.blenddisablelighting = li028.getinplayintensity * 40
p028off.blenddisablelighting = li028.getinplayintensity + 1
p55on.blenddisablelighting = light55.getinplayintensity * 40
p55off.blenddisablelighting = light55.getinplayintensity + 1
pSAon.blenddisablelighting = LightShootAgain.getinplayintensity * 40
pSAoff.blenddisablelighting = LightShootAgain.getinplayintensity + 1
p1on.blenddisablelighting = light001.getinplayintensity * 40
p1off.blenddisablelighting = light001.getinplayintensity + 1
p2on.blenddisablelighting = light2.getinplayintensity * 40
p2off.blenddisablelighting = light2.getinplayintensity + 1
p3on.blenddisablelighting = light3.getinplayintensity * 40
p3off.blenddisablelighting = light3.getinplayintensity + 1
p4on.blenddisablelighting = light4.getinplayintensity * 40
p4off.blenddisablelighting = light4.getinplayintensity + 1
p5on.blenddisablelighting = light5.getinplayintensity * 40
p5off.blenddisablelighting = light5.getinplayintensity + 1
p6on.blenddisablelighting = light6.getinplayintensity * 40
p6off.blenddisablelighting = light6.getinplayintensity + 1
p7on.blenddisablelighting = light7.getinplayintensity * 40
p7off.blenddisablelighting = light7.getinplayintensity + 1
p8on.blenddisablelighting = light8.getinplayintensity * 40
p8off.blenddisablelighting = light8.getinplayintensity + 1
p9on.blenddisablelighting = light9.getinplayintensity * 40
p9off.blenddisablelighting = light9.getinplayintensity + 1
p10on.blenddisablelighting = light10.getinplayintensity * 40
p10off.blenddisablelighting = light10.getinplayintensity + 1
p11on.blenddisablelighting = light11.getinplayintensity * 40
p11off.blenddisablelighting = light11.getinplayintensity + 1
p12on.blenddisablelighting = light12.getinplayintensity * 40
p12off.blenddisablelighting = light12.getinplayintensity + 1

p73on.blenddisablelighting = f73.getinplayintensity * 40
p73off.blenddisablelighting = f73.getinplayintensity + 1
p74on.blenddisablelighting = f74.getinplayintensity * 40
p74off.blenddisablelighting = f74.getinplayintensity + 1
p75on.blenddisablelighting = f75.getinplayintensity * 40
p75off.blenddisablelighting = f75.getinplayintensity + 1
BSon.blenddisablelighting = Light011.getinplayintensity * 40
BSoff.blenddisablelighting = Light011.getinplayintensity + 1
LPF1on.blenddisablelighting = LP1.getinplayintensity * 40
LPF1OFF.blenddisablelighting = LP1.getinplayintensity + 1
LPF2on.blenddisablelighting = LP2.getinplayintensity * 40
LPF2OFF.blenddisablelighting = LP2.getinplayintensity + 1
LPF3on.blenddisablelighting = LP3.getinplayintensity * 40
LPF3OFF.blenddisablelighting = LP3.getinplayintensity + 1
LPF4on.blenddisablelighting = LP4.getinplayintensity * 40
LPF4OFF.blenddisablelighting = LP4.getinplayintensity + 1



    RollingUpdate
    ' add any other real time update subs, like gates or diverters
    doorp.Roty = - DoorF.CurrentAngle + 90
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
    pick001.RotX = Spinner1.CurrentAngle
    pick004.RotX = Spinner2.CurrentAngle
    FlipperRSh.RotZ = RightFlipper.CurrentAngle
	FlipperLSh.RotZ = LeftFlipper.CurrentAngle
    
End Sub

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
' 12 colors: red, orange, amber, yellow...
'******************************************

'colors
Const red = 1
Const orange = 2
Const amber = 3
Const yellow = 4
Const darkgreen = 5
Const green = 6
Const blue = 7
Const darkblue = 8
Const purple = 9
Const white = 10
Const teal = 11
Const ledwhite = 12

Sub SetLightColor(n, col, stat) 'stat 0 = off, 1 = on, 2 = blink, -1= no change
    Select Case col
        Case red
            n.color = RGB(18, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case orange
            n.color = RGB(18, 3, 0)
            n.colorfull = RGB(255, 20, 147)
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
            n.color = RGB(0, 16, 0)
            n.colorfull = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 18, 18)
            n.colorfull = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 8, 8)
            n.colorfull = RGB(0, 64, 64)
        Case purple
            n.color = RGB(64, 0, 96)
            n.colorfull = RGB(128, 0, 192)
        Case white 'bulb
            n.color = RGB(193, 91, 0)
            n.colorfull = RGB(255, 197, 143)
        Case teal
            n.color = RGB(1, 64, 62)
            n.colorfull = RGB(2, 128, 126)
        Case ledwhite
            n.color = RGB(255, 197, 143)
            n.colorfull = RGB(255, 252, 224)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

Sub SetFlashColor(n, col, stat) 'stat 0 = off, 1 = on, -1= no change - no blink for the flashers, use FlashForMs
    Select Case col
        Case red
            n.color = RGB(255, 0, 0)
        Case orange
            n.color = RGB(255, 20, 147)
        Case amber
            n.color = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 64, 64)
        Case purple
            n.color = RGB(128, 0, 192)
        Case white 'bulb
            n.color = RGB(255, 197, 143)
        Case teal
            n.color = RGB(2, 128, 126)
         Case ledwhite
            n.color = RGB(255, 252, 224)
    End Select
    If stat <> -1 Then
        n.Visible = stat
    End If
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

Sub StopRainbow()
    Dim obj
    RainbowTimer.Enabled = 0
    RainbowTimer.Enabled = 0
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

' ********************************
'   Table info & Attract Mode
' ********************************

Sub ShowTableInfo
    Dim ii
    'info goes in a loop only stopped by the credits and the startkey
    If Score(1)Then
        DMD CL("LAST SCORE"), CL("PLAYER 1 " &FormatScore(Score(1))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(2)Then
        DMD CL("LAST SCORE"), CL("PLAYER 2 " &FormatScore(Score(2))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(3)Then
        DMD CL("LAST SCORE"), CL("PLAYER 3 " &FormatScore(Score(3))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    If Score(4)Then
        DMD CL("LAST SCORE"), CL("PLAYER 4 " &FormatScore(Score(4))), "", eNone, eNone, eNone, 3000, False, ""
    End If
    DMD "", CL("GAME OVER"), "", eNone, eBlink, eNone, 2000, False, ""
    If bFreePlay Then
        DMD "", CL("FREE PLAY"), "", eNone, eBlink, eNone, 2000, False, ""
    Else
        If Credits > 0 Then
            DMD CL("CREDITS " & Credits), CL("PRESS START"), "", eNone, eBlink, eNone, 2000, False, ""
        Else
            DMD CL("CREDITS " & Credits), CL("INSERT COIN"), "", eNone, eBlink, eNone, 2000, False, ""
        End If
    End If
    DMD "", "", "d_jppresents", eNone, eNone, eNone, 3000, False, ""
    DMD "", "", "Blizzard", eNone, eNone, eNone, 4000, False, ""
    DMD "", CL("ROM VERSION " &myversion), "", eNone, eNone, eNone, 2000, False, ""
    DMD CL("HIGHSCORES"), Space(dCharsPerLine(1)), "", eScrollLeft, eScrollLeft, eNone, 20, False, ""
    DMD CL("HIGHSCORES"), "", "", eBlinkFast, eNone, eNone, 1000, False, ""
    DMD CL("HIGHSCORES"), "1> " &HighScoreName(0) & " " &FormatScore(HighScore(0)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "2> " &HighScoreName(1) & " " &FormatScore(HighScore(1)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "3> " &HighScoreName(2) & " " &FormatScore(HighScore(2)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD "_", "4> " &HighScoreName(3) & " " &FormatScore(HighScore(3)), "", eNone, eScrollLeft, eNone, 2000, False, ""
    DMD Space(dCharsPerLine(0)), Space(dCharsPerLine(1)), "", eScrollLeft, eScrollLeft, eNone, 500, False, ""
End Sub

Sub StartAttractMode
   'ChangeSong
	PuPlayer.LabelSet pDMD,"CurrScore","" ,0,""
	pDMDSetpage pScores
    StartLightSeq
	DOF 323, DOFOn
    DMDFlush
    ShowTableInfo
	bAttract = True
	AttractTimerCount = 0
	AttractTimer.Enabled = 1
	'PlaySong "Mu_End"
End Sub

Sub StopAttractMode
    DMDScoreNow
	DOF 323, DOFOff
    LightSeqAttract.StopPlay
    LightSeqFlasher.StopPlay
	bAttractMode = False
	AttractTimer.Enabled = 0
	ClearPupAttractMessages
	bAttract = False
	StopSound "Mu_End"
	pDMDSetpage pScores
End Sub

Sub StartLightSeq()
    'lights sequences
	LightSeqFlasher.StopPlay
    LightSeqFlasher.UpdateInterval = 150
    LightSeqFlasher.Play SeqRandom, 10, , 50000
    LightSeqAttract.UpdateInterval = 25
    LightSeqAttract.Play SeqBlinking, , 5, 150
    LightSeqAttract.Play SeqRandom, 40, , 4000
    LightSeqAttract.Play SeqAllOff
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 50, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 40, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 40, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqRightOn, 30, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqLeftOn, 30, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 10
    LightSeqAttract.Play SeqCircleOutOn, 15, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 5
    LightSeqAttract.Play SeqStripe1VertOn, 50, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqCircleOutOn, 15, 2
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 50, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 25, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe1VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqStripe2VertOn, 25, 3
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqUpOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqDownOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqRightOn, 15, 1
    LightSeqAttract.UpdateInterval = 8
    LightSeqAttract.Play SeqLeftOn, 15, 1
End Sub

Sub LightSeqAttract_PlayDone()
    StartLightSeq()
End Sub

Sub LightSeqTilt_PlayDone()
    LightSeqTilt.Play SeqAllOff
End Sub

Sub LightSeqSkillshot_PlayDone()
    LightSeqSkillshot.Play SeqAllOff
End Sub

'************************************
'       LUT - Darkness control
' 10 normal level & 10 warmer levels 
'************************************

Dim bLutActive, LUTImage

Sub LoadLUT
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
	'Debug.print "Image:" &LUTImage
End Sub

Sub NextLUT
	bLutActive = True
	UpdateLUT
	SaveLUT
	SetLUTLine "Color LUT image " & table1.ColorGradeImage
End Sub

'Sub NextLUT:bLutActive = True:UpdateLUT:SaveLUT:SetLUTLine "Color LUT image " & table1.ColorGradeImage:End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT6"
        Case 7:table1.ColorGradeImage = "LUT7"
        Case 8:table1.ColorGradeImage = "LUT8"
        Case 9:table1.ColorGradeImage = "LUT9"
        Case 10:table1.ColorGradeImage = "LUT10"
        Case 11:table1.ColorGradeImage = "LUT Warm 0"
        Case 12:table1.ColorGradeImage = "LUT Warm 1"
        Case 13:table1.ColorGradeImage = "LUT Warm 2"
        Case 14:table1.ColorGradeImage = "LUT Warm 3"
        Case 15:table1.ColorGradeImage = "LUT Warm 4"
        Case 16:table1.ColorGradeImage = "LUT Warm 5"
        Case 17:table1.ColorGradeImage = "LUT Warm 6"
        Case 18:table1.ColorGradeImage = "LUT Warm 7"
        Case 19:table1.ColorGradeImage = "LUT Warm 8"
        Case 20:table1.ColorGradeImage = "LUT Warm 9"
        Case 21:table1.ColorGradeImage = "LUT Warm 10"
        Case 22:table1.ColorGradeImage = "Fleep Natural Dark 1"
'        Case 23:table1.ColorGradeImage = "Fleep Natural Dark 2"
'        Case 24:table1.ColorGradeImage = "Fleep Warm Dark"
'        Case 25:table1.ColorGradeImage = "Fleep Warm Bright"
'        Case 26:table1.ColorGradeImage = "Fleep Warm Vivid Soft"
'        Case 27:table1.ColorGradeImage = "Fleep Warm Vivid Hard"
'        Case 28:table1.ColorGradeImage = "Skitso Natural and Balanced"
'        Case 29:table1.ColorGradeImage = "Skitso Natural High Contrast"
'        Case 30:table1.ColorGradeImage = "3rdaxis Referenced THX Standard"
'        Case 31:table1.ColorGradeImage = "CalleV Punchy Brightness and Contrast"
'        Case 32:table1.ColorGradeImage = "HauntFreaks Desaturated"
'        Case 33:table1.ColorGradeImage = "Tomate Washed Out"
'        Case 34:table1.ColorGradeImage = "VPW Original 1 to 1"
'        Case 35:table1.ColorGradeImage = "Bassgeige"
'        Case 36:table1.ColorGradeImage = "Blacklight"
'        Case 37:table1.ColorGradeImage = "B&W Comic Book"
    End Select
End Sub

Dim GiIntensity
GiIntensity = 1   'can be used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

' New LUT postit
Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetLUTLine(String)
    Dim Index
    Dim xFor
    Index = 1
    LUBack.imagea="PostItNote"
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
SetLUTLine ""
LUBack.imagea="PostitBL"
End Sub
'***********************************************************************
' *********************************************************************
'                     Table Specific Script Starts Here
' *********************************************************************
'***********************************************************************

' droptargets, animations, etc
Sub VPObjects_Init
End Sub

' tables variables and Mode init

Dim LaneBonus
Dim TargetBonus
Dim RampBonus
Dim ComboCount
Dim ComboHits(4)
Dim ComboValue(4)
Dim BumperValue(4)
Dim BumperHits
Dim SuperBumperHits
Dim SpinnerValue(4)
Dim MonstersKilled(4)
Dim BLIZZARDMB(4)
Dim SpinCount
Dim RampHits3
Dim RampHits12
Dim OrbitHits
Dim TargetHits7
Dim TargetHits8
Dim CaptiveBallHits
Dim LightHits6
Dim LightHits9
Dim LightHits11
Dim loopCount
Dim BattlesWon(4)
Dim Battle(4, 15) '12 battles, 1 final battle
Dim NewBattle
Dim PowerupHits

Sub Game_Init() 'called at the start of a new game
    Dim i, j
    bExtraBallWonThisBall = False
    'Play some Music
    'ChangeSong
    'Init Variables
    LaneBonus = 0 'it gets deleted when a new ball is launched
    TargetBonus = 0
    RampBonus = 0
    BumperHits = 0
    InitBobble
    For i = 1 to 4
        SkillshotValue(i) = 500000
        Jackpot(i) = 150000
        ComboValue(i) = 500000
        MonstersKilled(i) = 0
        BallsInLock(i) = 0
        SpinnerValue(i) = 1000
        ComboHits(i) = 0
        BumperValue(i) = 210 'start at 210 and every 30 hits its value is increased by 500 points
    Next
    ResetBattles
    SpinCount = 0
    SuperBumperHits = 0
    RampHits3 = 0
    RampHits12 = 0
    OrbitHits = 0
    TargetHits7 = 0
    TargetHits8 = 0
    CaptiveBallHits = 0
    loopCount = 0
    PowerupHits = 0
    LightHits9 = 0
    LightHits11 = 0
    ComboCount = 0
    'Init Delays/Timers
    'MainMode Init()
    'Init lights
    TurnOffPlayfieldLights()
End Sub

Sub StopEndOfBallMode() 'this sub is called after the last ball is drained
    ResetSkillShotTimer_Timer
    StopBattle
End Sub

Sub ResetNewBallVariables() 'reset variables for a new ball or player
    Dim i

	' Reset Extraball light
	Light39.State = 0

	ResetLaneLights
    LaneBonus = 0
    TargetBonus = 0
    RampBonus = 0
    BumperHits = 0
    ' select a battle
    SelectBattle
    RiseTarget
    bBLIZZARDMode = False
    UpdateLockLights
    ResetBLIZZARDLights
	nBlizzTargetCount = 0
	bSupressModeMessages = False	
	bSupressModeProgress = False

	for i = 0 to 8
		BlizzardLetters(i) = 0
	Next
    
      'reset LaneSwitch variables
    For i = 0 to 4
        LaneSwitch(i) = 0
	next

	DMDUpdateBallNumber Balls
	DMDUpdatePlayerName
	
End Sub

Sub ResetNewBallLights() 'turn on or off the needed lights before a new ball is released
 UpdatePFXLights(PlayfieldMultiplier(CurrentPlayer)) 'ensure the multiplier is displayed right
Light13.State = 0
Light16.State = 0

End Sub

Sub ResetBLIZZARDLights
debug.print "rest blizz lights"
	Li001.State = 0 
	Li021.State = 0
	Li022.State = 0
	Li023.State = 0
	Li024.State = 0
	Li025.State = 0
	Li026.State = 0
	Li027.State = 0
	Li028.State = 0
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub UpdateSkillShot() 'Setup and updates the skillshot lights
    LightSeqSkillshot.Play SeqAllOff
    Light48.State = 2
    Light17.State = 2
	LaneSwitch(1) = 2
    Gate2.Open = 1
    Gate3.Open = 1
    DMD CL("HIT LIT LIGHT"), CL("FOR SKILLSHOT"), "", eNone, eNone, eNone, 1500, True, ""
End Sub

Sub ResetSkillShotTimer_Timer 'timer to reset the skillshot lights & variables
    ResetSkillShotTimer.Enabled = 0
    bSkillShotReady = False
    LightSeqSkillshot.StopPlay
    If Light17.State = 2 Then Light17.State = 0
    Light48.State = 0
    Gate2.Open = 0
    Gate3.Open = 0
    DMDScoreNow
End Sub

'*********************************************************
' Slingshots has been hit
' In this table the slingshots change the outlanes lights

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    ShakeLeftCat
    Playsound "bats"
    FlashForMs AP1, 1000, 50, 0
    FlashForMs AP2, 1000, 50, 0
    If Spots.Enabled = 0 Then FlashForms spot1, 1000, 50, 0:FlashForms spot2, 1000, 50, 0
    If Tilted Then Exit Sub
    'PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFcontactors), Lemk
	RandomSoundSlingShotLeft Lemk
    DOF 105, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add some points
    AddScore 210
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "LeftSlingShot"
    ChangeOutlanes
    

	
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    ShakeRightCat
    Playsound "bats-2"
    FlashForMs AP1, 1000, 50, 0
    FlashForMs AP2, 1000, 50, 0
    If Spots.Enabled = 0 Then FlashForms spot1, 1000, 50, 0:FlashForms spot2, 1000, 50, 0
    If Tilted Then Exit Sub
    'PlaySoundAt SoundFXDOF("fx_slingshot", 104, DOFPulse, DOFcontactors), Remk
	RandomSoundSlingShotRight Remk
    DOF 106, DOFPulse
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add some points
    AddScore 210
    ' add some effect to the table?
    ' remember last trigger hit by the ball
    LastSwitchHit = "RightSlingShot"
    ChangeOutlanes

	
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


Sub ChangeOutlanes
    Dim tmp
    tmp = light13.State
    light13.State = light16.State
    light16.State = tmp
End Sub

'*********
' Bumpers
'*********
' after each 30 hits the bumpers increase their score value by 500 points up to 3210
' and they increase the playfield multiplier.

Sub Bumper1_Hit
    If NOT Tilted Then
'        PlaySoundAt SoundFXDOF("Bumpers_Top_1", 109, DOFPulse, DOFContactors), Bumper1
'        FlashForMs bumpersmalllight1, 1000, 50, 0
	RandomSoundBumperTop Bumper1
	FlBumperFadeTarget(1) = 1   'Flupper bumper demo
	Bumper1.timerenabled = True
        DOF 138, DOFPulse
        ' add some points
        AddScore BumperValue(CurrentPlayer)
        If Battle(CurrentPlayer, 0) = 2 Then
            SuperBumperHits = SuperBumperHits + 1
            Addscore 5000
            CheckWinBattle
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper1"
    End If
    CheckBumpers
End Sub

Sub Bumper1_Timer
	FlBumperFadeTarget(1) = 0
End Sub

Sub Bumper2_Hit
    If NOT Tilted Then
'        PlaySoundAt SoundFXDOF("Bumpers_Top_2", 110, DOFPulse, DOFContactors), Bumper2
'        FlashForMs bumpersmalllight2, 1000, 50, 0
	RandomSoundBumperMiddle Bumper2
	FlBumperFadeTarget(2) = 1   'Flupper bumper demo
	Bumper2.timerenabled = True
        DOF 140, DOFPulse
        ' add some points
        AddScore BumperValue(CurrentPlayer)
        If Battle(CurrentPlayer, 0) = 2 Then
            SuperBumperHits = SuperBumperHits + 1
            Addscore 5000
            CheckWinBattle
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper2"
    End If
    CheckBumpers
End Sub

Sub Bumper2_Timer
	FlBumperFadeTarget(2) = 0
End Sub

Sub Bumper3_Hit
    If NOT Tilted Then
'        PlaySoundAt SoundFXDOF("Bumpers_Top_3", 107, DOFPulse, DOFContactors), Bumper3
'        FlashForMs bumpersmalllight3, 1000, 50, 0
	RandomSoundBumperBottom Bumper3
	FlBumperFadeTarget(3) = 1   'Flupper bumper demo
	Bumper3.timerenabled = True

        DOF 137, DOFPulse
        ' add some points
        AddScore BumperValue(CurrentPlayer)
        If Battle(CurrentPlayer, 0) = 2 Then
            SuperBumperHits = SuperBumperHits + 1
            Addscore 5000
            CheckWinBattle
        End If
        ' remember last trigger hit by the ball
        LastSwitchHit = "Bumper3"
    End If
    CheckBumpers
End Sub

Sub Bumper3_Timer
	FlBumperFadeTarget(3) = 0
End Sub

' Check the bumper hits

Sub CheckBumpers()
    ' increase the bumper hit count and increase the bumper value after each 30 hits
    BumperHits = BumperHits + 1
    If BumperHits MOD 30 = 0 Then
        If BumperValue(CurrentPlayer) < 3210 Then
            BumperValue(CurrentPlayer) = BumperValue(CurrentPlayer) + 500
        End If
        ' lit the playfield multiplier light
        light54.State = 1
        TimerJAG.Enabled = True
        FlashForMs BLight001, 1000, 50, 0
        Playsound "WW2"
        PlayPlayfieldMP
    End If
End Sub

'************Primitive Swap + Shake*******************

Wolf001.Visible = True
Wolf002.Visible = False

' Initialize the timer
TimerJAG.Interval = 3
TimerJAG.Enabled = False
Dim Timer1Count: Timer1Count = 0
Dim ShakePhase: ShakePhase = 0

Sub TimerJAG_Timer
    Timer1Count = Timer1Count + 1

    ' --- Small shake effect ---
    If Timer1Count < 30 Then    ' shake for first few frames
        ShakePhase = ShakePhase + 1
        Dim shakeAmount
        shakeAmount = 5 * Sin(ShakePhase * 0.8) ' adjust for strength and speed
        Wolf001.TransX = shakeAmount
        Wolf002.TransX = shakeAmount
    Else
        ' reset shake after first 30 ticks
        Wolf001.TransX = 0
        Wolf002.TransX = 0
    End If

    ' --- Swap frames ---
    Select Case Timer1Count
        Case 1
            Wolf001.Visible = False
            Wolf002.Visible = True
        Case 100 ' number of timer intervals before swapping back
            Wolf001.Visible = True
            Wolf002.Visible = False
            Timer1Count = 0
            ShakePhase = 0
            Wolf001.TransX = 0
            Wolf002.TransX = 0
            TimerJAG.Enabled = False
    End Select
End Sub

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
	DNA30 = 0 : DNA45 = (NightDay-10)/20 : DNA90 = 0 : DayNightAdjust = 0.4
Else
	DNA30 = (NightDay-10)/30 : DNA45 = (NightDay-10)/45 : DNA90 = (NightDay-10)/90 : DayNightAdjust = NightDay/25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt : For cnt = 1 to 6 : FlBumperActive(cnt) = False : Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight

FlInitBumper 1, "red"
FlInitBumper 2, "blue"
FlInitBumper 3, "yellow"

' ### uncomment the statement below to change the color for all bumpers ###
' Dim ind : For ind = 1 to 5 : FlInitBumper ind, "green" : next

Sub FlInitBumper(nr, col)
	FlBumperActive(nr) = True
	' store all objects in an array for use in FlFadeBumper subroutine
	FlBumperFadeActual(nr) = 1 : FlBumperFadeTarget(nr) = 1.1: FlBumperColor(nr) = col
	Set FlBumperTop(nr) = Eval("bumpertop" & nr) : FlBumperTop(nr).material = "bumpertopmat" & nr
	Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr) : Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
	Set FlBumperDisk(nr) = Eval("bumperdisk" & nr) : Set FlBumperBase(nr) = Eval("bumperbase" & nr)
	Set FlBumperBulb(nr) = Eval("bumperbulb" & nr) : FlBumperBulb(nr).material = "bumperbulbmat" & nr
	Set FlBumperscrews(nr) = Eval("bumperscrews" & nr): FlBumperscrews(nr).material = "bumperscrew" & col
	Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
	' set the color for the two VPX lights
	select case col
		Case "red"
			FlBumperSmallLight(nr).color = RGB(255,4,0) : FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
			FlBumperBigLight(nr).color = RGB(255,32,0) : FlBumperBigLight(nr).colorfull = RGB(255,32,0)
			FlBumperHighlight(nr).color = RGB(64,255,0)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
			FlBumperSmallLight(nr).TransmissionScale = 0
		Case "blue"
			FlBumperBigLight(nr).color = RGB(32,80,255) : FlBumperBigLight(nr).colorfull = RGB(32,80,255)
			FlBumperSmallLight(nr).color = RGB(0,80,255) : FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
			FlBumperSmallLight(nr).TransmissionScale = 0 : MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
			FlBumperHighlight(nr).color = RGB(255,16,8)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
		Case "green"
			FlBumperSmallLight(nr).color = RGB(8,255,8) : FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
			FlBumperBigLight(nr).color = RGB(32,255,32) : FlBumperBigLight(nr).colorfull = RGB(32,255,32)
			FlBumperHighlight(nr).color = RGB(255,32,255) : MaterialColor "bumpertopmat" & nr, RGB(16,255,16) 
			FlBumperSmallLight(nr).TransmissionScale = 0.005
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
		Case "orange"
			FlBumperHighlight(nr).color = RGB(205,170,109)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1 
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).color = RGB(205,170,109) : FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
			FlBumperBigLight(nr).color = RGB(194,178,128) : FlBumperBigLight(nr).colorfull = RGB(205,170,109)
		Case "white"
			FlBumperBigLight(nr).color = RGB(255,230,190) : FlBumperBigLight(nr).colorfull = RGB(255,230,190)
			FlBumperHighlight(nr).color = RGB(255,180,100) : 
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
		Case "blacklight"
			FlBumperBigLight(nr).color = RGB(32,32,255) : FlBumperBigLight(nr).colorfull = RGB(32,32,255)
			FlBumperHighlight(nr).color = RGB(48,8,255) : 
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
		Case "yellow"
			FlBumperSmallLight(nr).color = RGB(255,230,4) : FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
			FlBumperBigLight(nr).color = RGB(255,240,50) : FlBumperBigLight(nr).colorfull = RGB(255,240,50)
			FlBumperHighlight(nr).color = RGB(255,255,220)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1 
			FlBumperSmallLight(nr).TransmissionScale = 0
		Case "purple"
			FlBumperBigLight(nr).color = RGB(80,32,255) : FlBumperBigLight(nr).colorfull = RGB(80,32,255)
			FlBumperSmallLight(nr).color = RGB(80,32,255) : FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
			FlBumperSmallLight(nr).TransmissionScale = 0 : 
			FlBumperHighlight(nr).color = RGB(32,64,255)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
	end select
End Sub

Sub FlFadeBumper(nr, Z)
	FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
'	UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
'               OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
'               float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
	FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 )* DayNightAdjust	

	select case FlBumperColor(nr)

		Case "blue" :
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(38-24*Z,130 - 98*Z,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20  + 500 * Z / (0.5 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
			FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z +0.97 * Z^3)
			Flbumperbiglight(nr).intensity = 45 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 10000 * (Z^3) / (0.5 + DNA90)

		Case "green"	
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(16 + 16 * sin(Z*3.14),255,16 + 16 * sin(Z*3.14)), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
			FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 6000 * (Z^3) / (1 + DNA90)
		
		Case "red" 
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 16 - 11*Z + 16 * sin(Z*3.14),0), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30^2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z*4,8-Z*8)     
		
		Case "orange"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(205, 170 - 22*z  + 16 * sin(Z*3.14),Z*32), RGB(205,205,205), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30^2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,136, 50)

		Case "white"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
			FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 14 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
			FlBumperSmallLight(nr).color = RGB(255,255 - 20*Z,255-65*Z) : FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20*Z,255-65*Z)
			MaterialColor "bumpertopmat" & nr, RGB(255,235 - z*36,220 - Z*90)

		Case "blacklight"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 1, RGB(30-27*Z^0.03,30-28*Z^0.01, 255), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z^3
			Flbumperbiglight(nr).intensity = 40 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
			FlBumperSmallLight(nr).color = RGB(255-240*(Z^0.1),255 - 240*(Z^0.1),255) : FlBumperSmallLight(nr).colorfull = RGB(255-200*z,255 - 200*Z,255)
			MaterialColor "bumpertopmat" & nr, RGB(255-190*Z,235 - z*180,220 + 35*Z)

		Case "yellow"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 180 + 40*z, 48* Z), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30^2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z +0.97 * Z^10)
			Flbumperbiglight(nr).intensity = 20 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

		Case "purple" :
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(128-118*Z - 32 * sin(Z*3.14), 32-26*Z ,255), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 15  + 200 * Z / (0.5 + DNA30) 
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z +0.97 * Z^3)
			Flbumperbiglight(nr).intensity = 50 * Z / (1 + DNA45) 
			FlBumperHighlight(nr).opacity = 4000 * (Z^3) / (0.5 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(128-60*Z,32,255)


	end select
End Sub

Sub BumperTimer_Timer
	dim nr
	For nr = 1 to 6
		If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
			FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
			If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1 : End If
			FlFadeBumper nr, FlBumperFadeActual(nr)
		End If
		If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
			FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
			If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0 : End If
			FlFadeBumper nr, FlBumperFadeActual(nr)
		End If
	next
End Sub

'*************************
' Top & Inlanes: Bonus X
'*************************
' lit the 2 top lane lights and the 2 inlane lights to increase the bonus multiplier

Sub sw1_Hit
    DOF 128, DOFPulse
    'PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
	Addscore SCORE_LANES
    Light17.State = 1
    FlashForMs f8, 1000, 50, 0
    LaneSwitch(2) = 1:CheckLaneSwitch
            Light17.State = LaneSwitch(2)
    If bSkillShotReady Then
     Awardskillshot
        
    Else
        CheckBonusX
    End If
End Sub

Sub sw6_Hit
    DOF 129, DOFPulse
    'PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
	Addscore SCORE_LANES
    Light18.State = 1
    FlashForMs f8, 1000, 50, 0
    LaneSwitch(2) = 1:CheckLaneSwitch
            Light18.State = LaneSwitch(2)
    If bSkillShotReady Then
    ResetSkillShotTimer_Timer    
    Else
        CheckBonusX
    End If
End Sub

Sub sw4_Hit
    DOF 133, DOFPulse
    'PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
	Addscore SCORE_LANES
    Light14.State = 1
    FlashForMs f6, 1000, 50, 0
    LaneSwitch(3) = 1:CheckLaneSwitch
            Light14.State = LaneSwitch(3)
    AddScore 5000
    CheckBonusX

' Do some sound or light effect
End Sub

Sub sw3_Hit
    DOF 134, DOFPulse
    'PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
	Addscore SCORE_LANES
    Light15.State = 1
    FlashForMs f7, 1000, 50, 0
    LaneSwitch(4) = 1:CheckLaneSwitch
            Light15.State = LaneSwitch(4)
    AddScore 5000
    CheckBonusX

' Do some sound or light effect
End Sub


Sub CheckLaneSwitch
    Dim tmp
    tmp = LaneSwitch(1) + LaneSwitch(2) + LaneSwitch(3) + LaneSwitch(4)
    If tmp = 4 Then
        AddBonusMultiplier 1
        LaneSwitch(1) = 0:Light17.State = 0
        LaneSwitch(2) = 0:Light18.State = 0
        LaneSwitch(3) = 0:Light14.State = 0
        LaneSwitch(4) = 0:Light15.State = 0
        LightEffect 2
    End If
End Sub

Sub RotatelaneSwitchLeft
	if bSkillshotReady Then Exit Sub
    Dim tmp
    tmp = LaneSwitch(1)
    LaneSwitch(1) = LaneSwitch(2)
    LaneSwitch(2) = LaneSwitch(3)
    LaneSwitch(3) = LaneSwitch(4)
    LaneSwitch(4) = tmp
    Light17.State = LaneSwitch(1)
    Light18.State = LaneSwitch(2)
    Light15.State = LaneSwitch(3)
    Light14.State = LaneSwitch(4)
End Sub

Sub RotateLaneSwitchRight
	if bSkillshotReady Then Exit Sub
    Dim tmp
	tmp = LaneSwitch(1)
    LaneSwitch(1) = LaneSwitch(4)
	LaneSwitch(4)= LaneSwitch(3)
    LaneSwitch(3) = LaneSwitch(2)
    LaneSwitch(2) = tmp

    Light17.State = LaneSwitch(1)
    Light18.State = LaneSwitch(2)
    Light15.State = LaneSwitch(3)
    Light14.State = LaneSwitch(4)
   
End Sub

Sub ResetLaneLights
    Light17.State = 0
    Light18.State = 0
    Light14.State = 0
    Light15.State = 0
End Sub

Sub CheckBonusX
    If Light17.State + Light18.State + Light14.State + Light15.State = 4 Then
        AddBonusMultiplier 1
        GiEffect 1
        FlashForMs Light17, 1000, 50, 0
        FlashForMs Light18, 1000, 50, 0
        FlashForMs Light14, 1000, 50, 0
        FlashForMs Light15, 1000, 50, 0
        
    End IF
End Sub


'************************************
' Flipper OutLanes: Virtual kickback
'***************************************
' if the light is lit then activate the ballsave

Sub sw2_Hit
    DOF 132, DOFPulse
    'PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    AddScore 50000
    ' Do some sound or light effect
    ' do some check
    If light13.State = 1 Then
        EnableBallSaver 5
    End If

'	debug.print "OUT:" &gametime
'	debug.print "BS:" &bBallSaverActive

	LastSwitchHit = "LeftOutlane"
End Sub

Sub sw5_Hit
    DOF 135, DOFPulse
    'PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
    LaneBonus = LaneBonus + 1
    AddScore 50000
    ' Do some sound or light effect
    ' do some check
    If Light16.State = 1 Then
        EnableBallSaver 5
    End If

	LastSwitchHit = "RightOutlane"
End Sub

'******DROPTARGET******



Sub RiseTarget
    RandomSoundDropTargetReset Target009
    StopSpinner
    Target009.IsDropped = 0
	TargetResetTimer.Enabled = False
	HideBlizzTimer
    ' disable ramp locks
End Sub

' Stuck Ball Rescue for Target009 

Dim Target009BallTime
Target009BallTime = 0

' Trigger when ball enters behind target
Sub Trigger_Target009_Back_Hit()
    Target009BallTime = GameTime
    Target009Check.Enabled = True
End Sub

' Trigger when ball leaves behind target
Sub Trigger_Target009_Back_UnHit()
    Target009Check.Enabled = False
    Target009BallTime = 0
End Sub

' Timer to check for stuck ball
Sub Target009Check_Timer()
    If Target009BallTime > 0 Then
        If GameTime - Target009BallTime > 1500 Then ' 1.5 seconds stuck
            Target009.IsDropped = True               ' Drop target
            RaiseTargetTimer.Enabled = True         ' Enable raise target timer
            Target009Check.Enabled = False
            Target009BallTime = 0
        End If
    End If
End Sub

' Timer to raise target after delay
Sub RaiseTargetTimer_Timer()

    Target009.IsDropped = 0
    StopSpinner
    RandomSoundDropTargetReset Target009
    RaiseTargetTimer.Enabled = False
	TargetResetTimer.Enabled = False
    If Target009.IsDropped = False Then
    If(bBLIZZARDMode = False) Then RiseTarget:ResetBLIZZARDLights
    ResetCountdown
    End If
    
End Sub



'*******************
'BLIZZARD Multiball
'*******************

Sub Target009_Hit()
    SoundDropTargetDrop Target009
	TargetBonus = TargetBonus + 1
	DOF 120, DOFPulse
	StartSpinner

	if Battle(CurrentPlayer, 0) <> 0 OR bMultiBallMode Then Exit Sub
	

    StartCountdown
	TargetResetTimer.Enabled = True

	bBlizzardPrepMode = True
	pdmdlabelshow "BlizzTimerImage"
    PlayBlizzardTargetCall
'	GeneralPupQueue.Add "PlayBlizzardCollectVideo","PlayBlizzardCollectVideo",60,1000,0,0,0,False

    DMD "_", CL("HIT 8 TARGETS"), "", eNone, eNone, eNone, 5000, True, "" 

    if Battle(CurrentPlayer, 0) = 0 And bMultiBallMode = False And Li001.State = 0 Then Li001.State = 1':CheckBLIZZARDTargets

    LastSwitchHit = "Target009"
End Sub

Sub Target09_Hit   
	DOF 120, DOFPulse
	PlayTargetSound
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target09"
    ' Do some sound or light effect
    if Battle(CurrentPlayer, 0) = 0 And bMultiBallMode = False And Li021.State = 2 Then Li021.State = 1:BlizzardLetters(1) = 1:CheckBLIZZARDTargets
    ' do some check
End Sub

Sub Target010_Hit   
	DOF 120, DOFPulse
	PlayTargetSound
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target010"
    ' Do some sound or light effect
    if Battle(CurrentPlayer, 0) = 0 And bMultiBallMode = False And Li022.State = 2 Then Li022.State = 1:BlizzardLetters(2) = 1:CheckBLIZZARDTargets
    ' do some check
End Sub

Sub Target011_Hit   
	DOF 120, DOFPulse
	PlayTargetSound
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target011"
    ' Do some sound or light effect
    if Battle(CurrentPlayer, 0) = 0 And bMultiBallMode = False And Li023.State = 2 Then Li023.State = 1:BlizzardLetters(3) = 1:CheckBLIZZARDTargets
    ' do some check
End Sub

Sub Target012_Hit   
	DOF 120, DOFPulse
	PlayTargetSound
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target012"
    ' Do some sound or light effect
    if Battle(CurrentPlayer, 0) = 0 And bMultiBallMode = False And Li024.State = 2 Then Li024.State = 1:BlizzardLetters(4) = 1:CheckBLIZZARDTargets
    ' do some check
End Sub

Sub Target013_Hit   
	DOF 120, DOFPulse
	PlayTargetSound
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target013"
    ' Do some sound or light effect
    if Battle(CurrentPlayer, 0) = 0 And bMultiBallMode = False And Li025.State = 2 Then Li025.State = 1:BlizzardLetters(5) = 1:CheckBLIZZARDTargets
    ' do some check
End Sub

Sub Target014_Hit   
	DOF 120, DOFPulse
	PlayTargetSound
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target014"
    ' Do some sound or light effect
    if Battle(CurrentPlayer, 0) = 0 And bMultiBallMode = False And Li026.State = 2 Then Li026.State = 1:BlizzardLetters(6) = 1:CheckBLIZZARDTargets
    ' do some check
End Sub

Sub Target015_Hit   
	DOF 120, DOFPulse
	PlayTargetSound
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target015"
    ' Do some sound or light effect
    if Battle(CurrentPlayer, 0) = 0 And bMultiBallMode = False And Li027.State = 2 Then Li027.State = 1:BlizzardLetters(7) = 1:CheckBLIZZARDTargets
    ' do some check
End Sub

Sub Target016_Hit  
	DOF 120, DOFPulse
	PlayTargetSound 
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target016"
    ' Do some sound or light effect
    if Battle(CurrentPlayer, 0) = 0 And bMultiBallMode = False And Li028.State = 2 Then Li028.State = 1:BlizzardLetters(8) = 1:CheckBLIZZARDTargets
    ' do some check
End Sub


Sub CheckBLIZZARDTargets

	If bBlizzardPrepMode = False Then Exit Sub

    Dim tmp
    FlashForMs f1, 1000, 50, 0
    FlashForMs f3, 1000, 50, 0
    FlashForMs f4, 1000, 50, 0
    FlashForMs f4, 1000, 50, 0
'    PlaySoundAtBall "" &tmp
    ' if all 8 targets are hit then start multiball & activate the multiball light

	nBlizzTargetCount = nBlizzTargetCount + 1

 '   If li021.state + li022.state + li023.state + li024.state + li025.state + li026.state + li027.state + li028.state = 8 Then
	dim i,tot

	tot = 0
	for i = 1 to 8
		tot = tot + BlizzardLetters(i)
	next

	if tot = 8 Then
	
        ' BLIZZARD Multiball
        BLIZZARDMB(CurrentPlayer) = BLIZZARDMB(CurrentPlayer) + 1
		HideBlizzTimer
'		bBLIZZARDMode = True
		bMultiBallMode = True
		bBlizzardPrepMode = False
        ResetCountdown
        StartBliz
        StartSnow
        TargetResetTimer.Enabled = False
        DMD " BLIZZARD", " MULTIBALL", "blizzardmb" &tmp, eNone, eNone, eNone, 2500, True, ""
        StartIceGrow
        PlayBlizzardMBall
		GeneralPupQueue.Add "PlayBlizzardCollectVideo","PlayBlizzardCollectVideo",60,100,0,0,0,False
        Playsound "wind"
		EnableBlizzardTargetLights
        BackleftFlash
        BackRightFlash
        RightFlash
        LeftFlash
        LightEffect 1
        FlashEffect 1
		'bAutoPlunger = True
		addmultiball 6
		bJackpot = True
        light44.State = 2
        Light49.State = 2

		PuPlayer.LabelSet pDMD,"Event3A","",0,""
		PuPlayer.LabelSet pDMD,"Event3B","",0,""
		PuPlayer.LabelSet pDMD,"Event3Ca","",0,""
		PuPlayer.LabelSet pDMD,"Event3C","",0,""

' Turn off ball saver
		'bBallSaverActive = False
		enableballsaver 5
		LightShootAgain.State = 0
        ' Lit the BLIZZARD MB light if it is off
		AddScore 50000
        ' reset the lights

		if Scorbit.bSessionActive then
			GameModeStrTmp="NA{Yellow}:Blizzard MB "
			if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
		End If

    End If
End Sub

'************
'  Spinners
'************

Sub spinner1_Spin
    If Tilted Then Exit Sub
    Addscore spinnervalue(CurrentPlayer)
    'PlaySoundAt "fx_spinner", spinner1
	SoundSpinner spinner1
    DOF 136, DOFPulse
    Select Case Battle(CurrentPlayer, 0)
        Case 1
            Addscore 3000
            SpinCount = SpinCount + 1
			if NOT bSupressModeMessages Then UpdateSpinProgress
            CheckWinBattle
    End Select
End Sub

Sub spinner2_Spin
    If Tilted Then Exit Sub
    'PlaySoundAt "fx_spinner", spinner2
	SoundSpinner spinner2
    DOF 124, DOFPulse
    Addscore spinnervalue(CurrentPlayer)
    Select Case Battle(CurrentPlayer, 0)
        Case 1
            Addscore 3000
            SpinCount = SpinCount + 1
			if NOT bSupressModeMessages Then UpdateSpinProgress
            CheckWinBattle
    End Select
End Sub

'*******************
'    RAMP COMBOS
'*******************

Sub AwardCombo
    Playsound "van"
    BackleftFlash
    BackRightFlash
    FlashForMs VANLIGHT, 1000, 50, 0
  
    ComboCount = ComboCount + 1
	UpdateComboCount

    Select Case ComboCount
        Case 1
            MoveVansBackAndForth
            DMD CL("COMBO"), CL(FormatScore(ComboValue(CurrentPlayer))), "", eNone, eNone, eNone, 1500, True, ""
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            PlayCombo
        Case 2
            MoveVansBackAndForth
            DMD CL("DOUBLE COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 2)), "", eNone, eNone, eNone, 1500, True, ""
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            Play2XCombo
        Case 3
            MoveVansBackAndForth
            DMD CL("TRIPLE COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 3)), "", eNone, eNone, eNone, 1500, True, ""
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            Play3XCombo
        Case 4
            MoveVansBackAndForth
            DMD CL("SCARY COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 4)), "", eNone, eNone, eNone, 1500, True, ""
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            Play4XCombo
        Case 5
            MoveVansBackAndForth
            DMD CL("UNHOLY COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 5)), "", eNone, eNone, eNone, 1500, True, ""
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            Play5XCombo
        Case 6
            MoveVansBackAndForth
            DMD CL("SUPER COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 7)), "", eNone, eNone, eNone, 1500, True, ""
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            PlaySuperCombo
            DOF 126, DOFPulse
        Case 7
            MoveVansBackAndForth
			if NOT bExtraBallWonThisBall And Light39.State <> 2 Then
				Light39.State = 2
				' Wait for super duper combo audio to play
				AudioQueue.Add "PlayExtraBallisLit","PlayExtraBallisLit",65,1800,0,0,0,False	
			End If
            VansTiltForward       ' Tilt animation for Mr Crowely combo
            PuPlayer.playevent pDMDVideo,"Combos","Combo_2.mp4",nPupVideoVolume,67,3,0,""
            DMD CL("MR CROWELY COMBO"), CL(FormatScore(ComboValue(CurrentPlayer) * 10)), "", eNone, eNone, eNone, 1500, True, ""
            ComboHits(CurrentPlayer) = ComboHits(CurrentPlayer) + 1
            PlaySuperDuperCombo
            DOF 126, DOFPulse
    End Select

    AddScore ComboValue(CurrentPlayer) * ComboCount
	Tracker_ComboValue = Tracker_ComboValue + (ComboValue(CurrentPlayer) * ComboCount)
    ComboValue(CurrentPlayer) = ComboValue(CurrentPlayer) + 100000
End Sub

Sub aComboTargets_Hit(idx) 'reset the combo count if the ball hits another target/trigger
    ComboCount = 0
	UpdateComboCount
End Sub

'*********************************
'      The Lock Targets
'*********************************

Sub Target13_Hit
	DOF 116, DOFPulse
	PlayTargetSound
    PlaySoundAt SoundFXDOF("", 116, DOFPulse, DOFTargets), Target10
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    ' Do some sound or light effect
    Light19.State = 1
    FlashForMs f4, 1000, 50, 0
    ' do some check
    Check2BankTargets
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Mode5Lights(3) = 2 Then
				Mode5Lights(3) = 0
				CalcMode5Lights
                Light31.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light31.State = 2 Then
                Light33.State = 2
                Light31.State = 0
				LightHits6 = LightHits6 + 1
                Addscore 100000
                CheckWinBattle
            End If
        Case 8:TargetHits8 = TargetHits8 + 1:Addscore 25000:CheckWinBattle
        Case 9
            If Light31.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light31.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
    End Select
    LastSwitchHit = "Target13"
End Sub

Sub Target1_Hit
	DOF 117, DOFPulse
	PlayTargetSound
    PlaySoundAt SoundFXDOF("", 117, DOFPulse, DOFTargets), Target1
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    ' Do some sound or light effect
    Light20.State = 1
    FlashForMs f5, 1000, 50, 0
    ' do some check
    Check2BankTargets
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Mode5Lights(1) = 2 Then
				Mode5Lights(1) = 0
				CalcMode5Lights
                Light29.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light29.State = 2 Then
                Light29.State = 0
				LightHits6 = LightHits6 + 1
                Addscore 100000
                WinBattle
            End If
        Case 8:TargetHits8 = TargetHits8 + 1:Addscore 25000:CheckWinBattle
        Case 9
            If Light29.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light29.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
    End Select
    LastSwitchHit = "Target1"
End Sub

Sub PlayBallLockVideo
    SupressModeMessages 2700
	PuPlayer.playevent pDMDVideo,"BallLock","LOCKLIT.mp4",nPupVideoVolume,60,3,0,""
End Sub

Sub Check2BankTargets
    If light19.state + light20.state = 2 Then
        light19.state = 0
        light20.state = 0
        LightEffect 1
        FlashEffect 1
        Addscore 20000
        If(Light46.State = 0)AND(bMultiballMode = FALSE)Then
            Light46.State = 1
            openDoor
            Playsound "Riff1"
            RotateGuitarBackAndForth 3, 4
            DMD "_", CL("LOCK IS LIT"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
			GeneralPupQueue.Add "PlayBallLockVideo","PlayBallLockVideo",60,0,0,2800,0,False		
'            PuPlayer.playevent pDMDVideo,"BallLock","LOCKLIT.mp4",nPupVideoVolume,65,3,0,""
            UpdateLockLights
            RightFlash
            PlayLockisLit
        ElseIf light53.State = 0 Then 'lit the increase jackpot light if the lock light is lit
            light53.State = 1
        'PlaySound "vo_IncreaseJakpot"
        Else
            Addscore 30000       
    End If
    End If
End Sub

'**************************
' The Lock: Main Multiball
'**************************
' the lock is a virtual lock, where the locked balls are simply counted

Sub Door_Hit
    PlaySoundAt "fx_woodhit", doorf
    OpenDoor
End Sub

Sub lock_Hit  
    Dim delay
    delay = 500
	SoundSaucerLock
    'PlaySoundAt "fx_hole_enter", lock
    bsJackal.AddBall Me
    CloseDoor
    If(bJackpot = True)AND(light45.State = 2)Then
        if Not bSuper then light45.State = 0
        AwardJackpot
    End If
    If light46.State = 1 Then 'lock the ball
        RotateGuitarBackAndForth 3, 4
        LeftFlash
        BallsInLock(CurrentPlayer) = BallsInLock(CurrentPlayer) + 1
        delay = 4000
        Playsound "Riff2"

		if Scorbit.bSessionActive then
			GameModeStrTmp="NA{Blue}:Ball Locked: " &BallsInLock(CurrentPlayer)
			if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
		End If

        Select Case BallsInLock(CurrentPlayer)
            Case 1:DMD "_", CL("BALL 1 LOCKED"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
            SupressModeMessages 2700
            PuPlayer.playevent pDMDVideo,"BallLock","B1.mp4",nPupVideoVolume,60,3,0,""
            UpdateLockLights
            PlayLock1
            Case 2:DMD "_", CL("BALL 2 LOCKED"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
            SupressModeMessages 2700
            PuPlayer.playevent pDMDVideo,"BallLock","B2.mp4",nPupVideoVolume,60,3,0,""
            UpdateLockLights
             PlayLock2
            Case 3:DMD "_", CL("BALL 3 LOCKED"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
            SupressModeMessages 2700
            PuPlayer.playevent pDMDVideo,"BallLock","B3.mp4",nPupVideoVolume,60,3,0,""
            UpdateLockLights
            PlayLock3
        End Select
        light46.State = 0
        If BallsInLock(CurrentPlayer) = 3 Then 'start multiball
            vpmtimer.addtimer 2000, "StartMainMultiball '"
        End If
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Mode5Lights(9) = 2 Then
				Mode5Lights(9) = 0
				CalcMode5Lights
                Light37.State = 0
                Addscore 100000
                CheckWinBattle
                Delay = 1000
            End If
        Case 6
            If Light37.State = 2 Then
                Light34.State = 2
                Light37.State = 0
				LightHits6 = LightHits6 + 1
                Addscore 100000
                CheckWinBattle
                Delay = 1000
            End If
        Case 9
            If Light37.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light37.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
    End Select
    If(Battle(CurrentPlayer, NewBattle) = 2)AND(Battle(CurrentPlayer, 0) = 0)Then 'the battle is ready, so start it
        vpmtimer.addtimer 2000, "StartBattle '"
        delay = 6000
    End If
	JackalHole.timerinterval = Delay
    JackalHole.timerenabled = True
'vpmtimer.addtimer delay, "JackalExit '"

End Sub

Sub UpdateLockLights
    Select Case BallsInLock(CurrentPlayer)
        Case 0:f73.State = 0:f74.State = 0:f75.State = 0
        Case 1:f75.State = 1:f74.State = 2                 'lock 1
        Case 2:f74.State = 1:f73.State = 2                 'lock 2
        Case 3:f73.State = 0:f74.State = 0:f75.State = 0 'lock 3
    End Select
End Sub

Sub StartMainMultiball
	bJackpot = true
	bMainMultiballMode = True
	FullResetBlizzardPrep
    AddMultiball 3
    DMD "_", CL("MULTIBALL"), "_", eNone, eBlinkFast, eNone, 1000, True, ""
    SupressModeMessages 2700
    PuPlayer.playevent pDMDVideo,"Multiball","Multiball.mp4",nPupVideoVolume,65,3,0,""
    Playmultib
    StartJackpots
    ChangeGi 5
    'reset BallsInLock variable
    BallsInLock(CurrentPlayer) = 0
    EnableBallSaver 10
  
End Sub

Sub OpenDoor
	PlaySoundAt "Door-open", Doorf
    Doorf.RotateToEnd
    door.IsDropped = 1
    doorllight.state = 1
End Sub

Sub CloseDoor
	PlaySoundAt "Door-close", Doorf
    Doorf.RotateToStart 
    door.IsDropped = 0
    doorllight.state = 0
End Sub

'**********
' Jackpots
'**********
' Jackpots are enabled during the Main multiball and the wizard battles

Sub StartJackpots

    'turn on the jackpot lights
    Select Case Battle(CurrentPlayer, 0)
'        Case 9 'jackpots on the ramps
'            light44.State = 2
'            light40.State = 2
'        Case 10 'jackpots 
'            light42.State = 2
'            light40.State = 2
'            light45.State = 2
'        Case 11 'jackpots on the ramps
'            light44.State = 2
'            light40.State = 2
'        Case 12 'jackpots 
'            light42.State = 2
'            light40.State = 2
'            light45.State = 2
        Case 13 'final battle - all jackpots on
			bJackpot = true
            light42.State = 2
            light41.State = 2
            light40.State = 2
            light44.State = 2
            light45.State = 2
            Light49.State = 2
            light51.State = 2
        Case Else
            If bMultiballMode Then
				bJackpot = true
                light44.State = 2
                light49.State = 2
            End If
			if bWizMode1Active or bWizMode2Active or bWizMode3Active Then
	            light41.State = 2
                light51.State = 2			
			End If

    End Select
End Sub

Sub ResetJackpotLights 'when multiball is finished, resets jackpot and superjackpot lights
    bJackpot = False
'    Light48.State = 0	' ozzy light
    light42.State = 0
    light41.State = 0
    light40.State = 0
    light44.State = 0
    light45.State = 0
    Light49.State = 0
    light51.State = 0
    light52.State = 0
End Sub

Sub EnableSuperJackpot
    If bJackpot = True Then
        If light42.State + light41.State + light40.State + light44.State + light45.State + Light49.State + light51.State = 0 Then
            'PlaySound "vo_superjackpotislit"
			bSuper = True
            light48.State = 2
            light52.State = 2
            light42.State = 2
            light41.State = 2
            light40.State = 2
            light44.State = 2
            light45.State = 2
            Light49.State = 2
            light51.State = 2		
        End If
    End If
End Sub

'***********************************
' Blue Targets:  The Object Targets
'***********************************

Sub Target2_Hit   
	DOF 120, DOFPulse
	PlayTargetSound
    PlaySoundAtBall SoundFXDOF("", 120, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target2"
    ' Do some sound or light effect
    Light23.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target4_Hit
	DOF 120, DOFPulse
	PlayTargetSound
    PlaySoundAtBall SoundFXDOF("", 120, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target4"
    ' Do some sound or light effect
    Light24.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target5_Hit
	DOF 113, DOFPulse
	PlayTargetSound
    PlaySoundAtBall SoundFXDOF("", 113, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target5"
    ' Do some sound or light effect
    
    Light25.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target7_Hit
	DOF 113, DOFPulse
	PlayTargetSound
    PlaySoundAtBall SoundFXDOF("", 113, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target7"
    ' Do some sound or light effect
    Light26.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target10_Hit
	DOF 114, DOFPulse
	PlayTargetSound
    PlaySoundAtBall SoundFXDOF("", 114, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target10"
    ' Do some sound or light effect
    Light27.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Target8_Hit
	DOF 114, DOFPulse
	PlayTargetSound
    PlaySoundAtBall SoundFXDOF("", 114, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    AddScore 5000
    TargetBonus = TargetBonus + 1
    LastSwitchHit = "Target8"
    ' Do some sound or light effect
    Light28.State = 1
    ' do some check
    Select Case Battle(CurrentPlayer, 0)
        Case 7:TargetHits7 = TargetHits7 + 1:Addscore 10000:CheckWinBattle
    End Select
    Check6BankTargets
End Sub

Sub Check6BankTargets
    Dim tmp
    FlashForMs f1, 1000, 50, 0
    FlashForMs f3, 1000, 50, 0
    FlashForMs f4, 1000, 50, 0
    FlashForMs f4, 1000, 50, 0
    tmp = INT(RND * 1) + 1
    ' if all 6 targets are hit collect Burrito & activate the mystery light
    If light23.state + light24.state + light25.state + light26.state + light27.state + light28.state = 6 Then
       ' enable ball saver on the outlanes
        light13.State = 1
        light16.State = 0
        ' increase the spinner value
        spinnervalue(CurrentPlayer) = spinnervalue(CurrentPlayer) + 1000
        MonstersKilled(CurrentPlayer) = MonstersKilled(CurrentPlayer) + 1

		if Scorbit.bSessionActive then
			GameModeStrTmp="NA{Yellow}:Burrito Found: " &MonstersKilled(CurrentPlayer)
			if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
		End If


        DMD " YOU FOUND ", " A BURRITO ", "object_" &tmp, eNone, eNone, eNone, 2500, True, "itempickup"
		SupressModeMessages 3000
        PuPlayer.playevent pDMDVideo,"Multipliers","Burrito.mp4",nPupVideoVolume,67,3,0,""
        PlayObject
        LightEffect 1
        FlashEffect 1
        ' Lit the Mystery light if it is off
        If Light38.State = 1 Then
            AddScore 50000
        Else
            Light38.State = 1
            AddScore 25000
        End If
        ' reset the lights
        light23.state = 0
        light24.state = 0
        light25.state = 0
        light26.state = 0
        light27.state = 0
        light28.state = 0
    End If
End Sub

' Playfiel Multiplier timer: reduces the multiplier after 30 seconds

Sub pfxtimer_Timer
    If PlayfieldMultiplier(CurrentPlayer) > 1 Then
        PlayfieldMultiplier(CurrentPlayer) = PlayfieldMultiplier(CurrentPlayer)-1
        SetPlayfieldMultiplier PlayfieldMultiplier(CurrentPlayer)
    Else
        pfxtimer.Enabled = 0
    End If
End Sub



'*****************
'  Captive Target
'*****************

Sub Target9_Hit
    ShakeOZZ


    DOF 115, DOFPulse
    PlayTargetSound
    If Tilted Then Exit Sub

    ' ---- SKILLSHOT ----
    If bSkillshotReady Then
        Awardskillshot
        Exit Sub
    End If

    ' ---- TARGET SCORING ----
    AddScore 5000 ' all targets score 5000

    ' ---- SUPER JACKPOT ----
    If (bJackpot = True) AND (light52.State = 2) Then
        AwardSuperJackpot
'		bSuper = True
'        light52.State = 0
'        light48.State = 0
'        StartJackpots
    End If

    ' ---- BATTLE ----
    Select Case Battle(CurrentPlayer, 0)
        Case 0: SelectBattle ' no battle is active then change to another battle
    End Select


    ' ---- JACKPOT INCREASE ----
    If light53.State = 1 Then
		'bJackpot = True
		if bSuper Then
			DMDQueue.Add "AddJackpot 50000","AddJackpot 50000",45,3000,0,0,10000,False		
		Else
			AddJackpot 50000
		End If
'        light53.State = 0
    End If

    ' ---- PLAYFIELD MULTIPLIER ----
    If light54.State = 1 Then
 '       AddPlayfieldMultiplier 1
		if bSuper And bJackpot Then
			DMDQueue.Add "AddPlayfieldMultiplier 1","AddPlayfieldMultiplier 1",45,6000,0,0,10000,False	
		Elseif bSuper OR bJackpot Then	
			DMDQueue.Add "AddPlayfieldMultiplier 1","AddPlayfieldMultiplier 1",45,3000,0,0,10000,False		
		Elseif light53.State = 1 Then
			DMDQueue.Add "AddPlayfieldMultiplier 1","AddPlayfieldMultiplier 1",45,3000,0,0,10000,False
		Else
			AddPlayfieldMultiplier 1	
		End If
'        light54.State = 0
	Else
		'if bSuper = False and bJackpot = False Then 
     PlayTaunt
    End If

	' Moved out of clauses to better handle stacking of calls
	light53.State = 0
    light54.State = 0
End Sub
'****************************
'  Jackal Hole Hit & Awards
'****************************

Sub JackalHole_Hit
    StartTRAINShake
    Playsound "Train_2"
    StartFastSmoke()
    FlashTrainlights 800, 100
    Dim Delay
    Delay = 200
	SoundSaucerLock
    'PlaySoundAt "fx_hole_enter", JackalHole
    bsJackal.AddBall Me
    If NOT Tilted Then
        ' do something
        If(bJackpot = True)AND(light40.State = 2)Then
            if Not bSuper Then light40.State = 0
            AwardJackpot
            Delay = 2000
        End If
        If light39.State = 2 Then ' extra ball is lit
            light39.State = 0
			if light38.state = 1 Then
				GeneralPupQueue.Add "AwardExtraBall","AwardExtraBall",65,5900,0,0,0,False
			Else
				AwardExtraBall
			End If
            Delay = 2000
        End If
        If light38.State = 1 Then ' mystery light is lit
            light38.State = 0
			PlayMadMan
			GeneralPupQueue.Add "GiveRandomAward","GiveRandomAward",65,2900,0,0,0,False
            Delay = 3500
        End If
        Select Case Battle(CurrentPlayer, 0)
            Case 5
				If Mode5Lights(4) = 2 Then
					Mode5Lights(4) = 0
					CalcMode5Lights
                    Light32.State = 0
                    Addscore 100000
                    CheckWinBattle
                    Delay = 1000
                End If
            Case 6
                If Light32.State = 2 Then
                    Light36.State = 2
                    Light32.State = 0
					LightHits6 = LightHits6 + 1
                    Addscore 100000
                    CheckWinBattle
                    Delay = 1000
                End If
            Case 9
                If Light32.State = 2 Then
                    AddScore 100000
                    FlashEffect 3
					LightHits9 = LightHits9 + 1
					CheckWinBattle
                    DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
                End If
            Case 11
                If Light32.State = 2 Then
                    AddScore 120000
                    FlashEffect 3
					LightHits11 = LightHits11 + 1
					CheckWinBattle
                    DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
                End If
        End Select
    End If

	JackalHole.timerinterval = Delay
	JackalHole.timerenabled = True
'    vpmtimer.addtimer Delay, "JackalExit '"
End Sub
Sub JackalHole_timer
	JackalHole.timerenabled = False
	JackalHole.timerinterval = 1000
    If bsJackal.Balls > 0 Then
        FlashForMs f1, 1000, 50, 0
		SoundSaucerKick 1, JackalHole
        PlaySoundAt SoundFXDOF("", 119, DOFPulse, DOFContactors), JackalHole
        DOF 121, DOFPulse
        
        'add a small delay before actually kicking the ball
		JackalHole.timerenabled = True
'       vpmtimer.addtimer 500, "bsJackal.ExitSol_On '"
		Jackalkick.enabled = True
    End If
	'kick out all the balls
'    If bsJackal.Balls > 0 Then
'		JackalHole.timerinterval = 500
'		JackalHole.timerenabled = True
''       vpmtimer.Addtimer 500, "JackalExit '"
'    End If
End Sub

Sub Jackalkick_Timer
	Jackalkick.enabled = False
	bsJackal.ExitSol_On	
End Sub


Sub GiveRandomAward() 
    SupressModeMessages 2700
    Dim tmp, tmp2
    BackleftFlash
    BackRightFlash
    LeftFlash
    RightFlash
    ' show some random values on the dmd
    DMD CL("ITEM AWARD"), "", "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("PLAYFIELD X"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("BUMPER VALUE"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("CRAZY POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA BALL"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("BONUS X"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("SPINNER VALUE"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("BUMPER VALUE"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("BEASTLY POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("PLAYFIELD X"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("BUMPER VALUE"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA POINTS"), "", eNone, eNone, eNone, 50, False, "fx_spinner"
    DMD "_", CL("EXTRA BALL"), "", eNone, eNone, eNone, 50, False, "fx_spinner"

	if light39.state = 2 Then
		tmp = INT(RND(1) * 74) + 6
	Else
		tmp = INT(RND(1) * 80)
	End If

    Select Case tmp
        Case 1, 2, 3, 4, 5, 6 'Lit Extra Ball
            DMD "", CL("EXTRA BALL IS LIT"), "", eNone, eBlink, eNone, 1500, True, "fx_win"
            light39.State = 2
            PlayExtraBallisLit
            PuPlayer.playevent pDMDVideo,"ExtraBall","Extraballlit.mp4",nPupVideoVolume,65,3,0,""
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: EB Lit"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 7, 8, 13, 14, 15 '100,000 points
            DMD CL("BIG POINTS"), CL("6666"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
            AddScore 6666
            PlayBigPoints
            PuPlayer.playevent pDMDVideo,"Mystery","BIGPOINTS.mp4",nPupVideoVolume,65,3,0,""
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: BIG POINTS: 6,666"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 9, 10, 11, 12 'Hold Bonus
            DMD CL("BONUS HELD"), CL("ACTIVATED"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
            bBonusHeld = True
            PlayBonusheld
            PuPlayer.playevent pDMDVideo,"Mystery","bonusheld.mp4",nPupVideoVolume,65,3,0,""
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: BONUS HELD"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 16, 17, 18 'Increase Bonus Multiplier
            DMD CL("INCREASED"), CL("BONUS X"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
            AddBonusMultiplier 1
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: BONUS X INCREASED: " &BonusMultiplier(CurrentPlayer)
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 19, 20, 21 'Complete Battle
            If Battle(CurrentPlayer, 0) > 0 AND Battle(CurrentPlayer, 0) < 13 Then
                DMD CL("BATTLE"), CL("COMPLETED"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
                WinBattle
				if Scorbit.bSessionActive then
					GameModeStrTmp="NA{Blue}:Mystery Award: BATTLE WON"
					if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
				End If
            Else
                DMD CL("BEASTLY POINTS"), CL("66666"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
                AddScore 66666
                PlayBeastlyPoints
                PuPlayer.playevent pDMDVideo,"Mystery","Beastlypoints.mp4",nPupVideoVolume,65,3,0,""
				if Scorbit.bSessionActive then
					GameModeStrTmp="NA{Blue}:Mystery Award: BEASTLY POINTS: 66,666"
					if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
				End If
            End If
        Case 22, 23, 36, 37, 38 'PlayField multiplier
            DMD CL("INCREASED"), CL("PLAYFIELD X"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
            AddPlayfieldMultiplier 1
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: PFX INCREASED: "&PlayfieldMultiplier(CurrentPlayer)
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 24, 25, 26, 27, 28 '100,000 points
            DMD CL("CRAZY POINTS"), CL("666666"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
            AddScore 666666
            PlayCrazyPoints
            PuPlayer.playevent pDMDVideo,"Mystery","CRAZYPOINTS.mp4",nPupVideoVolume,65,3,0,""
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: CRAZY POINTS: 666,666" 
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 29, 30, 31, 32, 33, 34, 35 'Increase Bumper value
            BumperValue(CurrentPlayer) = BumperValue(CurrentPlayer) + 500
            DMD CL("BUMPER VALUE"), CL(BumperValue(CurrentPlayer)), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
            PlayBumperValue
            PuPlayer.playevent pDMDVideo,"Mystery","bumpervalue.mp4",nPupVideoVolume,65,3,0,""
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: BUMPERS INCREASED"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 39, 40, 43, 44 'extra multiball
            DMD CL("EXTRA"), CL("MULTIBALL"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
			FullResetBlizzardPrep
            AddMultiball 1
            EnableBallSaver 10
            PlayInstantMultiball
            PuPlayer.playevent pDMDVideo,"Mystery","InstantMultiball.mp4",nPupVideoVolume,65,3,0,""
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: ADD-A-BALL"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 45, 46, 47, 48 ' Ball Save
            DMD CL("BALL SAVE"), CL("ACTIVATED"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
            EnableBallSaver 20
            PlayBallSaverAct
            PuPlayer.playevent pDMDVideo,"Mystery","20sec.mp4",nPupVideoVolume,65,3,0,""
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: BALL SAVER ENABLED"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 49, 50, 51, 52
            DMD CL("OUTLANE SAVER"), CL("ACTIVATED"), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
            Light13.State = 1
            PlayOutlaneSaverAct
            PuPlayer.playevent pDMDVideo,"Mystery","outlansaver.mp4",nPupVideoVolume,65,3,0,""
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: OUTLANE SAVER ENABLED"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case ELSE 'Add a Random score from 10.000 to 100,000 points
            tmp2 = INT((RND) * 9) * 10000 + 10000
            DMD CL("EXTRA POINTS"), CL(tmp2), "", eBlink, eBlink, eNone, 1500, True, "fx_win"
            AddScore tmp2
            PlayExtraPoints
            PuPlayer.playevent pDMDVideo,"Mystery","EXTRAPOINTS.mp4",nPupVideoVolume,65,3,0,""
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:Mystery Award: EXTRA POINTS: " &tmp2
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
    End Select
End Sub


'*******************
'   The Orbit lanes
'*******************

Sub sw8_Hit
    DOF 130, DOFPulse
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
 '   LaneBonus = LaneBonus + 1
	OrbitHits = OrbitHits + 1
    If(bJackpot = True)AND(light41.State = 2)Then
		if bWizMode2Active  Then
			light41.State = 2
			AwardJackpot
		Elseif bWizMode3Active Then
			light41.State = 2
			AwardSuperJackpot
		Elseif bSuper And bFinalWizModeActive Then
			light41.State = 2
			AwardJackpot
		Else
			light41.State = 0
			AwardJackpot
		End If
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 4:Addscore 70000:CheckWinBattle
        Case 5
            If Mode5Lights(5) = 2 Then
				Mode5Lights(5) = 0
				CalcMode5Lights
                Light33.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light33.State = 2 Then
                Light32.State = 2
                Light33.State = 0
				LightHits6 = LightHits6 + 1
                Addscore 100000
                CheckWinBattle
            End If
        Case 9
            If Light33.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 10
            If LastSwitchHit = "sw7" Then
                LastSwitchHit = ""
                loopCount = loopCount + 1
                Addscore 140000
                CheckWinBattle
            End If
        Case 11
            If Light33.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 12
            If Light33.State = 2 Then
                RampHits12 = RampHits12 + 1
                Light34.State = 2
                Light36.State = 2
                Light33.State = 0
                Light35.State = 0
                Addscore 100000
                CheckWinBattle
            End If
    End Select
    LastSwitchHit = "sw8"

End Sub

Sub sw7_Hit
    DOF 131, DOFPulse
    PlaySoundAtBall "fx_sensor"
    If Tilted Then Exit Sub
'    LaneBonus = LaneBonus + 1
	OrbitHits = OrbitHits + 1:
    If(bJackpot = True)AND(light51.State = 2)Then
		if bWizMode2Active  Then
			light51.State = 2
			AwardJackpot
		Elseif bWizMode3Active  Then
			light51.State = 2
			AwardSuperJackpot
		Elseif bSuper And bFinalWizModeActive Then
			light51.State = 2
			AwardJackpot
		Else
			light51.State = 0
			AwardJackpot
		End If
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 4:Addscore 70000:CheckWinBattle
        Case 5
              If Mode5Lights(7) = 2 Then
				Mode5Lights(7) = 0
				CalcMode5Lights
                Light35.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light35.State = 2 Then
                Light29.State = 2
                Light35.State = 0
				LightHits6 = LightHits6 + 1
                Addscore 100000
                CheckWinBattle
            End If
        Case 9
            If Light35.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 10
            If LastSwitchHit = "sw8" Then
                LastSwitchHit = ""
                loopCount = loopCount + 1
                Addscore 100000
                CheckWinBattle
            End If
        Case 11
            If Light35.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 12
            If Light35.State = 2 Then
                RampHits12 = RampHits12 + 1
                Light34.State = 2
                Light36.State = 2
                Light33.State = 0
                Light35.State = 0
                Addscore 120000
                CheckWinBattle
            End If
    End Select
    LastSwitchHit = "sw7"
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'****************
'     Ramps
'****************
' Ramp Sounds
Sub RightRampStart_Hit
	'ActiveBall.VelY = ActiveBall.VelY * 1.05 ' Ramp Helper for balance
	WireRampOn True	 'Play Plastic Ramp Sound

'debug.print "right Start-" &gametime

	if FTLstep = 5 AND (Battle(CurrentPlayer, 0) = 9 or Battle(CurrentPlayer, 0) = 11) Then
		ResetFTLTimer
	End If
End Sub


Sub RightRampDone_UnHit
	'ActiveBall.VelY = ActiveBall.VelY * 1.05 ' Ramp Helper for balance
	WireRampOn False	'On Wire Ramp, Play Wire Ramp Sound
End Sub

Sub LeftRampStart_Hit
	'ActiveBall.VelY = ActiveBall.VelY * 1.05 ' Ramp Helper for balance
	WireRampOn True	 'Play Plastic Ramp Sound

'	debug.print "LEFT Start-" &gametime

	if FTLstep = 7 AND (Battle(CurrentPlayer, 0) = 9 or Battle(CurrentPlayer, 0) = 11) Then
		ResetFTLTimer
	End If

End Sub


Sub LeftRampDone_UnHit
	'ActiveBall.VelY = ActiveBall.VelY * 1.05 ' Ramp Helper for balance
	WireRampOn False	'On Wire Ramp, Play Wire Ramp Sound
End Sub

Sub LeftRampDone_Hit
    Dim tmp
    DOF 311, DOFPulse
    WireRampOff


'debug.print "LEFT END-" &gametime

    If Tilted Then Exit Sub

    'increase the ramp bonus
    RampBonus = RampBonus + 1
	Addscore SCORE_RAMPS

    ' Normal jackpot logic
    If (bJackpot = True) AND (light44.State = 2) Then
        If bFinalWizModeActive And Not bSuper Then
            light44.State = 0
        Else
            light44.State = 2
        End If

        AwardJackpot
		If bBlizzardMode = False And bMainMultiballMode = False Then
			If Minion1wallUp.Enabled = False And Minion1wallDown.Enabled = False Then
				Minion1wall.z = -110
				Minion1wall.collidable = True
				Minion1wallUp.Enabled = True
			End If
		End If
    End If

    ' 💥 Blizzard mode always awards a jackpot, even outside normal jackpot logic
    If bBLIZZARDMode Then
        AwardJackpot
        ' 💥 Raise Minion1 during Blizzard jackpots if desired:
        'If Minion1wall.transz <= -100 Then Minion1wallUp.Enabled = True
    End If

    'Powerup - ramps the variable and give the jackpots
    If light50.State = 2 Then
        DMD CL("POWER PULSE"), CL(jackpot(CurrentPlayer)), "_", eNone, eBlinkFast, eNone, 1000, True, "" 
        SupressModeMessages 2500
        PuPlayer.playevent pDMDVideo,"Jackpot","Jackpot.mp4",nPupVideoVolume,65,3,0,""    

        AddScore Jackpot(CurrentPlayer)
        LightEffect 2
        FlashEffect 2
        PlayJackpotsound
        If Not JackpotFlashTimer.Enabled Then
            JPCount = 0
            JackpotFlashTimer.Enabled = True
        End If
    Else
        PowerupHits = PowerupHits + 1
        CheckPowerup
    End If

    'Battles
    Select Case Battle(CurrentPlayer, 0)
        Case 3: RampHits3 = RampHits3 + 1: Addscore 300000: CheckWinBattle
        Case 5
            If Mode5Lights(8) = 2 Then
                Mode5Lights(8) = 0
                CalcMode5Lights
                Light36.State = 0
                Addscore 300000
                CheckWinBattle
            End If
        Case 6
            If Light36.State = 2 Then
                Light37.State = 2
                Light36.State = 0
				LightHits6 = LightHits6 + 1
                Addscore 300000
                CheckWinBattle
            End If
        Case 9
            If Light36.State = 2 Then
                AddScore 300000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("300000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light36.State = 2 Then
                AddScore 320000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("320000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 12
            If Light36.State = 2 Then
                RampHits12 = RampHits12 + 1
                Light34.State = 0
                Light36.State = 0
                Light33.State = 2
                Light35.State = 2
                Addscore 300000
                CheckWinBattle
            End If
        Case Else
            ' no special action
    End Select

    ' Combos
    If LastSwitchHit = "RightRampDone" Or LastSwitchHit = "LeftRampDone" Then
        AwardCombo
    End If
    LastSwitchHit = "LeftRampDone"
End Sub

Sub RightRampDone_Hit
    Dim tmp
    DOF 312, DOFPulse
    WireRampOff


'debug.print "right END-" &gametime

    If Tilted Then Exit Sub

    'increase the ramp bonus
    RampBonus = RampBonus + 1
	Addscore SCORE_RAMPS

    ' Normal jackpot logic
    If (bJackpot = True) AND (light49.State = 2) Then
        If bFinalWizModeActive And Not bSuper Then
            light49.State = 0
        Else
            light49.State = 2
        End If

        AwardJackpot
		If bBlizzardMode = False And bMainMultiballMode = False Then
			If Minion2wallUp.Enabled = False And Minion2wallDown.Enabled = False Then
				Minion2wall.z = -110
				Minion2wall.collidable = True
				Minion2wallUp.Enabled = True
			End If
		End If
    End If

    ' 💥 Blizzard mode always awards a jackpot, even outside normal jackpot logic
    If bBLIZZARDMode Then
        AwardJackpot
        ' 💥 Raise Minion2 during Blizzard jackpots if desired:
        'If Minion2wall.transz <= -100 Then Minion2wallUp.Enabled = True
    End If

    'Powerup - ramps the variable and give the jackpots
    If light50.State = 2 Then
        DMD CL("POWER PULSE"), CL(jackpot(CurrentPlayer)), "_", eNone, eBlinkFast, eNone, 1000, True, ""
        SupressModeMessages 2500
        PuPlayer.playevent pDMDVideo,"Jackpot","Jackpot.mp4",nPupVideoVolume,65,3,0,""

        AddScore Jackpot(CurrentPlayer)
        LightEffect 2
        FlashEffect 2
        PlayJackpotsound
        If Not JackpotFlashTimer.Enabled Then
            JPCount = 0
            JackpotFlashTimer.Enabled = True
        End If
    Else
        PowerupHits = PowerupHits + 1
        CheckPowerup
    End If

    'Battles
    Select Case Battle(CurrentPlayer, 0)
        Case 3: RampHits3 = RampHits3 + 1: Addscore 300000: CheckWinBattle
        Case 5
            If Mode5Lights(6) = 2 Then
                Mode5Lights(6) = 0
                CalcMode5Lights
                Light34.State = 0
                Addscore 300000
                CheckWinBattle
            End If
        Case 6
            If Light34.State = 2 Then
                Light35.State = 2
                Light34.State = 0
				LightHits6 = LightHits6 + 1
                Addscore 300000
                CheckWinBattle
            End If
        Case 9
            If Light34.State = 2 Then
                AddScore 300000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("300000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light34.State = 2 Then
                AddScore 320000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("320000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 12
            If Light34.State = 2 Then
                RampHits12 = RampHits12 + 1
                Light34.State = 0
                Light36.State = 0
                Light33.State = 2
                Light35.State = 2
                Addscore 300000
                CheckWinBattle
            End If
        Case Else
            ' no special action
    End Select

    ' Combos
    If LastSwitchHit = "RightRampDone" Or LastSwitchHit = "LeftRampDone" Then
        AwardCombo
    End If
    LastSwitchHit = "RightRampDone"
End Sub


'******************
' Left Target
'******************

Sub Target12_Hit
	DOF 112, DOFPulse
	PlayTargetSound
    If Tilted Then Exit Sub
    If(bJackpot = True)AND(light42.State = 2)Then
        if Not bSuper Then light42.State = 0
        AwardJackpot
    End If
    Select Case Battle(CurrentPlayer, 0)
        Case 5
            If Mode5Lights(2) = 2 Then
				Mode5Lights(2) = 0
				CalcMode5Lights
                Light30.State = 0
                Addscore 100000
                CheckWinBattle
            End If
        Case 6
            If Light30.State = 2 Then
                Light31.State = 2
                Light30.State = 0
				LightHits6 = LightHits6 + 1
                Addscore 100000
                CheckWinBattle
            End If
        Case 9
            If Light30.State = 2 Then
                AddScore 100000
                FlashEffect 3
                LightHits9 = LightHits9 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("100000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
        Case 11
            If Light30.State = 2 Then
                AddScore 120000
                FlashEffect 3
                LightHits11 = LightHits11 + 1
                CheckWinBattle
                DMD "_", CL(FormatScore("120000")), "_", eNone, eBlinkFast, eNone, 500, True, ""
            End If
    End Select
    LastSwitchHit = "Target12"
End Sub

'************************
'       BLIZZARD Events
'************************

' This table has 12 main battles, and a final battle
' you may choose any the 12 main battles you want to play
' After completing all 12 battles you play the final battle

' current active battle number is stored in Battle(CurrentPlayer,0)

Sub SelectBattle 'select a new random battle if none is active
    Dim i

	' Prevent battle from starting during mini-wizard modes
	If bWizMode1Active or bWizMode2Active or bWizMode3Active Then Exit Sub

    If Battle(CurrentPlayer, 0) = 0 Then
        ' reset the battles that are not finished
        For i = 1 to 12
            If Battle(CurrentPlayer, i) = 2 Then Battle(CurrentPlayer, i) = 0
        Next

        If (BattlesWon(CurrentPlayer) Mod 12) = 0 And BattlesWon(CurrentPlayer) > 11 Then
            NewBattle = 13:Battle(CurrentPlayer, NewBattle) = 2:UpdateBattleLights:StartBattle '13 battle is the wizard
        Else
            NewBattle = INT(RND * 12 + 1)
            do while Battle(CurrentPlayer, NewBattle) <> 0
                NewBattle = INT(RND * 11 + 1)
            loop
            Battle(CurrentPlayer, NewBattle) = 2
            Light47.State = 2
            UpdateBattleLights
        End iF
    'debug.print "newbatle " & newbattle
    End If
End Sub

' Update the lights according to the battle's state
Sub UpdateBattleLights
    Light9.State = Battle(CurrentPlayer, 1)
    Light11.State = Battle(CurrentPlayer, 2)
    Light12.State = Battle(CurrentPlayer, 3)
    Light10.State = Battle(CurrentPlayer, 4)
    Light3.State = Battle(CurrentPlayer, 5)
    Light8.State = Battle(CurrentPlayer, 6)
    Light5.State = Battle(CurrentPlayer, 7)
    Light6.State = Battle(CurrentPlayer, 8)
    Light001.State = Battle(CurrentPlayer, 9)
    Light2.State = Battle(CurrentPlayer, 10)
    Light4.State = Battle(CurrentPlayer, 11)
    Light7.State = Battle(CurrentPlayer, 12)
    
End Sub

Sub PlayModeVideo(nMode)	
   PuPlayer.playevent pDMDVideo,"Mode","M"&nMode&".mp4",nPupVideoVolume,60,1,0,""
End Sub

Sub PlayBlizzardCollectVideo
	if bsupressmodemessages Then GeneralPupQueue.Add "PlayBlizzardCollectVideo","PlayBlizzardCollectVideo",60,0,0,0,0,False : Exit Sub
	SupressModeMessages 3100
'	PuPlayer.LabelSet pDMD,"Event3A","",0,""
	PuPlayer.LabelSet pDMD,"Event3B","",0,""
	PuPlayer.LabelSet pDMD,"Event3Ca","",0,""
	PuPlayer.LabelSet pDMD,"Event3C","",0,""
	PuPlayer.playevent pDMDVideo,"BlizMultiball","Blizzard.mp4",nPupVideoVolume,65,3,0,""
	GeneralPupQueue.Add "PlayBlizzCollectBackground","PlayBlizzCollectBackground",60,3300,0,0,0,False
End Sub

Sub PlayBlizzCollectBackground
	PuPlayer.playevent pDMDVideo,"Misc","BlizzardCollect.mp4",nPupVideoVolume,60,1,0,""
	bBLIZZARDMode = True
End Sub


Sub FullResetBlizzardPrep
'	debug.print "CD Timer:" & CountdownTimer.Enabled
'	if CountdownTimer.Enabled Then
		bBlizzardPrepMode = False
		ResetCountdown
		DisableBlizzardTargetLights
		RiseTarget
'	End If
End Sub

' Starting a battle means to setup some lights and variables, maybe timers
' Battle lights will always blink during an active battle
Sub StartBattle
dim i



	'MerlinRTP remove before release
'	pDMDLabelSetColorGradientPercent "TESTQTitle",  cWhite, cBlue, 20
'	pDMDLabelSetColorGradientPercent "TESTQ",  cWhite, cBlue, 60
'
'	PuPlayer.LabelSet pDMD,"TESTQTitle","Active Mode",1,"{'mt':2,'fonth':6,'xalign':0,'yalign':0,'ypos':12,'xpos':"&(0 + nOffsetX) &"}"
'	PuPlayer.LabelSet pDMD,"TESTQ",Battle(CurrentPlayer,0),1,"{'mt':2,'fonth':20,'xalign':0,'yalign':0,'ypos':15,'xpos':"&(3 + nOffsetX) &"}"

	PuPlayer.LabelSet pDMD,"Event3Ca","",1,"{'mt':2,'color': " & cRed &"}"
    StartSpots

	if bBLIZZARDMode Then Exit Sub

	'Make sure it clears timer and pup png
	GeneralPupQueue.Add "HideBlizzTimer","HideBlizzTimer",60,1000,0,0,0,False
	

	FullResetBlizzardPrep

    Battle(CurrentPlayer, 0) = NewBattle
    Light47.State = 0
    'ChangeSong
    EnableBallSaver 15 'start a 15 seconds ball save

	' RESET BORDERS
	PuPlayer.LabelSet pDMD,"Event3A","",0,""
	PuPlayer.LabelSet pDMD,"Event3B","",0,""
	PuPlayer.LabelSet pDMD,"Event3Ca","",0,""
	PuPlayer.LabelSet pDMD,"Event3C","",0,""

	pDMDLabelSetBorder "Event3B",cRed,3,3,1
	pDMDLabelSetBorder "Event3Ca",cYellow,3,3,1
	pDMDLabelSetBorder "Event3C",cRed,3,3,1

	pDMDLabelSetColorGradient "Event3B", cOrange, cRed
	pDMDLabelSetColorGradient "Event3C",  cOrange, cRed


    Select Case NewBattle
        Case 1 'NO REST = Super Spinners
            DMD CL("NO REST FOR THE WICKED"), CL("SHOOT THE SPINNERS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootSpinners
            'PuPlayer.playevent pDMDVideo,"Mode","M1.mp4",nPupVideoVolume,60,1,0,""	
			GeneralPupQueue.Add "PlayModeVideo 1","PlayModeVideo 1",60,0,0,0,0,False
            Light33.State = 2
            Light35.State = 2
            SpinCount = 0
            ChangeGi blue
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - NO REST FOR THE WICKED"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 2 'NO MORE TEARS = Super Pop Bumpers
            DMD CL("NO MORE TEARS"), CL("HIT THE POP BUMPERS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayHitBumpers
			GeneralPupQueue.Add "PlayModeVideo 2","PlayModeVideo 2",60,0,0,0,0,False
            Light55.State = 2
            LightSeqBumpers.Play SeqRandom, 10, , 1000
            SuperBumperHits = 0
            ChangeGi yellow
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - NO MORE TEARS"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 3 'PATIENT NO.9 = Ramps
            DMD CL("PATIENT NO.9"), CL("SHOOT THE RAMPS"), "", eNone, eNone, eNone, 1500, True, ""   
            PlayShootRamps
			GeneralPupQueue.Add "PlayModeVideo 3","PlayModeVideo 3",60,0,0,0,0,False
            Light36.State = 2
            Light34.State = 2
            RampHits3 = 0
            ChangeGi amber   
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - PATIENT NO.9"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 4 'SCREAM = Orbits
            DMD CL("SCREAM"), CL("SHOOT THE ORBITS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootOrbits
			GeneralPupQueue.Add "PlayModeVideo 4","PlayModeVideo 4",60,0,0,0,0,False
            Light33.State = 2
            Light35.State = 2
            ChangeGi Red
            OrbitHits = 0
           	if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - SCREAM"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 5 'DIARY OF A MADMAN = Shoot the lights 2
            DMD CL("DIARY OF A MADMAN"), CL("SHOOT THE LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootLights
			GeneralPupQueue.Add "PlayModeVideo 5","PlayModeVideo 5",60,0,0,0,0,False
            EnableBallSaver 10
			for i = 1 to 9 
				Mode5Lights(i) = 2
			Next
			CalcMode5Lights
            Light29.State = 2
            Light30.State = 2
            Light31.State = 2
            Light32.State = 2
            Light33.State = 2
            Light34.State = 2
            Light35.State = 2
            Light36.State = 2
            Light37.State = 2
            ChangeGi purple
 			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - DIARY OF A MADMAN"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If           
        Case 6 'BLACK RAIN= Shoot the lights 1
            DMD CL("BLACK RAIN"), CL("SHOOT THE LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootRampOrbits
			GeneralPupQueue.Add "PlayModeVideo 6","PlayModeVideo 6",60,0,0,0,0,False
			LightHits6 = 0
            Light30.State = 2
            ChangeGi red
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - BLACK RAIN"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 7 'BARK AT THE MOON =  Blue Target Frenzy
            DMD CL("BARK AT THE MOON"), CL("SHOOT THE TARGETS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootTargets
 			GeneralPupQueue.Add "PlayModeVideo 7","PlayModeVideo 7",60,0,0,0,0,False
            LightSeqBlueTargets.Play SeqRandom, 10, , 1000
            TargetHits7 = 0
            ChangeGi red
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - BARK AT THE MOON"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If            
        Case 8 'UNDERCOVER = Left & Right Targets
            DMD CL("UNDER COVER"), CL("SHOOT THE TARGETS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootTargets
			GeneralPupQueue.Add "PlayModeVideo 8","PlayModeVideo 8",60,0,0,0,0,False
            Light31.State = 2
            Light29.State = 2
            ChangeGi blue
            TargetHits8 = 0
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - UNDER COVER"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 9 'ORDINARY MAN = Follow the Lights 1
            DMD CL("ORDINARY MAN"), CL("CHASE THE LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootLights
			GeneralPupQueue.Add "PlayModeVideo 9","PlayModeVideo 9",60,0,0,0,0,False
            FollowTheLights.Enabled = 1
            LightHits9 = 0
            ChangeGi purple
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - ORDINARY MAN"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If            
        Case 10 'OZZMOSIS = Super Loops
            DMD CL("OZZMOSIS"), CL("SHOOT THE LOOPS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootloops
			GeneralPupQueue.Add "PlayModeVideo 10","PlayModeVideo 10",60,0,0,0,0,False
            Light33.State = 2
            Light35.State = 2
            Gate2.Open = 1
            Gate3.Open = 1
            loopCount = 0
            ChangeGi darkgreen
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - OZZMOSIS"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 11 'DOWN TO EARTH = Follow the Lights 2
            DMD CL("DOWN TO EARTH"), CL("CHASE THE LIGHTS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootLights
			GeneralPupQueue.Add "PlayModeVideo 11","PlayModeVideo 11",60,0,0,0,0,False
            FollowTheLights.Enabled = 1
            LightHits11 = 0
            ChangeGi blue
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - DOWN TO EARTH"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 12 'ULTIMATE SIN = Ramps and Orbits 'uses the ramphits12 to count the hits
            DMD CL("THE ULTIMATE SIN"), CL("SHOOT RAMPS ORBITS"), "", eNone, eNone, eNone, 1500, True, ""
            PlayShootRampOrbits
			GeneralPupQueue.Add "PlayModeVideo 12","PlayModeVideo 12",60,0,0,0,0,False
            Light36.State = 2
            Light34.State = 2
            RampHits12 = 0
            ChangeGi amber
			if Scorbit.bSessionActive then
				GameModeStrTmp="NA{Blue}:MODE - THE ULTIMATE SIN"
				if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
			End If
        Case 13 'PRINCE OF DARKNESS - the final battle
            DMD CL("PRINCE OF DARKNESS"), CL("SHOOT THE JACKPOTS"), "", eNone, eNone, eNone, 1500, True, ""
			StartFinalWizMode        
    End Select

	' Slightly delay the pulsing progress text
	DMDQueue.Add "UpdateModeProgress False","UpdateModeProgress False",45,500,0,0,0,False
	DMDQueue.Add "UpdateModeProgress False","UpdateModeProgress False",45,1000,0,0,0,False
	
End Sub

Sub CheckMode5lights
	dim i,j
	for i = 1 to 9
		j = j + Mode5Lights(i)
	Next

	if j = 9 then WinBattle
End Sub

Sub CalcMode5Lights
	dim i
	Mode5Lights(0) = 0
	for i = 1 to 9
		if Mode5Lights(i) = 0 Then
			Mode5Lights(0) = Mode5Lights(0) + 1
		End If
	Next
End Sub

' check if the battle is completed
Sub CheckWinBattle
    dim tmp
    tmp = INT(RND * 7) + 1
    PlaySound "fx_thunder" & tmp
    LightningStrike()
    DOF 126, DOFPulse
    LightSeqInserts.StopPlay 'stop the light effects before starting again so they don't play too long.
    LightEffect 3
    FlashEffect 3
    Select Case NewBattle
        Case 1
            If SpinCount >= 100 Then WinBattle:End if
        Case 2
            If SuperBumperHits >= 25 Then WinBattle:End if
        Case 3
            If RampHits3 >= 6 Then WinBattle:End if
        Case 4
            If OrbitHits >= 6 Then WinBattle:End if
        Case 5
              If Light29.State + Light30.State + Light31.State + Light32.State + Light33.State + Light34.State + Light35.State + Light36.State + Light37.State = 0 Then WinBattle:End if
			CheckMode5Lights
        Case 6 'the last light win the battle
				' target1 wins battle
        Case 7
            If TargetHits7 >= 20 Then WinBattle:End if
        Case 8
            If TargetHits8 >= 6 Then WinBattle:End if
        Case 9
            If LightHits9 >= 8 Then WinBattle:End if
        Case 10:
            If loopCount >= 6 Then WinBattle
        Case 11
            If LightHits11 >= 8 Then WinBattle:End if
        Case 12
            If RampHits12 >= 6 Then WinBattle:End if
    End Select

	UpdateModeProgress True
End Sub

Sub StopBattle 'called at the end of a ball
'	PuPlayer.LabelSet pDMD,"TESTQTitle","Active Mode",1,"{'mt':2,'fonth':6,'xalign':0,'yalign':0,'ypos':12,'xpos':"&(0 + nOffsetX) &"}"
'	PuPlayer.LabelSet pDMD,"TESTQ","0",1,"{'mt':2,'fonth':20,'xalign':0,'yalign':0,'ypos':15,'xpos':"&(3 + nOffsetX) &"}"


	StopSpinner	
	RiseTarget

	Dim r
	' Make sure Demon GIF is off
	pDMDHideAnimate "DemonGIF"
	r = Battle(CurrentPlayer, 0)
	PuPlayer.playevent pDMDVideo,"Mode","M"&r&".mp4",nPupVideoVolume,60,5,0,""
    StopSpots
    Dim i
    Battle(CurrentPlayer, 0) = 0
    For i = 0 to 15
        If Battle(CurrentPlayer, i) = 2 Then Battle(CurrentPlayer, i) = 0
    Next

	if bBlizzardPrepMode Then
		bBlizzardPrepMode = False
		PuPlayer.playevent pDMDVideo,"Misc","BlizzardCollect.mp4",0,60,5,0,""
	Else
		if Target009.isDropped Then ResetBLIZZARDLights : ResetCountdown
	End If
	
    UpdateBattleLights
    StopBattle2
    NewBattle = 0
End Sub

Sub StopBattle2
    'Turn off the bomb lights
    Light29.State = 0
    Light30.State = 0
    Light31.State = 0
    Light32.State = 0
    Light33.State = 0
    Light34.State = 0
    Light35.State = 0
    Light36.State = 0
    Light37.State = 0
    ' stop some timers or reset battle variables
    Select Case NewBattle
        Case 1:'SpinCount = 0
        Case 2:Light55.State = 0:LightSeqBumpers.StopPlay':SuperBumperHits = 0
        Case 3, 12:
        Case 4:'OrbitHits = 0
        Case 7:LightSeqBlueTargets.StopPlay
        Case 8:
        Case 9, 11:FollowTheLights.Enabled = 0
        Case 10:LoopCount = 0:Gate2.Open = 0:Gate3.Open = 0
        Case 13:ResetBattles:SelectBattle
    End Select
End Sub

Sub ResetBattles
    Dim i, j
    For j = 0 to 4
        BattlesWon(j) = 0
        For i = 0 to 12
            Battle(CurrentPlayer, i) = 0
        Next
    Next
    NewBattle = 0
End Sub

'called after completing a battle
Sub WinBattle

'	PuPlayer.LabelSet pDMD,"TESTQTitle","Active Mode",1,"{'mt':2,'fonth':6,'xalign':0,'yalign':0,'ypos':12,'xpos':"& (0 + nOffsetX) &"}"
'	PuPlayer.LabelSet pDMD,"TESTQ","0",1,"{'mt':2,'fonth':20,'xalign':0,'yalign':0,'ypos':15,'xpos':"& (3 + nOffsetX) &"}"


	StopSpinner	
	RiseTarget

	Dim r
	r = Battle(CurrentPlayer, 0)
	PuPlayer.playevent pDMDVideo,"Mode","M"&r&".mp4",0,60,5,0,""
    StopSpots
    Dim tmp
    BattlesWon(CurrentPlayer) = BattlesWon(CurrentPlayer) + 1
    Battle(CurrentPlayer, 0) = 0
    Battle(CurrentPlayer, NewBattle) = 1
    UpdateBattleLights
    FlashEffect 2
    LightEffect 2
    GiEffect 2
    DMD "", CL("ALBUM COLLECTED"), "_", eNone, eBlinkFast, eNone, 1000, True, "triumph"
    PlayAlbumCollected
    AddScore SCORE_MODESCOMPLETED
    PlayExcelent
    DOF 139, DOFPulse
    StopBattle2
	UpdateAlbumCount
' wiz1 ramps jackpots
'wiz2 ramps & obits jackpots
'wiz3  ramps supers, orbits jackpots
'finall all supers
    'add a multiball after each 2 won battles
    Select Case (BattlesWon(CurrentPlayer) Mod 12)
        Case 3
			if Target009.isDropped Then ResetBLIZZARDLights : ResetCountdown
			GeneralPupQueue.Add "StartWiz1","StartWiz1",60,3100,0,0,0,False		
        Case 6
			if Target009.isDropped Then ResetBLIZZARDLights : ResetCountdown
			GeneralPupQueue.Add "StartWiz2","StartWiz2",60,3100,0,0,0,False
		Case 9
			if Target009.isDropped Then ResetBLIZZARDLights : ResetCountdown
			GeneralPupQueue.Add "StartWiz3","StartWiz3",60,3100,0,0,0,False
		Case Else     
			NewBattle = 0
			SelectBattle 'automatically select a new battle
    End Select


	if Scorbit.bSessionActive then
		GameModeStrTmp="NA{Blue}:MODE - BATTLE WON"
		if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
	End If
End Sub


Sub StartWiz1
	pDMDPNGAnimate "DemonGIF", 30
    StartFlameFollowers
    FlameActive = True
    StartDemFlame
    StartDemonMode
    PlayDEMTaunt
	PlayWiz1Video
	bWizMode1Active = True	
	bJackpot = True
	EnableBallSaver 20
	if bFlamingBalls Then Startfire
	StartMinionMode
	FullResetBlizzardPrep
	AddMultiball 2
	' turn on ramp Lights
	Light44.State = 2
	Light49.State = 2
End Sub

Sub StartWiz2
	pDMDPNGAnimate "DemonGIF", 30
    StartFlameFollowers
     FlameActive = True
    StartDemFlame
    StartDemonMode
    PlayDEMTaunt
	PlayWiz2Video
	bWizMode2Active = True	
	bJackpot = True
	EnableBallSaver 30
	if bFlamingBalls Then Startfire
	StartMinionMode
	FullResetBlizzardPrep
	AddMultiball 3
	' turn on ramp & orbitLights
	Light44.State = 2
	Light49.State = 2
	Light41.State = 2
	Light51.State = 2
End Sub

Sub StartWiz3
	pDMDPNGAnimate "DemonGIF", 30
    StartFlameFollowers
     FlameActive = True
    StartDemFlame
    StartDemonMode
    PlayDEMTaunt
	PlayWiz3Video
	bWizMode3Active = True	
	bJackpot = True
	EnableBallSaver 40
	if bFlamingBalls Then Startfire
	StartMinionMode
	FullResetBlizzardPrep
	AddMultiball 4
	' turn on ramp & orbitLights
	Light44.State = 2
	Light49.State = 2
	Light41.State = 2
	Light51.State = 2
End Sub

Sub StopWizMode1
    PlayWizMend
	'Turn off ramp lights
	Light44.State = 0
	Light49.State = 0
	' Make sure Demon GIF is off
	pDMDHideAnimate "DemonGIF"
    StopFlameFollowers
     FlameActive = False
	PuPlayer.playevent pDMDVideo,"Wizard","MiniWizard1.mp4",nPupVideoVolume,60,5,0,"" : bWizMode1Active = False
	bWizMode1Active = False
	bJackpot = False
	if bFlamingBalls Then StopFire
	NewBattle = 0
	SelectBattle 'automatically select a new battle

	if NOT bExtraBallWonThisBall And Light39.State <> 2 Then
		Light39.State = 2
		AudioQueue.Add "PlayExtraBallisLit","PlayExtraBallisLit",65,100,0,0,0,False	
	End If
End Sub

Sub StopWizMode2
    PlayWizMend
	'Turn off ramp & orbit lights
	Light44.State = 0
	Light49.State = 0
	Light41.State = 0
	Light51.State = 0   
	' Make sure Demon GIF is off
	pDMDHideAnimate "DemonGIF"
    StopFlameFollowers
     FlameActive = False
	PuPlayer.playevent pDMDVideo,"Wizard","MiniWizard2.mp4",nPupVideoVolume,60,5,0,"" : bWizMode1Active = False
	bWizMode2Active = False
	bJackpot = False
	if bFlamingBalls Then StopFire
	NewBattle = 0
	SelectBattle 'automatically select a new battle

	if NOT bExtraBallWonThisBall And Light39.State <> 2 Then
		Light39.State = 2
		AudioQueue.Add "PlayExtraBallisLit","PlayExtraBallisLit",65,100,0,0,0,False	
	End If
End Sub

Sub StopWizMode3
    PlayWizMend
	'Turn off ramp & orbit lights
	Light44.State = 0
	Light49.State = 0
	Light41.State = 0
	Light51.State = 0
	' Make sure Demon GIF is off
	pDMDHideAnimate "DemonGIF"
    StopFlameFollowers
     FlameActive = False
	PuPlayer.playevent pDMDVideo,"Wizard","MiniWizard3.mp4",nPupVideoVolume,60,5,0,"" : bWizMode1Active = False
	bWizMode3Active = False
	bJackpot = False
	if bFlamingBalls Then StopFire
	NewBattle = 0
	SelectBattle 'automatically select a new battle

	if NOT bExtraBallWonThisBall And Light39.State <> 2 Then
		Light39.State = 2
		AudioQueue.Add "PlayExtraBallisLit","PlayExtraBallisLit",65,100,0,0,0,False	
	End If
End Sub

Sub StartFinalWizMode
	pDMDPNGAnimate "DemonGIF", 40
	'Turn off ramp & orbit lights
'	Light44.State = 2
'	Light49.State = 2
'	Light41.State = 2
'	Light51.State = 2
'	Light40.State = 2
    StartFlameFollowers
    FlameActive = True
    StartDemFlame
    StartDemonMode
    PlayDEMTaunt
    StartBats
	StartMinionMode
	GeneralPupQueue.Add "PlayModeVideo 13","PlayModeVideo 13",60,0,0,0,0,False
	FullResetBlizzardPrep
    AddMultiball 5
    startfire
	bFinalWizModeActive = True
	nWizardModeMultiplier = 3
    StartJackpots
    ChangeGi blue
    PlayShootJackpot
    '20 second ball saver
     EnableBallSaver 20

	if Scorbit.bSessionActive then
		GameModeStrTmp="NA{Blue}:MODE - PRINCE OF DARKNESS"
		if nSlowPC = 0 Then Scorbit.SetGameMode(GameModeStrTmp)
	End If    

End Sub

Sub StopFinalWizMode
	' Make sure Demon GIF is off
	pDMDHideAnimate "DemonGIF"
    PlayPrince
	'Turn off ramp & orbit lights
	Light44.State = 0
	Light49.State = 0
	Light41.State = 0
	Light51.State = 0
	Light40.State = 0
	nWizardModeMultiplier = 1
	StopBattle2
    FlameActive = False
    StopFlameFollowers
    StopBats
	PuPlayer.playevent pDMDVideo,"Wizard","FinalWizard.mp4",nPupVideoVolume,65,5,0,"" : bFinalWizModeActive = False
	bFinalWizModeActive = False
	bJackpot = False
	if bFlamingBalls Then StopFire
End Sub

'Extra subs for the battles

Sub LightSeqAllTargets_PlayDone()
    LightSeqAllTargets.Play SeqRandom, 10, , 1000
End Sub

Sub LightSeqBumpers_PlayDone()
    LightSeqBumpers.Play SeqRandom, 10, , 1000
End Sub

Sub LightSeqBlueTargets_PlayDone()
    LightSeqBlueTargets.Play SeqRandom, 10, , 1000
End Sub

' Wizards modes timer
Dim FTLstep:FTLstep = 0

Sub FollowTheLights_Timer
    Light29.State = 0
    Light30.State = 0
    Light31.State = 0
    Light32.State = 0
    Light33.State = 0
    Light34.State = 0
    Light35.State = 0
    Light36.State = 0
    Light37.State = 0
    Select Case Battle(CurrentPlayer, 0)
        Case 9
            Select case FTLstep
                Case 0:FTLstep = 1:Light29.State = 2
                Case 1:FTLstep = 2:Light30.State = 2
                Case 2:FTLstep = 3:Light31.State = 2
                Case 3:FTLstep = 4:Light32.State = 2
                Case 4:FTLstep = 5:Light33.State = 2
                Case 5:FTLstep = 6:Light34.State = 2
                Case 6:FTLstep = 7:Light35.State = 2
                Case 7:FTLstep = 8:Light36.State = 2
                Case 8:FTLstep = 0:Light37.State = 2
            End Select
        Case 11
            if bShotClear Then FTLstep = INT(RND * 9)
            Select case FTLstep
                Case 0:Light29.State = 2
                Case 1:Light30.State = 2
                Case 2:Light31.State = 2
                Case 3:Light32.State = 2
                Case 4:Light33.State = 2
                Case 5:Light34.State = 2
                Case 6:Light35.State = 2
                Case 7:Light36.State = 2
                Case 8:Light37.State = 2
            End Select
			bShotClear = True
    End Select
End Sub

Sub ResetFTLTimer
	FollowTheLights.Enabled = 0
	bShotClear = False
	FollowTheLights.Enabled = 1
End Sub

'**********************
' Power up Jackpot
'**********************
' 30 seconds hurry up with jackpots on the right ramp
' uses variable PowerupHits and the light50

Sub CheckPowerup
    If light50.State = 0 Then
    If light43.State = 0 Then
    If Light011.State = 0 Then
        If PowerupHits MOD 10 = 0 Then
            EnablePowerup
        End If
    End If
    End If
    End If
End Sub

Sub EnablePowerup
    PlayPowerPulse
   DMD "", CL("POWER PULSE HURRY-UP"), "_", eNone, eBlinkFast, eNone, 1000, True, "itempickup"
    SupressModeMessages 2500
    PuPlayer.playevent pDMDVideo,"PowerPulse","PowerPulse.mp4",nPupVideoVolume,67,3,0,""
    ' start the timers
    PowerupTimerExpired.Enabled = True
    PowerupSpeedUpTimer.Enabled = True
    ' turn on the light
    Light50.BlinkInterval = 160
    Light50.State = 2
    Light43.BlinkInterval = 160
    Light43.State = 2
    Light011.BlinkInterval = 160
    Light011.State = 2
End Sub

Sub PowerupTimerExpired_Timer()
    PowerupTimerExpired.Enabled = False
    ' turn off the light
    Light50.State = 0
    Light43.State = 0 
    Light011.State = 0
End Sub

Sub PowerupSpeedUpTimer_Timer()
    PowerupSpeedUpTimer.Enabled = False
    ' Speed up the blinking
    PlayHurryUp
    Light50.BlinkInterval = 80
    Light50.State = 2
    Light43.BlinkInterval = 80
    Light43.State = 2
    Light011.BlinkInterval = 80
    Light011.State = 2
End Sub

' Turntable - Spinner disk

Sub StartSpinner
    ttSpinDisk.MotorOn = True
    TurnTT.Enabled = True
    gi_spin.state = 1
End Sub

Sub StopSpinner
    gi_spin.state = 0
    ttSpinDisk.MotorOn = False
    If ttSpinDisk.Speed < 5 Then
        TurnTT.Enabled = False
    End If
End Sub

Sub TurnTT_Timer
    Dim tmp
    tmp = (SpinnDisk.Rotz + ttSpinDisk.Speed) MOD 360
    SpinnDisk.Rotz = tmp
End Sub

'******************************************************
'	ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity

'*******************************************
' Early 90's and after

Sub InitPolarity()
	Dim x, a
	a = Array(LF, RF)
	For Each x In a
		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
		x.enabled = True
		x.TimeDelay = 60
		x.DebugOn=False ' prints some info in debugger
		
		x.AddPt "Polarity", 0, 0, 0
		x.AddPt "Polarity", 1, 0.05, -5.5
		x.AddPt "Polarity", 2, 0.4, -5.5
		x.AddPt "Polarity", 3, 0.6, -5.0
		x.AddPt "Polarity", 4, 0.65, -4.5
		x.AddPt "Polarity", 5, 0.7, -4.0
		x.AddPt "Polarity", 6, 0.75, -3.5
		x.AddPt "Polarity", 7, 0.8, -3.0
		x.AddPt "Polarity", 8, 0.85, -2.5
		x.AddPt "Polarity", 9, 0.9,-2.0
		x.AddPt "Polarity", 10, 0.95, -1.5
		x.AddPt "Polarity", 11, 1, -1.0
		x.AddPt "Polarity", 12, 1.05, -0.5
		x.AddPt "Polarity", 13, 1.1, 0
		x.AddPt "Polarity", 14, 1.3, 0
		
		x.AddPt "Velocity", 0, 0,	   1
		x.AddPt "Velocity", 1, 0.160, 1.06
		x.AddPt "Velocity", 2, 0.410, 1.05
		x.AddPt "Velocity", 3, 0.530, 1'0.982
		x.AddPt "Velocity", 4, 0.702, 0.968
		x.AddPt "Velocity", 5, 0.95,  0.968
		x.AddPt "Velocity", 6, 1.03,  0.945
	Next
	
	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
	LF.SetObjects "LF", LeftFlipper, TriggerLF
	RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt		'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay		'delay before trigger turns off and polarity is disabled
	Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
	Private Balls(20), balldata(20)
	Private Name
	
	Dim PolarityIn, PolarityOut
	Dim VelocityIn, VelocityOut
	Dim YcoefIn, YcoefOut
	Public Sub Class_Initialize
		ReDim PolarityIn(0)
		ReDim PolarityOut(0)
		ReDim VelocityIn(0)
		ReDim VelocityOut(0)
		ReDim YcoefIn(0)
		ReDim YcoefOut(0)
		Enabled = True
		TimeDelay = 50
		LR = 1
		Dim x
		For x = 0 To UBound(balls)
			balls(x) = Empty
			Set Balldata(x) = new SpoofBall
		Next
	End Sub
	
	Public Sub SetObjects(aName, aFlipper, aTrigger)
		
		If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
		If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
		If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
		If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
		Name = aName
		Set Flipper = aFlipper
		FlipperStart = aFlipper.x
		FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
		FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y
		
		Dim str
		str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
		ExecuteGlobal(str)
		str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
		ExecuteGlobal(str)
		
	End Sub
	
	' Legacy: just no op
	Public Property Let EndPoint(aInput)
		
	End Property
	
	Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
		Select Case aChooseArray
			Case "Polarity"
				ShuffleArrays PolarityIn, PolarityOut, 1
				PolarityIn(aIDX) = aX
				PolarityOut(aIDX) = aY
				ShuffleArrays PolarityIn, PolarityOut, 0
			Case "Velocity"
				ShuffleArrays VelocityIn, VelocityOut, 1
				VelocityIn(aIDX) = aX
				VelocityOut(aIDX) = aY
				ShuffleArrays VelocityIn, VelocityOut, 0
			Case "Ycoef"
				ShuffleArrays YcoefIn, YcoefOut, 1
				YcoefIn(aIDX) = aX
				YcoefOut(aIDX) = aY
				ShuffleArrays YcoefIn, YcoefOut, 0
		End Select
	End Sub
	
	Public Sub AddBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If IsEmpty(balls(x)) Then
				Set balls(x) = aBall
				Exit Sub
			End If
		Next
	End Sub
	
	Private Sub RemoveBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If TypeName(balls(x) ) = "IBall" Then
				If aBall.ID = Balls(x).ID Then
					balls(x) = Empty
					Balldata(x).Reset
				End If
			End If
		Next
	End Sub
	
	Public Sub Fire()
		Flipper.RotateToEnd
		processballs
	End Sub
	
	Public Property Get Pos 'returns % position a ball. For debug stuff.
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x) ) Then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next
	End Property
	
	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x) ) Then
				balldata(x).Data = balls(x)
			End If
		Next
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = abs(PartialFlipCoef-1)
	End Sub
	'Timer shutoff for polaritycorrect
	Private Function FlipperOn()
		If GameTime < FlipAt+TimeDelay Then
			FlipperOn = True
		End If
	End Function
	
	Public Sub PolarityCorrect(aBall)
		If FlipperOn() Then
			Dim tmp, BallPos, x, IDX, Ycoef
			Ycoef = 1
			
			'y safety Exit
			If aBall.VelY > -8 Then 'ball going down
				RemoveBall aBall
				Exit Sub
			End If
			
			'Find balldata. BallPos = % on Flipper
			For x = 0 To UBound(Balls)
				If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)								'find safety coefficient 'ycoef' data
				End If
			Next
			
			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)												'find safety coefficient 'ycoef' data
			End If
			
			'Velocity correction
			If Not IsEmpty(VelocityIn(0) ) Then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
				
				If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
				
				If Enabled Then aBall.Velx = aBall.Velx*VelCoef
				If Enabled Then aBall.Vely = aBall.Vely*VelCoef
			End If
			
			'Polarity Correction (optional now)
			If Not IsEmpty(PolarityIn(0) ) Then
				Dim AddX
				AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
				
				If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
			End If
			If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
		End If
		RemoveBall aBall
	End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
	Dim x, aCount
	aCount = 0
	ReDim a(UBound(aArray) )
	For x = 0 To UBound(aArray)		'Shuffle objects in a temp array
		If Not IsEmpty(aArray(x) ) Then
			If IsObject(aArray(x)) Then
				Set a(aCount) = aArray(x)
			Else
				a(aCount) = aArray(x)
			End If
			aCount = aCount + 1
		End If
	Next
	If offset < 0 Then offset = 0
	ReDim aArray(aCount-1+offset)		'Resize original array
	For x = 0 To aCount-1				'set objects back into original array
		If IsObject(a(x)) Then
			Set aArray(x) = a(x)
		Else
			aArray(x) = a(x)
		End If
	Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
	ShuffleArray aArray1, offset
	ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
	BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)		'Set up line via two points, no clamping. Input X, output Y
	Dim x, y, b, m
	x = input
	m = (Y2 - Y1) / (X2 - X1)
	b = Y2 - m*X2
	Y = M*x+b
	PSlope = Y
End Function

' Used for flipper correction
Class spoofball
	Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
	Public Property Let Data(aBall)
		With aBall
			x = .x
			y = .y
			z = .z
			velx = .velx
			vely = .vely
			velz = .velz
			id = .ID
			mass = .mass
			radius = .radius
		End With
	End Property
	Public Sub Reset()
		x = Empty
		y = Empty
		z = Empty
		velx = Empty
		vely = Empty
		velz = Empty
		id = Empty
		mass = Empty
		radius = Empty
	End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	Dim y 'Y output
	Dim L 'Line
	'find active line
	Dim ii
	For ii = 1 To UBound(xKeyFrame)
		If xInput <= xKeyFrame(ii) Then
			L = ii
			Exit For
		End If
	Next
	If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)		'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )
	
	If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )		 'Clamp lower
	If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )		'Clamp upper
	
	LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
	FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
	FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
	FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
	FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
	Dim b
	Dim gBOT
	gBOT = GetBalls
	
	If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		'   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					Exit Sub
				End If
			Next
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
					gBOT(b).velx = gBOT(b).velx / 1.3
					gBOT(b).vely = gBOT(b).vely - 0.5
				End If
			Next
		End If
	Else
		If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
	End If
End Sub

'*****************
' Maths
'*****************

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
	dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
	dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
	If dx > 0 Then
		Atn2 = Atn(dy / dx)
	ElseIf dx < 0 Then
		If dy = 0 Then
			Atn2 = pi
		Else
			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
		End If
	ElseIf dx = 0 Then
		If dy = 0 Then
			Atn2 = 0
		Else
			Atn2 = Sgn(dy) * pi / 2
		End If
	End If
End Function

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
	Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
	DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
	Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
	AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
	DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
	Dim DiffAngle
	DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
	If DiffAngle > 180 Then DiffAngle = DiffAngle - 360
	
	If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
		FlipperTrigger = True
	Else
		FlipperTrigger = False
	End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
	Case 0
		SOSRampup = 2.5
	Case 1
		SOSRampup = 6
	Case 2
		SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
'	Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
	FlipperPress = 1
	Flipper.Elasticity = FElasticity
	
	Flipper.eostorque = EOST
	Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
	FlipperPress = 0
	Flipper.eostorqueangle = EOSA
	Flipper.eostorque = EOST * EOSReturn / FReturn
	
	If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
		Dim b, gBOT
		gBOT = GetBalls
		
		For b = 0 To UBound(gBOT)
			If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper
	
	If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
		If FState <> 1 Then
			Flipper.rampup = SOSRampup
			Flipper.endangle = FEndAngle - 3 * Dir
			Flipper.Elasticity = FElasticity * SOSEM
			FCount = 0
			FState = 1
		End If
	ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
		If FCount = 0 Then FCount = GameTime
		
		If FState <> 2 Then
			Flipper.eostorqueangle = EOSAnew
			Flipper.eostorque = EOSTnew
			Flipper.rampup = EOSRampup
			Flipper.endangle = FEndAngle
			FState = 2
		End If
	ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
		If FState <> 3 Then
			Flipper.eostorque = EOST
			Flipper.eostorqueangle = EOSA
			Flipper.rampup = Frampup
			Flipper.Elasticity = FElasticity
			FState = 3
		End If
	End If
End Sub

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle)	'-1 for Right Flipper
	Dim LiveCatchBounce																														'If live catch is not perfect, it won't freeze ball totally
	Dim CatchTime
	CatchTime = GameTime - FCount

	If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
		If CatchTime <= LiveCatch * 0.5 Then												'Perfect catch only when catch time happens in the beginning of the window
			LiveCatchBounce = 0
		Else
			LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)		'Partial catch when catch happens a bit late
		End If
		
		If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
		ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
		ball.angmomx = 0
		ball.angmomy = 0
		ball.angmomz = 0
	Else
		If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
	End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************

'******************************************************
' 	ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
	RubbersD.dampen ActiveBall
	TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
	SleevesD.Dampen ActiveBall
	TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD				'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False	'shows info in textbox "TBPout"
RubbersD.Print = False	  'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1		 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967	'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64	   'there's clamping so interpolate up to 56 at least

Dim SleevesD	'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False	'shows info in textbox "TBPout"
SleevesD.Print = False	  'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
	Public Print, debugOn   'tbpOut.text
	Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
	Public ModIn, ModOut
	Private Sub Class_Initialize
		ReDim ModIn(0)
		ReDim Modout(0)
	End Sub
	
	Public Sub AddPoint(aIdx, aX, aY)
		ShuffleArrays ModIn, ModOut, 1
		ModIn(aIDX) = aX
		ModOut(aIDX) = aY
		ShuffleArrays ModIn, ModOut, 0
		If GameTime > 100 Then Report
	End Sub
	
	Public Sub Dampen(aBall)
		If threshold Then
			If BallSpeed(aBall) < threshold Then Exit Sub
		End If
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
		"actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
		If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)
		
		aBall.velx = aBall.velx * coef
		aBall.vely = aBall.vely * coef
		If debugOn Then TBPout.text = str
	End Sub
	
	Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
			aBall.velx = aBall.velx * coef
			aBall.vely = aBall.vely * coef
		End If
	End Sub
	
	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		Dim x
		For x = 0 To UBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
		Next
	End Sub
	
	Public Sub Report() 'debug, reports all coords in tbPL.text
		If Not debugOn Then Exit Sub
		Dim a1, a2
		a1 = ModIn
		a2 = ModOut
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		TBPout.text = str
	End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
	Public ballvel, ballvelx, ballvely
	
	Private Sub Class_Initialize
		ReDim ballvel(0)
		ReDim ballvelx(0)
		ReDim ballvely(0)
	End Sub
	
	Public Sub Update()	'tracks in-ball-velocity
		Dim str, b, AllBalls, highestID
		allBalls = GetBalls
		
		For Each b In allballs
			If b.id >= HighestID Then highestID = b.id
		Next
		
		If UBound(ballvel) < highestID Then ReDim ballvel(highestID)	'set bounds
		If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)	'set bounds
		If UBound(ballvely) < highestID Then ReDim ballvely(highestID)	'set bounds
		
		For Each b In allballs
			ballvel(b.id) = BallSpeed(b)
			ballvelx(b.id) = b.velx
			ballvely(b.id) = b.vely
		Next
	End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
Sub RDampen_Timer
	Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************

'******************************************************
' 	ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1	  '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7	 'Level of bounces. Recommmended value of 0.7

Sub TargetBouncer(aBall,defvalue)
	Dim zMultiplier, vel, vratio
	If TargetBouncerEnabled = 1 And aball.z < 30 Then
		'   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		vel = BallSpeed(aBall)
		If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
		Select Case Int(Rnd * 6) + 1
			Case 1
				zMultiplier = 0.2 * defvalue
			Case 2
				zMultiplier = 0.25 * defvalue
			Case 3
				zMultiplier = 0.3 * defvalue
			Case 4
				zMultiplier = 0.4 * defvalue
			Case 5
				zMultiplier = 0.45 * defvalue
			Case 6
				zMultiplier = 0.5 * defvalue
		End Select
		aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
		aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
		aBall.vely = aBall.velx * vratio
		'   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		'   debug.print "conservation check: " & BallSpeed(aBall)/vel
	End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
	TargetBouncer ActiveBall, 1
End Sub

'******************************************************
' 	ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'	 Metals (all metal objects, metal walls, metal posts, metal wire guides)
'	 Apron (the apron walls and plunger wall)
'	 Walls (all wood or plastic walls)
'	 Rollovers (wire rollover triggers, star triggers, or button triggers)
'	 Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'	 Gates (plate gates)
'	 GatesWire (wire gates)
'	 Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
'	- Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
'	- Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
'	- Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
'	- Ball rolling sounds sub
'
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1				https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2				https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3				https://youtu.be/ESXWGJZY_EI

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1					  'volume level; range [0, 1]
NudgeLeftSoundLevel = 1				 'volume level; range [0, 1]
NudgeRightSoundLevel = 1				'volume level; range [0, 1]
NudgeCenterSoundLevel = 1			   'volume level; range [0, 1]
StartButtonSoundLevel = 0.1			 'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1			   'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010		'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635		'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0					   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45					'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel		'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel	   'sound helper; not configurable
SlingshotSoundLevel = 0.95					  'volume level; range [0, 1]
BumperSoundFactor = 4.25						'volume multiplier; must not be zero
KnockerSoundLevel = 1						   'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2		  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5			 'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5			   'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.375 / 5			'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025		   'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025		   'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8	  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075				   'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5			'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10	 'volume multiplier; must not be zero
DTSoundLevel = 0.25				 'volume multiplier; must not be zero
RolloverSoundLevel = 0.25		   'volume level; range [0, 1]
SpinnerSoundLevel = 0.5			 'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8				   'volume level; range [0, 1]
BallReleaseSoundLevel = 1			   'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2	'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015	 'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5			 'volume multiplier; must not be zero

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
	PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
	PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
	PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
	PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
	PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
	PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
	PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.y * 2 / tableheight - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 10)
	Else
		AudioFade = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.x * 2 / tablewidth - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 10)
	Else
		AudioPan = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
	Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
	Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
	Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
	BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
	VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
	PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
	RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
	RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
	PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
	PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
	PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
	PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
	PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
	PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
	PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
	PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
	FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
	FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
	FlipperLeftHitParm = parm / 10
	If FlipperLeftHitParm > 1 Then
		FlipperLeftHitParm = 1
	End If
	FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
	FlipperRightHitParm = parm / 10
	If FlipperRightHitParm > 1 Then
		FlipperRightHitParm = 1
	End If
	FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
	PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
	PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
	RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 5 Then
		RandomSoundRubberStrong 1
	End If
	If finalspeed <= 5 Then
		RandomSoundRubberWeak()
	End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
	Select Case Int(Rnd * 10) + 1
		Case 1
			PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 2
			PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 3
			PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 4
			PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 5
			PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 6
			PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 7
			PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 8
			PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 9
			PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 10
			PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
	End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
	PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
	RandomSoundWall()
End Sub

Sub RandomSoundWall()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 5) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 5
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 4) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 3) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
	PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
	RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
	RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////

Sub RandomSoundBottomArchBallGuide()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
	PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
	If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
		RandomSoundBottomArchBallGuideHardHit()
	Else
		RandomSoundBottomArchBallGuide
	End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
	If finalspeed < 6 Then
		PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 10 Then
		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft ActiveBall
	Else
		RandomSoundTargetHitWeak()
	End If
End Sub

Sub Targets_Hit (idx)
	PlayTargetSound
	Addscore SCORE_TARGETS
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
	Select Case Int(Rnd * 9) + 1
		Case 1
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 2
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 3
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
		Case 4
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 5
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 6
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 7
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 8
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 9
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
	End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
	PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
	Select Case Int(Rnd * 5) + 1
		Case 1
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 2
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 3
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 4
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 5
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
	End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
	PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
	PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
	SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
	SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
	PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
	PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
	If ActiveBall.velx > 1 Then SoundPlayfieldGate
	StopSound "Arch_L1"
	StopSound "Arch_L2"
	StopSound "Arch_L3"
	StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
	If ActiveBall.velx <  - 8 Then
		RandomSoundRightArch
	End If
End Sub

Sub Arch2_hit()
	If ActiveBall.velx < 1 Then SoundPlayfieldGate
	StopSound "Arch_R1"
	StopSound "Arch_R2"
	StopSound "Arch_R3"
	StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
	If ActiveBall.velx > 10 Then
		RandomSoundLeftArch
	End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
	PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
	Select Case scenario
		Case 0
			PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
		Case 1
			PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
	End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)
	Dim snd
	Select Case Int(Rnd * 7) + 1
		Case 1
			snd = "Ball_Collide_1"
		Case 2
			snd = "Ball_Collide_2"
		Case 3
			snd = "Ball_Collide_3"
		Case 4
			snd = "Ball_Collide_4"
		Case 5
			snd = "Ball_Collide_5"
		Case 6
			snd = "Ball_Collide_6"
		Case 7
			snd = "Ball_Collide_7"
	End Select
	
	PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
	PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
	PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05	  'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
	End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
	End Select
End Sub

'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************

'******************************************************
'	ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
	Dim i
	For i = 0 To tnob
		rolling(i) = False
	Next
End Sub

Sub RollingUpdate()
	Dim b
	Dim gBOT
	gBOT = GetBalls
	
	' stop the sound of deleted balls
	For b = UBound(gBOT) + 1 To tnob - 1
		' Comment the next line if you are not implementing Dyanmic Ball Shadows
		'If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
		rolling(b) = False
		StopSound("BallRoll_" & b)
	Next
	
	' exit the sub if no balls on the table
	If UBound(gBOT) =  lob- 1 Then Exit Sub

'Rotate the idols
      'OZZ002.Rotz = -120 + (gBOT(2).Y)\15
      'OZZ001.Rotz = -120 + (gBOT(2).Y)\15

	
	' play the rolling sound for each ball
	For b = 0 To UBound(gBOT)
		If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
			rolling(b) = True
			PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
		Else
			If rolling(b) = True Then
				StopSound("BallRoll_" & b)
				rolling(b) = False
			End If
		End If
		
		' Ball Drop Sounds
		If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
			If DropCount(b) >= 5 Then
				DropCount(b) = 0
				If gBOT(b).velz >  - 7 Then
					RandomSoundBallBouncePlayfieldSoft gBOT(b)
				Else
					RandomSoundBallBouncePlayfieldHard gBOT(b)
				End If
			End If
		End If
		
		If DropCount(b) < 5 Then
			DropCount(b) = DropCount(b) + 1
		End If
		
		' "Static" Ball Shadows
		' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
		'If AmbientBallShadowOn = 0 Then
		'	If gBOT(b).Z > 30 Then
		'		BallShadowA(b).height = gBOT(b).z - BallSize / 4		'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
		'	Else
		'		BallShadowA(b).height = 0.1
		'	End If
		'	BallShadowA(b).Y = gBOT(b).Y + offsetY
		'	BallShadowA(b).X = gBOT(b).X + offsetX
		'	BallShadowA(b).visible = 1
		'End If
	Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'******************************************************
' 	ZRRL: RAMP ROLLING SFX
'******************************************************

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

Sub WireRampOn(input)
	Waddball ActiveBall, input
	RampRollUpdate
End Sub

Sub WireRampOff()
	WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
	' This will loop through the RampBalls array checking each element of the array x, position 1
	' To see if the the ball was already added to the array.
	' If the ball is found then exit the subroutine
	Dim x
	For x = 1 To UBound(RampBalls)	'Check, don't add balls twice
		If RampBalls(x, 1) = input.id Then
			If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub	'Frustating issue with BallId 0. Empty variable = 0
		End If
	Next
	
	' This will itterate through the RampBalls Array.
	' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
	' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
	' The RampType(BallId) is set to RampInput
	' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
	For x = 1 To UBound(RampBalls)
		If IsEmpty(RampBalls(x, 1)) Then
			Set RampBalls(x, 0) = input
			RampBalls(x, 1) = input.ID
			RampType(x) = RampInput
			RampBalls(x, 2) = 0
			'exit For
			RampBalls(0,0) = True
			RampRoll.Enabled = 1	 'Turn on timer
			'RampRoll.Interval = RampRoll.Interval 'reset timer
			Exit Sub
		End If
		If x = UBound(RampBalls) Then	 'debug
			Debug.print "WireRampOn error, ball queue Is full: " & vbNewLine & _
			RampBalls(0, 0) & vbNewLine & _
			TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
			TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
			TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
			TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
			TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
			" "
		End If
	Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
	'   Debug.Print "In WRemoveBall() + Remove ball from loop array"
	Dim ballcount
	ballcount = 0
	Dim x
	For x = 1 To UBound(RampBalls)
		If ID = RampBalls(x, 1) Then 'remove ball
			Set RampBalls(x, 0) = Nothing
			RampBalls(x, 1) = Empty
			RampType(x) = Empty
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
		'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
		If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
	Next
	If BallCount = 0 Then RampBalls(0,0) = False	'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
	RampRollUpdate
End Sub

Sub RampRollUpdate()	'Timer update
	Dim x
	For x = 1 To UBound(RampBalls)
		If Not IsEmpty(RampBalls(x,1) ) Then
			If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
				If RampType(x) Then
					PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
					StopSound("wireloop" & x)
				Else
					StopSound("RampLoop" & x)
					PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
				End If
				RampBalls(x, 2) = RampBalls(x, 2) + 1
			Else
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
			End If
			If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then	'if ball is on the PF, remove  it
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
				Wremoveball RampBalls(x,1)
			End If
		Else
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
	Next
	If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()	'debug textbox
	Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
	"1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
	"2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
	"3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
	"4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
	"5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
	"6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
	" "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
	BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
	BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************




'***************************************************************************
' VR Plunger Code
'***************************************************************************
Sub TimerVRPlunger2_Timer
	If PinCab_Shooter.Y < 100 then
		PinCab_Shooter.Y = PinCab_Shooter.Y + 5
	End If
End Sub


Sub TimerVRPlunger1_Timer
	PinCab_Shooter.Y = 0 + (5* Plunger.Position) - 20
End Sub

Sub RTP
        li021.state = 1
        li022.state = 1
        li023.state = 1
        li024.state = 1
        li025.state = 1
        li026.state = 1
        li027.state = 1
        li028.state = 1
		CheckBLIZZARDTargets

End Sub

Sub CheckTrustPost
	if RemoveTrustPost Then
		TrustPost.Visible = 0
        TrustSkull.Visible = 0
		TrustPost.Collidable = 0
		TrustPostRubber.visible = 0
	Else
		TrustPost.Visible = 0
        TrustSkull.Visible = 1
		TrustPost.Collidable = 1
		TrustPostRubber.visible = 1
	End If
End Sub

Sub SetTrustPost(Opt)
	Select Case Opt
		Case 0:
			RemoveTrustPost = 0
			CheckTrustPost
		Case 1:
			RemoveTrustPost = 1
			CheckTrustPost
		End Select
End Sub

'==========================
' TRUSTSKULL SHAKE
'==========================

Dim TrustPos, TrustDir

Sub ShakeTrustSkull()
    TrustPos = 10         
    TrustDir = 2
    TrustTimer.Enabled = True
End Sub

Sub TrustTimer_Timer()
    ' Move skull slightly up/down
    TrustSkull.TransY = TrustSkull.TransY + (TrustDir * 2)

    ' Reverse direction each tick
    TrustDir = -TrustDir
    TrustPos = TrustPos - 1

    ' Stop when done
    If TrustPos <= 0 Then
        TrustTimer.Enabled = False
        TrustSkull.TransY = 0   
    End If
End Sub

Sub TrustTrigger_Hit()
    ShakeTrustSkull
End Sub

Sub UpdateLeftOutlanePosts(Opt)
	Select Case Opt
		Case 0
			zCol_Rubber_Post007.y = 1534.33
			Primitive12.y = 1534.33
			RubberOutlaneLeftEasy.visible = True
			RubberOutlaneLeftMed.visible = False
			RubberOutlaneLeftHard.visible = False
		Case 1
			zCol_Rubber_Post007.y = 1536.33
			Primitive12.y = 1536.33
			RubberOutlaneLeftEasy.visible = False
			RubberOutlaneLeftMed.visible = True
			RubberOutlaneLeftHard.visible = False
		Case 2
			zCol_Rubber_Post007.y = 1538.33
			Primitive12.y = 1538.33
			RubberOutlaneLeftEasy.visible = False
			RubberOutlaneLeftMed.visible = False
			RubberOutlaneLeftHard.visible = True
	End Select
End Sub

Sub UpdateRightOutlanePosts(Opt)
	Select Case Opt
		Case 0
			zCol_Rubber_Post035.y = 1534.14
			Primitive1.y = 1534.14
			RubberOutlaneRightEasy.visible = True
			RubberOutlaneRightMed.visible = False
			RubberOutlaneRightHard.visible = False
		Case 1
			zCol_Rubber_Post035.y = 1536.14
			Primitive1.y = 1536.14
			RubberOutlaneRightEasy.visible = False
			RubberOutlaneRightMed.visible = True
			RubberOutlaneRightHard.visible = False
		Case 2
			zCol_Rubber_Post035.y = 1538.14
			Primitive1.y = 1538.14
			RubberOutlaneRightEasy.visible = False
			RubberOutlaneRightMed.visible = False
			RubberOutlaneRightHard.visible = True
	End Select
End Sub

Dim BallsPerGame  ' change in F12 menu
Sub SetBallsPerGame(Opt)
	Select Case Opt
		Case 0: BallsPerGame = 3
		Case 1:	BallsPerGame = 4
		Case 2:	BallsPerGame = 5
	End Select
End Sub

'***ANIMATIONS********

'OZZY

Sub OZZtimer_Timer
	nOzzyDuration = nOzzyDuration - 100
	if nOzzyDuration <= 0 Then OZZtimerstop_Timer : Exit Sub

	countr1 = countr1 + 1 : If Countr1 > 6 then Countr1 = 1 : end If

	select case countr1
		case 1 : OZZ001.z=56:OZZ002.z=-320:OZZ003.z=-320:OZZ004.z=-320:OZZ005.z=-320:OZZ006.z=-320
		case 2 : OZZ001.z=-320:OZZ002.z=56:OZZ003.z=-320:OZZ004.z=-320:OZZ005.z=-320:OZZ006.z=-320
		case 3 : OZZ001.z=-320:OZZ002.z=-320:OZZ003.z=56:OZZ004.z=-320:OZZ005.z=-320:OZZ006.z=-320
		case 4 : OZZ001.z=-320:OZZ002.z=-320:OZZ003.z=-320:OZZ004.z=56:OZZ005.z=-320:OZZ006.z=-320
		case 5 : OZZ001.z=-320:OZZ002.z=-320:OZZ003.z=-320:OZZ004.z=-320:OZZ005.z=56:OZZ006.z=-320
		case 6 : OZZ001.z=-320:OZZ002.z=-320:OZZ003.z=-320:OZZ004.z=-320:OZZ005.z=-320:OZZ006.z=56
	end Select
End Sub

Sub OZZtimerstop_Timer()
	OZZtimer.enabled = False
	OZZTimerstop.enabled = False
	countr1 = 0
End Sub


'DEVIL1

Sub DEVtimer_Timer
'	debug.print "DEV:" &nDevilDuration
'	debug.print "CNT:" &countr2
	nDevilDuration = nDevilDuration - 100
	if nDevilDuration <= 0 Then Devtimerstop_Timer 

	countr2 = countr2 + 1 : If Countr2 > 2 then Countr2 = 1 : end If

	select case countr2
		case 1 : DEV001.z=55:DEV002.z=-350
		case 2 : DEV001.z=-350:DEV002.z=55
	end Select
End Sub

Dim Running
Sub DEVtimerstop_Timer()
'	debug.print "STOP"
	DEVTimerstop.enabled = False
	countr2 = 0
	DEVtimer.enabled = False 

End Sub

'DEVIL2
Sub DEV2timer_Timer
'	nDevilDuration = nDevilDuration - 100
'	if nDevilDuration <= 0 Then Dev2timerstop_Timer : Exit Sub

	countr3 = countr3 + 1 : If Countr3 > 2 then Countr3 = 1 : end If

	select case countr3
		case 1 : DEV001.z=55:DEV002.z=-350
		case 2 : DEV001.z=-350:DEV002.z=55
	end Select


End Sub

Sub DEV2timerstop_Timer()
    if countr3 = 1 then
		DEV2timer.enabled = False
		DEV2Timerstop.enabled = False
	end if
	countr3 = 0
End Sub

'DEVIL3
Sub DEMtimer_Timer
'	nDevilDuration = nDevilDuration - 100
'	if nDevilDuration <= 0 Then Dev2timerstop_Timer : Exit Sub

	countr4 = countr4 + 1 : If Countr4 > 2 then Countr4 = 1 : end If

	select case countr4
		case 1 : DEM002.z=55:DEM003.z=-350
		case 2 : DEM002.z=-350:DEM003.z=55
	end Select


End Sub

Sub DEMtimerstop_Timer()
    if countr4 = 1 then
		DEMtimer.enabled = False
		DEMTimerstop.enabled = False
	end if
	countr4 = 0
End Sub


'************************************
' DEMON MODE DEVIL SWITCH LOGIC
'************************************
' Call StartDemonMode when the mode begins.
' DEV1 hides for 3s while DEV3 shows, then switches back.
'************************************

'--- Start demon mode devil behavior ---
Sub StartDemonMode()
    HideDevil1
    ShowDevil3
    
    DevilSwitchBackTimer.Interval = 2500
    DevilSwitchBackTimer.Enabled = True
End Sub

'--- After 3 seconds, switch back to DEV1 ---
Sub DevilSwitchBackTimer_Timer()
    DevilSwitchBackTimer.Enabled = False
    HideDevil3
    ShowDevil1
End Sub


'************************************
' DEVIL1 ANIMATION HELPERS
'************************************
Sub ShowDevil1()
    DEVTimer.Enabled = True
    DEV001.Z = 55
    DEV002.Z = -350
End Sub

Sub HideDevil1()
    DEVTimer.Enabled = False
    DEV001.Z = -350
    DEV002.Z = -350
End Sub


'************************************
' DEVIL3 ANIMATION HELPERS
'************************************
Sub ShowDevil3()
    DEMTimer.Enabled = True
    DEM002.Visible = True
    DEM003.Visible = True
    DEM002.Z = 55
    DEM003.Z = -350
End Sub

Sub HideDevil3()
    DEMTimer.Enabled = False
    DEM002.Visible = False
    DEM003.Visible = False
    DEM002.Z = -350
    DEM003.Z = -350
End Sub



'***SATAN CALLOUTS****

Sub SatanDone_Hit
    LeftSplat
    RightSplat
    If NOT NewBattle Then PlayDevilTaunt
End Sub
Sub PlayDevilTaunt
	if DEVtimer.enabled then exit sub
    Dim tmp
    tmp = RndNbr(18)

	Select Case tmp
		case 1
			PlaySound "Taunt_1"
			nDevilDuration = 1700
		Case 2
			PlaySound "Taunt_2"
			nDevilDuration = 1400
		Case 3
			PlaySound "Taunt_3"
			nDevilDuration = 1800
		Case 4
			PlaySound "Taunt_4"
			nDevilDuration = 1800
		Case 5
			PlaySound "Taunt_5"
			nDevilDuration = 1800
		case 6
			PlaySound "Taunt_6"
			nDevilDuration = 1900
		Case 7
			PlaySound "Taunt_7"
			nDevilDuration = 1700
		Case 8
			PlaySound "Taunt_19"
			nDevilDuration = 1400
		Case 9
			PlaySound "Taunt_9"
			nDevilDuration = 2200
        Case 10
			PlaySound "Taunt_10"
			nDevilDuration = 1600
        Case 11
			PlaySound "Taunt_11"
			nDevilDuration = 1600
        Case 12
			PlaySound "Taunt_12"
			nDevilDuration = 1600
        Case 13
			PlaySound "Taunt_13"
			nDevilDuration = 1600
        Case 14
			PlaySound "Taunt_14"
			nDevilDuration = 1800
        Case 15
			PlaySound "Taunt_15"
			nDevilDuration = 1800
        Case 16
			PlaySound "Taunt_16"
			nDevilDuration = 1400
        Case 17
			PlaySound "Taunt_17"
			nDevilDuration = 1000
        Case 18
			PlaySound "Taunt_18"
			nDevilDuration = 1200
       
	End Select

	DEVtimer.enabled = True
End Sub

Sub PlayDEMTaunt
	if DEVtimer.enabled then exit sub
    Dim tmp
    tmp = RndNbr(7)
	Select Case tmp
		case 1
			PlaySound "Wizard_1"
			nDevilDuration = 1200
		Case 2
			PlaySound "Wizard_2"
			nDevilDuration = 1200
		Case 3
			PlaySound "Wizard_3"
			nDevilDuration = 1200
		Case 4
			PlaySound "Wizard_4"
			nDevilDuration = 1400
		Case 5
			PlaySound "Wizard_5"
			nDevilDuration = 1400
		case 6
			PlaySound "Wizard_6"
			nDevilDuration = 1400
		Case 7
			PlaySound "Wizard_7"
			nDevilDuration = 2200
	
	End Select

	DEVtimer.enabled = True
End Sub

'*****Ozzy Callouts******

Sub PlayBallLost
	if DEVtimer.enabled then exit sub
    Dim tmp
    tmp = RndNbr(30)

	PlayBallLostVideo
	Select Case tmp
		case 1
			PlaySound "lost_1"
			nDevilDuration = 1000
		Case 2
			PlaySound "lost_2"
			nDevilDuration = 1200
		Case 3
			PlaySound "lost_3"
			nDevilDuration = 1400
		Case 4
			PlaySound "lost_4"
			nDevilDuration = 2000
		Case 5
			PlaySound "lost_5"
			nDevilDuration = 2000
		case 6
			PlaySound "lost_6"
			nDevilDuration = 1200
		Case 7
			PlaySound "lost_7"
			nDevilDuration = 2000
		Case 8
			PlaySound "lost_8"
			nDevilDuration = 2000
		Case 9
			PlaySound "lost_9"
			nDevilDuration = 1400
		Case 10
			PlaySound "lost_10"
			nDevilDuration = 2000
		case 11
			PlaySound "lost_11"
			nDevilDuration = 1800
		Case 12
			PlaySound "lost_12"
			nDevilDuration = 2000
		Case 13
			PlaySound "lost_13"
			nDevilDuration = 1700
		Case 14
			PlaySound "lost_14"
			nDevilDuration = 2000
		Case 15
			PlaySound "lost_15"
			nDevilDuration = 1500
		case 16
			PlaySound "lost_16"
			nDevilDuration = 1200
		Case 17
			PlaySound "lost_17"
			nDevilDuration = 1800
		Case 18
			PlaySound "lost_18"
			nDevilDuration = 1200
		Case 19
			PlaySound "lost_19"
			nDevilDuration = 1200
		Case 20
			PlaySound "lost_20"
			nDevilDuration = 1200
		case 21
			PlaySound "lost_21"
			nDevilDuration = 2000
		Case 22
			PlaySound "lost_22"
			nDevilDuration = 1700
		Case 23
			PlaySound "lost_23"
			nDevilDuration = 2500
		Case 24
			PlaySound "lost_24"
			nDevilDuration = 1700
        Case 25
			PlaySound "lost_25"
			nDevilDuration = 1700
        Case 26
			PlaySound "lost_26"
			nDevilDuration = 1700
        Case 27
			PlaySound "lost_27"
			nDevilDuration = 1700
        Case 28
			PlaySound "lost_28"
			nDevilDuration = 1200
        Case 29
			PlaySound "lost_29"
			nDevilDuration = 1200
        Case 30
			PlaySound "lost_30"
			nDevilDuration = 1200
       
	End Select

	DEVtimer.enabled = True
End Sub

Sub PlayWizMend
	if Devtimer.enabled then exit sub

	PlaySound "taunt_8"
	nDevilDuration = 3500

	Devtimer.enabled = True
End Sub

Sub PlayTaunt
	if OZZtimer.enabled then exit sub

	PlaySound "Ozlaugh_1"
	nOzzyDuration = 1200

	OZZtimer.enabled = True

End Sub

Sub PlayTaunt2
    if OZZtimer.enabled then exit sub

    PlaySound "OZLaugh_2"
    nOzzyDuration = 2600

    OZZtimer.enabled = True

End Sub

Sub PlayPrince
	if OZZtimer.enabled then exit sub
	PlaySound "Prince_1"
	nOzzyDuration = 2400

	OZZtimer.enabled = True

End Sub



Sub PlayBALLSAVE
	if OZZtimer.enabled then exit sub
    Dim tmp
    tmp = RndNbr(3)
	PlayBallSaveVideo
	Select Case tmp
		case 1
			PlaySound "BALLSAVE_1"
			nOzzyDuration = 1200
		Case 2
			PlaySound "BALLSAVE_2"
			nOzzyDuration = 1200
		Case 3
			PlaySound "BALLSAVE_3"
			nOzzyDuration = 1200
	End Select
	OZZtimer.enabled = True   
End Sub

Sub NeedCoins
	if OZZtimer.enabled then exit sub

	PlaySound "vo_needcoins"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayLock1
	if OZZtimer.enabled then exit sub

	PlaySound "vo_ball1locked"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayLock2
	if OZZtimer.enabled then exit sub

	PlaySound "vo_ball2locked"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayLock3
	if OZZtimer.enabled then exit sub

	PlaySound "vo_ball3locked"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayMultib
	if OZZtimer.enabled then exit sub

	PlaySound "vo_multiball"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayCombo
	if OZZtimer.enabled then exit sub

	PlaySound "vo_Combo"
	nOzzyDuration = 600

	OZZtimer.enabled = True
End Sub

Sub Play2XCombo
	if OZZtimer.enabled then exit sub

	PlaySound "vo_2XCombo"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub Play3XCombo
	if OZZtimer.enabled then exit sub

	PlaySound "vo_3XCombo"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub Play4XCombo
	if OZZtimer.enabled then exit sub

	PlaySound "vo_4XCombo"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub Play5XCombo
	if OZZtimer.enabled then exit sub

	PlaySound "vo_5XCombo"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlaySuperCombo
	if OZZtimer.enabled then exit sub

	PlaySound "vo_SuperCombo"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlaySuperDuperCombo
	if OZZtimer.enabled then exit sub

	PlaySound "vo_SuperDuperCombo"
	nOzzyDuration = 1600

	OZZtimer.enabled = True
End Sub

Sub PlayBlizzardMBall
	if OZZtimer.enabled then exit sub

	PlaySound "vo_blizzardmb"
	nOzzyDuration = 1400

	OZZtimer.enabled = True
End Sub

Sub PlayBonusmp
	if OZZtimer.enabled then exit sub

	PlaySound "vo_bonusmp"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayCareful
	if OZZtimer.enabled then exit sub

	PlaySound "vo_careful"
	nOzzyDuration = 800

	OZZtimer.enabled = True
End Sub

Sub PlayExcelent
	if OZZtimer.enabled then exit sub
    Dim tmp
    tmp = RndNbr(5)
	Select Case tmp
		case 1
			PlaySound "vo_excelent"
			nOzzyDuration = 1200
		Case 2
			PlaySound "vo_impressive"
			nOzzyDuration = 1200
		Case 3
			PlaySound "vo_welldone"
			nOzzyDuration = 1200
        Case 4
			PlaySound "vo_YouWon"
			nOzzyDuration = 1200
        Case 5
			PlaySound "Prince_1"
			nOzzyDuration = 1200
	End Select
	OZZtimer.enabled = True   
End Sub

Sub PlayExtraBall
	if OZZtimer.enabled then exit sub
    Dim tmp
    tmp = RndNbr(1)
	Select Case tmp
		case 1
			PlaySound "vo_extraball"
			nOzzyDuration = 1200
	End Select
	OZZtimer.enabled = True
End Sub

Sub PlayExtraBallisLit

	PuPlayer.playevent pDMDVideo,"ExtraBall","ExtraBallisLit.mp4",nPupVideoVolume,65,3,0,""
	PlaySound "vo_extraballislit"

	if OZZtimer.enabled then exit sub

	nOzzyDuration = 1800

	OZZtimer.enabled = True
End Sub

Sub PlayExtragame
	if OZZtimer.enabled then exit sub

	PlaySound "vo_extragame"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayHitBumpers
	if OZZtimer.enabled then exit sub

	PlaySound "vo_hitbumpers"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayJackpotsound
	if OZZtimer.enabled then exit sub

	PlaySound "vo_Jackpot"
	nOzzyDuration = 600

	OZZtimer.enabled = True
End Sub

Sub PlayJackpotIncr
	if OZZtimer.enabled then exit sub

	PlaySound "vo_jackpotinc"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayLockisLit
	if OZZtimer.enabled then exit sub

	PlaySound "vo_lockislit"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayMadMan
	Playsound "Train"
    StopSound "train_2"
    Flashforms Flasher022, 1000, 50, 0
    Flashforms Flasher023, 1000, 50, 0
	'PlaySound "vo_madmanmystery"
	SupressModeMessages 3000
	PuPlayer.playevent pDMDVideo,"Mystery","MadmanMystery.mp4",nPupVideoVolume,65,3,0,""

	AudioQueue.Add "playmadmanaudio","playmadmanaudio",65,250,0,0,0,False	

	if OZZtimer.enabled then exit sub
	AudioQueue.Add "PlayMadmanAnim","PlayMadmanAnim",65,250,0,0,0,False	

End Sub

Sub playmadmanaudio
	PlaySound "vo_madmanmystery" 
End Sub

Sub PlayMadmanAnim
	nOzzyDuration = 1200
	OZZtimer.enabled = True
End Sub

Sub PlayObject
	if OZZtimer.enabled then exit sub

	PlaySound "vo_object"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayPlayersInGame

	PuPlayer.playevent pDMDVideo,"Misc","p"&PlayersPlayingGame&".mp4",nPupVideoVolume,65,3,0,""

	if OZZtimer.enabled then exit sub

	PlaySound "vo_players" &PlayersPlayingGame
	nOzzyDuration = 1000

	OZZtimer.enabled = True
End Sub

Sub PlayPlayerUp
	if OZZtimer.enabled then exit sub

	PlaySound "vo_player_" &CurrentPlayer
	nOzzyDuration = 1000

	OZZtimer.enabled = True
End Sub

Sub PlayPlayfieldInc
	if OZZtimer.enabled then exit sub

	PlaySound "vo_playfield"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayShootAgainsng
	if OZZtimer.enabled then exit sub

	PlaySound "vo_shootagain"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayShootJackpot
	if OZZtimer.enabled then exit sub

	PlaySound "vo_shootjackpots"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayShootLights
	if OZZtimer.enabled then exit sub

	PlaySound "vo_shootlights"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayShootloops
	if OZZtimer.enabled then exit sub

	PlaySound "vo_shootloops"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayShootOrbits
	if OZZtimer.enabled then exit sub

	PlaySound "vo_shootorbits"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayShootRamps
	if OZZtimer.enabled then exit sub

	PlaySound "vo_shootramp"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub


Sub PlayShootRampOrbits
	if OZZtimer.enabled then exit sub

	PlaySound "vo_shootramporb"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayShootSpinners
	if OZZtimer.enabled then exit sub

	PlaySound "vo_shootspinners"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayShootTargets
	if OZZtimer.enabled then exit sub

	PlaySound "vo_shoottargets"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlaySuperJackpotsnd
	if OZZtimer.enabled then exit sub

	PlaySound "vo_superjackpot"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayYouTilted
	if OZZtimer.enabled then exit sub

	PlaySound "vo_youtilted"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlaySkillshot
	if OZZtimer.enabled then exit sub
    Dim tmp
    tmp = RndNbr(1)
	Select Case tmp
		case 1
			PlaySound "fx_skillshot"
            
			nOzzyDuration = 1200
	End Select
	OZZtimer.enabled = True
End Sub

Sub PlayPlayfieldMP
	if OZZtimer.enabled then exit sub
    Dim tmp
    tmp = RndNbr(1)
	Select Case tmp
		case 1
			PlaySound "vo_PlayfieldMP"
			nOzzyDuration = 1600
	End Select
	OZZtimer.enabled = True
End Sub

Sub PlayBlizzardTargetCall
	if OZZtimer.enabled then exit sub
    Dim tmp
    tmp = RndNbr(1)
	Select Case tmp
		case 1
			PlaySound "vo_BlizzardTargets"
			nOzzyDuration = 1800
	End Select
	OZZtimer.enabled = True
End Sub


Sub PlayPowerPulse
	if OZZtimer.enabled then exit sub

	PlaySound "vo_powerpulse"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayHurryUp
	if OZZtimer.enabled then exit sub

	PlaySound "vo_hurryup"
	nOzzyDuration = 600

	OZZtimer.enabled = True
End Sub

Sub PlayCrazyPoints
	if OZZtimer.enabled then exit sub

	PlaySound "vo_crazypointz"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayInstantMultiball
	if OZZtimer.enabled then exit sub

	PlaySound "vo_instantmultiball"
	nOzzyDuration = 1800

	OZZtimer.enabled = True
End Sub

Sub PlayOutlaneSaverAct
	if OZZtimer.enabled then exit sub

	PlaySound "vo_outlaneactivated"
	nOzzyDuration = 2200

	OZZtimer.enabled = True
End Sub

Sub PlayExtraPoints
	if OZZtimer.enabled then exit sub

	PlaySound "vo_extrapoints"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub


Sub PlayBumperValue
	if OZZtimer.enabled then exit sub

	PlaySound "vo_bumpervalue"
	nOzzyDuration = 1800

	OZZtimer.enabled = True
End Sub

Sub PlayBonusheld
	if OZZtimer.enabled then exit sub

	PlaySound "vo_bonusheld"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayBigPoints
	if OZZtimer.enabled then exit sub

	PlaySound "vo_bigpoints"
	nOzzyDuration = 1200

	OZZtimer.enabled = True
End Sub

Sub PlayBeastlyPoints
	if OZZtimer.enabled then exit sub

	PlaySound "vo_beastlypointz"
	nOzzyDuration = 1400

	OZZtimer.enabled = True
End Sub

Sub PlayBallsaveract
	if OZZtimer.enabled then exit sub

	PlaySound "vo_ballsaver"
	nOzzyDuration = 1800

	OZZtimer.enabled = True
End Sub

'****Shake Bats with Bigger Shake and Faster Spin****

Dim CatLPos, CatRPos, DevRPos
Dim CatLSpinStep, CatLSpinCount, CatLSpinning
Dim CatRSpinStep, CatRSpinCount, CatRSpinning

'--- Shake and spin left cat ---
Sub ShakeLeftCat()
    CatLPos = 16              ' bigger shake amplitude (was 8)
    CatLSpinStep = 90         ' faster spin (90 degrees per tick)
    CatLSpinCount = 0
    CatLSpinning = True
    CatLTimer.Enabled = True
End Sub

'--- Timer for left cat ---
Sub CatLTimer_Timer()
    ' Shake movement
    CatL.TransZ = CatLPos
    If CatLPos = 0 Then
        If Not CatLSpinning Then Me.Enabled = False : Exit Sub
    End If

    If CatLPos < 0 Then
        CatLPos = Abs(CatLPos) - 2    ' decrease faster for faster shake
    Else
        CatLPos = -CatLPos + 2
    End If

    ' Spin movement
    If CatLSpinning Then
        CatL.RotZ = (CatL.RotZ + CatLSpinStep) Mod 360
        CatLSpinCount = CatLSpinCount + 1
        If CatLSpinCount >= 8 Then ' shorter spin duration (8 ticks)
            CatLSpinning = False
            CatL.RotZ = 0
            If CatLPos = 0 Then Me.Enabled = False
        End If
    End If
End Sub

'--- Shake and spin right cat ---
Sub ShakeRightCat()
    CatRPos = 16
    CatRSpinStep = 90
    CatRSpinCount = 0
    CatRSpinning = True
    CatRTimer.Enabled = True
End Sub

'--- Timer for right cat ---
Sub CatRTimer_Timer()
    ' Shake movement
    CatR.TransZ = CatRPos
    If CatRPos = 0 Then
        If Not CatRSpinning Then Me.Enabled = False : Exit Sub
    End If

    If CatRPos < 0 Then
        CatRPos = Abs(CatRPos) - 2
    Else
        CatRPos = -CatRPos + 2
    End If

    ' Spin movement
    If CatRSpinning Then
        CatR.RotZ = (CatR.RotZ + CatRSpinStep) Mod 360
        CatRSpinCount = CatRSpinCount + 1
        If CatRSpinCount >= 8 Then
            CatRSpinning = False
            CatR.RotZ = 0
            If CatRPos = 0 Then Me.Enabled = False
        End If
    End If
End Sub

'--- Call these subs on slingshot hit events ---
Sub LeftSlingshot_Hit()
    ShakeLeftCat
End Sub

Sub RightSlingshot_Hit()
    ShakeRightCat
End Sub
'*****Shake Ozzy****

Sub ShakeOZZ
    DEVRPos = 8
    DEVRTimer.Enabled = 1
End Sub

Sub DEVRTimer_Timer
    OZZ001.TransY = DEVRPos
    OZZ002.TransY = DEVRPos
    OZZ003.TransY = DEVRPos
    OZZ004.TransY = DEVRPos
    OZZ005.TransY = DEVRPos
    OZZ006.TransY = DEVRPos
    Halo.TransY = DEVRPos
    If DEVRPos = 0 Then Me.Enabled = 0:Exit Sub
    If DEVRPos < 0 Then
        DEVRPos = ABS(DEVRPos)- 1
    Else
        DEVRPos = - DEVRPos + 1
    End If
End Sub

'****** GUITAR ROTATION ******

Dim GuitarOriginalRotZ
Dim GuitarRotationTarget
Dim GuitarRotationStep
Dim GuitarRotatingBack
Dim GuitarShakesRemaining
Dim GuitarSpeed

Sub RotateGuitarBackAndForth(speed, totalShakes)
    ' provide defaults if arguments omitted
    If IsEmpty(speed) Then speed = 1
    If IsEmpty(totalShakes) Then totalShakes = 1

    GuitarSpeed = CDbl(speed)
    GuitarOriginalRotZ = Guitar.RotZ
    GuitarRotationTarget = GuitarOriginalRotZ - 5
    GuitarRotationStep = -GuitarSpeed
    GuitarRotatingBack = False
    GuitarShakesRemaining = CInt(totalShakes)
    GuitarRotateTimer.Enabled = True
End Sub

Sub GuitarRotateTimer_Timer()
    Guitar.RotZ = Guitar.RotZ + GuitarRotationStep

    If GuitarRotatingBack = False Then
        ' rotating toward target angle
        If Guitar.RotZ <= GuitarRotationTarget Then
            Guitar.RotZ = GuitarRotationTarget
            GuitarRotationStep = GuitarSpeed   ' reverse direction (positive)
            GuitarRotatingBack = True
        End If
    Else
        ' rotating back to original position
        If Guitar.RotZ >= GuitarOriginalRotZ Then
            Guitar.RotZ = GuitarOriginalRotZ
            GuitarShakesRemaining = GuitarShakesRemaining - 1

            If GuitarShakesRemaining > 0 Then
                ' start next shake
                GuitarRotatingBack = False
                GuitarRotationStep = -GuitarSpeed
            Else
                ' finished all shakes
                GuitarRotateTimer.Enabled = False
            End If
        End If
    End If
End Sub

' existing trigger (normal single shake)
Sub rotateguitarccw_Hit()
    RotateGuitarBackAndForth 1, 1
End Sub
'****VAN + VAN2******

' --- Global variables for horizontal movement (X axis) ---
Dim VanOriginalX, VanTargetX, VanStepX, VanMovingBackX
Dim Van2OriginalX, Van2TargetX

' --- Start horizontal move (X axis) ---
Sub MoveVansBackAndForth()
    VanMoveTimer.Enabled = False  ' reset timer if running
    
    ' Van1 setup
    VanOriginalX = Van.x
    VanTargetX = VanOriginalX - 5
    VanStepX = -0.5
    VanMovingBackX = False

    ' Van2 setup (same offset relative to Van2’s current position)
    Van2OriginalX = Van2.x
    Van2TargetX = Van2OriginalX - 5

    VanMoveTimer.Enabled = True
End Sub

' --- Horizontal move timer handler ---
Sub VanMoveTimer_Timer()
    If VanMovingBackX = False Then
        Van.x = Van.x + VanStepX
        Van2.x = Van2.x + VanStepX

        If Van.x <= VanTargetX Then
            Van.x = VanTargetX
            Van2.x = Van2TargetX
            VanMovingBackX = True
            VanStepX = -VanStepX ' reverse direction
        End If
    Else
        Van.x = Van.x + VanStepX
        Van2.x = Van2.x + VanStepX

        If Van.x >= VanOriginalX Then
            Van.x = VanOriginalX
            Van2.x = Van2OriginalX
            VanMoveTimer.Enabled = False
        End If
    End If
End Sub


' --- Global variables for rotation tilt ---
Dim VanOriginalRotX, VanTargetRotX
Dim Van2OriginalRotX, Van2TargetRotX
Dim VanRotStepX, VanTiltingForward, VanTiltingBack, VanTiltPauseCounter

' --- Start van tilt (like a low rider tilt) ---
Sub VansTiltForward()
    If VanTiltingForward Or VanTiltingBack Then Exit Sub ' prevent re-trigger

    ' Van1 setup
    VanOriginalRotX = Van.RotX
    VanTargetRotX = VanOriginalRotX + 10

    ' Van2 setup
    Van2OriginalRotX = Van2.RotX
    Van2TargetRotX = Van2OriginalRotX + 10

    VanRotStepX = 0.5
    VanTiltingForward = True
    VanTiltingBack = False
    VanTiltPauseCounter = 0
    VanTiltTimer.Enabled = True
End Sub

' --- Van tilt timer handler ---
Sub VanTiltTimer_Timer()
    If VanTiltingForward Then
        Van.RotX = Van.RotX + VanRotStepX
        Van2.RotX = Van2.RotX + VanRotStepX

        If Van.RotX >= VanTargetRotX Then
            Van.RotX = VanTargetRotX
            Van2.RotX = Van2TargetRotX
            VanTiltingForward = False
            VanTiltPauseCounter = 30   ' pause duration (~1.5 sec)
        End If

    ElseIf VanTiltPauseCounter > 0 Then
        VanTiltPauseCounter = VanTiltPauseCounter - 1

    ElseIf Not VanTiltingBack Then
        VanTiltingBack = True

    ElseIf VanTiltingBack Then
        Van.RotX = Van.RotX - VanRotStepX
        Van2.RotX = Van2.RotX - VanRotStepX

        If Van.RotX <= VanOriginalRotX Then
            Van.RotX = VanOriginalRotX
            Van2.RotX = Van2OriginalRotX
            VanTiltingBack = False
            VanTiltTimer.Enabled = False
        End If
    End If
End Sub

'**************
' Head Tracking
'**************

Dim HeadPos, OldHeadPos
Dim bHeadMoving
Dim bBreathingFire
Dim StartZ, finalZ
HeadPos = 3:UpdateHead

Sub UpdateHead
    If bBreathingFire or bHeadMoving then Exit Sub
	Select Case HeadPos
        Case 1:
			Select Case OldHeadPos
				Case 1:
					StartZ = -56 : finalZ = -56 : bHeadMoving = False ': HeadTimer.Enabled = 1
				Case 2:
					StartZ = -32 : finalZ = -56 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1			
				Case 3:
					StartZ = -4 : finalZ = -56 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1
				Case 4:
					StartZ = 32 : finalZ = -56 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1
				Case 5:
					StartZ = 56 : finalZ = -56 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1
			End Select
			OldHeadPos = 1
        Case 2:
			Select Case OldHeadPos
				Case 1:
					StartZ = -56 : finalZ = -32 : bHeadMoving =  True:HeadCountUpTimer.Enabled = 1
				Case 2:
					StartZ = -32 : finalZ = -32 : bHeadMoving = False ':HeadCountDownTimer.Enabled = 1			
				Case 3:
					StartZ = -4 : finalZ = -32 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1
				Case 4:
					StartZ = 32 : finalZ = -32 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1
				Case 5:
					StartZ = 56 : finalZ = -32 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1
			End Select
			OldHeadPos = 2
        Case 3:
			Select Case OldHeadPos
				Case 1:
					StartZ = -56 : finalZ = -4 : bHeadMoving = True :HeadCountUpTimer.Enabled = 1
				Case 2:
					StartZ = -32 : finalZ = -4 : bHeadMoving = True :HeadCountUpTimer.Enabled = 1			
				Case 3:
					StartZ = -4 : finalZ = -4 : bHeadMoving = False ':HeadTimer.Enabled = 1
				Case 4:
					StartZ = 32 : finalZ = -4 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1
				Case 5:
					StartZ = 56 : finalZ = -4 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1
			End Select
			OldHeadPos = 3
        Case 4:
			Select Case OldHeadPos
				Case 1:
					StartZ = -56 : finalZ = 32 : bHeadMoving = True :HeadCountUpTimer.Enabled = 1
				Case 2:
					StartZ = -32 : finalZ = 32 : bHeadMoving = True :HeadCountUpTimer.Enabled = 1			
				Case 3:
					StartZ = -4 : finalZ = 32 : bHeadMoving = True :HeadCountUpTimer.Enabled = 1
				Case 4:
					StartZ = 32 : finalZ = 32 : bHeadMoving = False' :HeadTimer.Enabled = 1
				Case 5:
					StartZ = 56 : finalZ = 32 : bHeadMoving = True :HeadCountDownTimer.Enabled = 1
			End Select
			OldHeadPos = 4
        Case 5:
			Select Case OldHeadPos
				Case 1:
					StartZ = -56 : finalZ = 56 : bHeadMoving = True : HeadCountUpTimer.Enabled = 1
				Case 2:
					StartZ = -32 : finalZ = 56 : bHeadMoving = True : HeadCountUpTimer.Enabled = 1			
				Case 3:
					StartZ = -4 : finalZ = 56 : bHeadMoving = True : HeadCountUpTimer.Enabled = 1
				Case 4:
					StartZ = 32 : finalZ = 56 : bHeadMoving = True :HeadCountUpTimer.Enabled = 1
				Case 5:
					StartZ = 56 : finalZ = 56 : bHeadMoving = False ': HeadTimer.Enabled = 1
			End Select
			OldHeadPos = 5
    End Select
End Sub

Sub Head1_Hit:HeadPos = 1:UpdateHead:End Sub
Sub Head2_Hit:HeadPos = 2:UpdateHead:End Sub
Sub Head3_Hit:HeadPos = 3:UpdateHead:End Sub
Sub Head4_Hit:HeadPos = 4:UpdateHead:End Sub
Sub Head5_Hit:HeadPos = 5:UpdateHead:End Sub

Sub HeadCountUpTimer_Timer()
	if StartZ = finalZ Then bHeadMoving = False : HeadCountUpTimer.Enabled = 0 : Exit Sub
	StartZ = StartZ + 4
	OZZ002.rotz = StartZ
	OZZ001.rotz = StartZ
    OZZ003.rotz = StartZ
	OZZ004.rotz = StartZ
    OZZ005.rotz = StartZ
    OZZ006.rotz = StartZ
    Halo.rotz = StartZ
End Sub

Sub HeadCountDownTimer_Timer()
	if StartZ = finalZ Then bHeadMoving = False : HeadCountDownTimer.Enabled = 0 :Exit Sub
	StartZ = StartZ - 4
	OZZ001.rotz = StartZ
	OZZ002.rotz = StartZ
    OZZ003.rotz = StartZ
	OZZ004.rotz = StartZ
    OZZ005.rotz = StartZ
    OZZ006.rotz = StartZ
    Halo.rotz = StartZ
End Sub

' TRAIN Shake

Dim TRAINShake:TRAINShake = 0
Dim TRAINREShake:TRAINREShake=0
Dim TRAINLEShake:TRAINLEShake=0

Sub StartTRAINShake
    TRAINShake = 4:TRAINREShake=4:TRAINLEShake=4
    ShakeTRAIN.Enabled = True
End Sub

Sub ShakeTRAIN_Timer
    TRAIN.Roty = TRAINShake
	TRAIN2.Roty = TRAINShake
	TRAIN3.Roty = TRAINShake
    If TRAINShake = 0 Then ShakeTRAIN.Enabled = False:Exit Sub
    If TRAINShake <0 Then

        TRAINShake = ABS(TRAINShake)- 0.1
		TRAINREShake= ABS(TRAINREShake)- 0.1
		TRAINLEShake= ABS(TRAINLEShake)-1
    Else
        TRAINShake = - TRAINShake + 0.1
		TRAINREShake= -TRAINREShake + 0.1
		TRAINLEShake= -TRAINREShake -0.1
    End If
End Sub

Dim fDuration
Dim fPeriod

Sub FlashTrainlights(fDuration, fPeriod)
	FlashTrainTimer.interval = fPeriod
	FlashTrainTimer.enabled = True
	StopTrainTimer.interval = fDuration
	StopTrainTimer.Enabled = True
End Sub
Sub FlashTrainTimer_Timer
	if Train3.visible = false Then
		Train3.visible = True
	Else
		Train3.visible = False
	End If
   
End Sub

Sub StopTrainTimer_Timer
	StopTrainTimer.enabled = False
	FlashTrainTimer.enabled = False
	Train3.visible = False
End Sub


'***********************
' spotflasher animation
'***********************

Dim MyPii, SpotStep, SpotDir
Dim sRGBStep, sRGBFactor, sRed, sGreen, sBlue

Sub StartSpots
    Spot1.visible = 1
    Spot2.visible = 1
    MyPii = Round(4 * Atn(1), 6) / 90
    SpotStep = 0
    sRGBStep = 0
    sRGBFactor = 5
    sRed = 255
    sGreen = 0
    sBlue = 0
    Spots.Enabled = 1
End Sub

Sub StopSpots
    Spot1.visible = 0
    Spot2.visible = 0
    Spots.Enabled = 0
    Spot1.RotZ = 210
    Camera1.RotZ = 210
    Spot2.RotZ = 170
    Camera2.RotZ = 170
    Spot1.color = RGB(255, 252, 224)
    Spot2.color = RGB(255, 252, 224)
End Sub

Sub Spots_Timer()
    Spot1.visible = 1
    Spot2.visible = 1
    'rotate spots
    SpotDir = SIN(SpotStep * MyPii) * 50
    SpotStep = (SpotStep + 1)MOD 360
    Spot1.RotZ = 210 - SpotDir
    Camera1.RotZ = 210 - SpotDir
    Spot2.RotZ = 170 + SpotDir
    Camera2.RotZ = 170 + SpotDir
    ' color the spotlights
    Select Case sRGBStep
        Case 0 'Green
            sGreen = sGreen + sRGBFactor
            If sGreen > 255 then
                sGreen = 255
                sRGBStep = 1
            End If
        Case 1 'Red
            sRed = sRed - sRGBFactor
            If sRed < 0 then
                sRed = 0
                sRGBStep = 2
            End If
        Case 2 'Blue
            sBlue = sBlue + sRGBFactor
            If sBlue > 255 then
                sBlue = 255
                sRGBStep = 3
            End If
        Case 3 'Green
            sGreen = sGreen - sRGBFactor
            If sGreen < 0 then
                sGreen = 0
                sRGBStep = 4
            End If
        Case 4 'Red
            sRed = sRed + sRGBFactor
            If sRed > 255 then
                sRed = 255
                sRGBStep = 5
            End If
        Case 5 'Blue
            sBlue = sBlue - sRGBFactor
            If sBlue < 0 then
                sBlue = 0
                sRGBStep = 0
            End If
    End Select
    Spot1.color = RGB(sRed, sGreen, sBlue)
    Spot2.color = RGB(sRed, sGreen, sBlue)
End Sub

'**************************
' CRAZY TRAIN DUAL SMOKE SYSTEM (ALTERNATING + FAST MODE)
'**************************

Dim SAVA1Pos, SAVA2Pos, SAVAFLOW
Dim Smoke1FadeDir, Smoke2FadeDir
Dim Smoke1Opacity, Smoke2Opacity
Dim PuffTimerCount
Dim SmokeFastMode, SmokeFastTimer

SAVAFLOW = Array("S_1", "S_2", "S_3", "S_4", "S_5", "S_6", "S_7", "S_8", "S_9", "S_10", _
    "S_11", "S_12", "S_13", "S_14", "S_15", "S_16", "S_17", "S_18", "S_19", "S_20", _
    "S_21", "S_22", "S_23", "S_24", "S_25", "S_26", "S_27", "S_28", "S_29", "S_30", _
    "S_31", "S_32", "S_33", "S_34", "S_35", "S_36", "S_37", "S_38", "S_39", "S_40", _
    "S_41", "S_42", "S_43", "S_44", "S_45", "S_46", "S_47", "S_48", "S_49", "S_50", _
    "S_51", "S_52", "S_53", "S_54", "S_55", "S_56", "S_57", "S_58", "S_59")

'----------------------------
' Start / Stop Smoke
'----------------------------
Sub StartSAVA()
    SAVA1Pos = 0
    SAVA2Pos = 0
    PuffTimerCount = 0
    Smoke1FadeDir = 1
    Smoke2FadeDir = 0
    Smoke1Opacity = 0
    Smoke2Opacity = 0
    SAVA1.Visible = True
    SAVA2.Visible = True
    SmokeFastMode = False
    SAVATimer.Interval = 20
    SAVATimer.Enabled = True
End Sub

Sub StopSAVA()
    SAVATimer.Enabled = False
    SAVA1.Visible = False
    SAVA2.Visible = False
End Sub

'----------------------------
' Smoke Timer
'----------------------------
Sub SAVATimer_Timer()
    UpdateFastSmoke()   ' checks if we’re in fast puff mode
    PuffTimerCount = PuffTimerCount + 1

    ' Alternate puff timing
    Dim puffSpeed
    If SmokeFastMode Then
        puffSpeed = 60   ' faster alternation during fast mode
    Else
        puffSpeed = 120  ' normal rhythm
    End If

    ' alternate puffs between stacks
    If PuffTimerCount Mod (puffSpeed * 2) = 0 Then
        Smoke1FadeDir = 1
        Smoke2FadeDir = 0
    ElseIf PuffTimerCount Mod (puffSpeed * 2) = puffSpeed Then
        Smoke1FadeDir = 0
        Smoke2FadeDir = 1
    End If

    ' animate both
    UpdateSmoke SAVA1, SAVA1Pos, Smoke1FadeDir, Smoke1Opacity
    UpdateSmoke SAVA2, SAVA2Pos, Smoke2FadeDir, Smoke2Opacity
End Sub

'----------------------------
' Smoke Update Sub
'----------------------------
Sub UpdateSmoke(obj, ByRef pos, ByRef dir, ByRef op)
    obj.ImageA = SAVAFLOW(pos)
    pos = (pos + 1) Mod 59

    'fade up/down
    If dir = 1 Then
        op = op + 0.04
        If op >= 1 Then dir = -1
    ElseIf dir = -1 Then
        op = op - 0.04
        If op <= 0 Then dir = 0
    End If

    obj.IntensityScale = op
End Sub

'----------------------------
' Fast Puff Mode Control
'----------------------------
Sub StartFastSmoke()
    SmokeFastMode = True
    SAVATimer.Interval = 10    ' faster frame rate
    SmokeFastTimer = 150       ' ~3 seconds
End Sub

Sub UpdateFastSmoke()
    If SmokeFastMode Then
        SmokeFastTimer = SmokeFastTimer - 1
        If SmokeFastTimer <= 0 Then
            SmokeFastMode = False
            SAVATimer.Interval = 20   ' back to normal
        End If
    End If
End Sub



' VLM  Arrays - Start
' Arrays per baked part

' Arrays per lighting scenario
Dim BL_All_Lights_gi_001: BL_All_Lights_gi_001=Array(LM_All_Lights_gi_001_Playfield)
Dim BL_All_Lights_gi_002: BL_All_Lights_gi_002=Array(LM_All_Lights_gi_002_Playfield)
Dim BL_All_Lights_gi_003: BL_All_Lights_gi_003=Array(LM_All_Lights_gi_003_Playfield)
Dim BL_All_Lights_gi_004: BL_All_Lights_gi_004=Array(LM_All_Lights_gi_004_Playfield)
Dim BL_All_Lights_gi_005: BL_All_Lights_gi_005=Array(LM_All_Lights_gi_005_Playfield)
Dim BL_All_Lights_gi_006: BL_All_Lights_gi_006=Array(LM_All_Lights_gi_006_Playfield)
Dim BL_All_Lights_gi_007: BL_All_Lights_gi_007=Array(LM_All_Lights_gi_007_Playfield)
Dim BL_All_Lights_gi_008: BL_All_Lights_gi_008=Array(LM_All_Lights_gi_008_Playfield)
Dim BL_All_Lights_gi_009: BL_All_Lights_gi_009=Array(LM_All_Lights_gi_009_Playfield)
Dim BL_All_Lights_gi_010: BL_All_Lights_gi_010=Array(LM_All_Lights_gi_010_Playfield)
Dim BL_All_Lights_gi_011: BL_All_Lights_gi_011=Array(LM_All_Lights_gi_011_Playfield)
Dim BL_All_Lights_gi_012: BL_All_Lights_gi_012=Array(LM_All_Lights_gi_012_Playfield)
Dim BL_All_Lights_gi_013: BL_All_Lights_gi_013=Array(LM_All_Lights_gi_013_Playfield)
Dim BL_All_Lights_gi_014: BL_All_Lights_gi_014=Array(LM_All_Lights_gi_014_Playfield)
Dim BL_All_Lights_gi_015: BL_All_Lights_gi_015=Array(LM_All_Lights_gi_015_Playfield)
Dim BL_All_Lights_gi_016: BL_All_Lights_gi_016=Array(LM_All_Lights_gi_016_Playfield)
Dim BL_All_Lights_gi_017: BL_All_Lights_gi_017=Array(LM_All_Lights_gi_017_Playfield)
Dim BL_All_Lights_gi_018: BL_All_Lights_gi_018=Array(LM_All_Lights_gi_018_Playfield)
Dim BL_All_Lights_gi_019: BL_All_Lights_gi_019=Array(LM_All_Lights_gi_019_Playfield)
Dim BL_All_Lights_gi_020: BL_All_Lights_gi_020=Array(LM_All_Lights_gi_020_Playfield)
Dim BL_All_Lights_gi_021: BL_All_Lights_gi_021=Array(LM_All_Lights_gi_021_Playfield)
Dim BL_All_Lights_gi_022: BL_All_Lights_gi_022=Array(LM_All_Lights_gi_022_Playfield)
Dim BL_All_Lights_gi_023: BL_All_Lights_gi_023=Array(LM_All_Lights_gi_023_Playfield)
' Global arrays
Dim BG_Lightmap: BG_Lightmap=Array(LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_014_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield, LM_All_Lights_gi_019_Playfield, LM_All_Lights_gi_020_Playfield, LM_All_Lights_gi_021_Playfield, LM_All_Lights_gi_022_Playfield, LM_All_Lights_gi_023_Playfield)
Dim BG_All: BG_All=Array(LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_014_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield, LM_All_Lights_gi_019_Playfield, LM_All_Lights_gi_020_Playfield, LM_All_Lights_gi_021_Playfield, LM_All_Lights_gi_022_Playfield, LM_All_Lights_gi_023_Playfield)
' VLM  Arrays - End

'**********************
' Multi Primitive Bobble (10 heads, velocity based, all in sync)
'**********************
Dim BobbleX, BobbleY
Dim BobbleVX, BobbleVY

Sub InitBobble()
    BobbleX = 0 : BobbleY = 0
    BobbleVX = 0 : BobbleVY = 0
    BobbleTimer.Interval = 20   ' ~50 fps
    BobbleTimer.Enabled = True
End Sub

' Triggered when ball hits the invisible trigger near the heads
Sub SatanBob_Hit()
    Dim impactX, impactY
    ' Scale wobble by ball velocity
    impactX = ActiveBall.VelX / 8
    impactY = ActiveBall.VelY / 8
    
    ' Add wobble energy
    BobbleVX = BobbleVX + impactX + (Rnd - 0.5) * 2
    BobbleVY = BobbleVY + impactY + (Rnd - 0.5) * 2
End Sub

Sub BobbleTimer_Timer()
    ' X axis
    BobbleX = BobbleX + BobbleVX
    BobbleVX = BobbleVX - (BobbleX * 0.12)   ' spring pullback
    BobbleVX = BobbleVX * 0.92               ' damping

    ' Y axis
    BobbleY = BobbleY + BobbleVY
    BobbleVY = BobbleVY - (BobbleY * 0.12)
    BobbleVY = BobbleVY * 0.92

    ' Apply same wobble to all Dev primitives
    Dev001.RotX = BobbleY : Dev001.RotY = BobbleX
    Dev002.RotX = BobbleY : Dev002.RotY = BobbleX
    DeM003.RotX = BobbleY : DeM003.RotY = BobbleX
    DeM002.RotX = BobbleY : DeM002.RotY = BobbleX
End Sub

'***Speakers****

Sub speakertimer_Timer
countr = countr + 1 : If Countr > 2 then Countr = 1 : end If
select case countr
case 1 : Lr1.z=250:Lr2.z=-250
case 2 : Lr1.z=-250:Lr2.z=250

end Select
End Sub 

Sub speaker2timer_Timer
countra = countra + 1 : If Countra > 2 then Countra = 1 : end If
select case countra
case 1 :Sr1.z=250:Sr2.z=-250
case 2 : Sr1.z=-250:Sr2.z=250

end Select
End Sub 



'**************************
' LAVA lamps
'**************************

Dim LAVA1POS,LAVA2POS, LAVAFLOW
LAVAFLOW = Array("f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", _
    "f10", "f11", "f12", "f13", "f14", "f15", "f16")

Sub StartLAVA
    LAVA2POS = 2
    LAVA1POS = 0
    LAVATimer.Enabled = 1
End Sub

Sub LAVATimer_Timer
    'debug.print fire1pos
    LAVA2.ImageA = LAVAFLOW(LAVA2Pos)
    LAVA2Pos = (LAVA2Pos + 1) MOD 16
    LAVA1.ImageA = LAVAFLOW(LAVA1Pos)
    LAVA1Pos = (LAVA1Pos + 1) MOD 16
End Sub

Dim RSplat, LSplat, R2Splat

Sub RightSplat
    RSplat = 0
    Rightblood_Timer
End Sub

Sub Rightblood_Timer
    Select Case RSplat
        Case 0:Rightblood.ImageA = "blood1":Rightblood.Visible = 1:Rightblood.TimerEnabled = 1
        Case 1:Rightblood.ImageA = "blood2"
        Case 2:Rightblood.ImageA = "blood3"
        Case 3:Rightblood.ImageA = "blood4"
        Case 4:Rightblood.ImageA = "blood5"
        Case 5:Rightblood.ImageA = "blood6"
        Case 6:Rightblood.Visible = 0:Rightblood.TimerEnabled = 0
    End Select
    RSplat = RSplat + 1
End Sub

Sub LeftSplat
    LSplat = 0
    Leftblood_Timer
End Sub

Sub Leftblood_Timer
    Select Case LSplat
        Case 0:Leftblood.ImageA = "blood1a":Leftblood.Visible = 1:Leftblood.TimerEnabled = 1
        Case 1:Leftblood.ImageA = "blood2a"
        Case 2:Leftblood.ImageA = "blood3a"
        Case 3:Leftblood.ImageA = "blood4a"
        Case 4:Leftblood.ImageA = "blood5a"
        Case 5:Leftblood.ImageA = "blood6a"
        Case 6:Leftblood.Visible = 0:Leftblood.TimerEnabled = 0
    End Select
    LSplat = LSplat + 1
End Sub


Sub loadBG
	PuPlayer.LabelSet pDMD, "DMDOverlay", "PupOverlays\\defaultDMD.png",1,"{'mt':2,'zback':1}"

	if DMDType <> 2 And PlatformOS = "windows" Then 
		SetRuleCards 0
	Elseif PlatformOS <> "windows" Then
		PuPlayer.playevent pBackglass,"Backglass","Backglass.mp4",0,20,6,0,""
	End If


	PuPlayer.playevent pDMDVideo,"Background","Blank2.mp4",0,20,6,0,""

End Sub


sub LoadAttract
	PuPlayer.LabelSet pDMD, "DMDOverlay", "PupOverlays\\defaultDMD.png",1,"{'mt':2,'zback':1}"
End Sub


Sub ResetOverlay
	PuPlayer.LabelSet pDMD, "DMDOverlay", "PupOverlays\\defaultDMD.png",1,"{'mt':2,'zback':1}"
End Sub

' *********************************************************************
'              Supporting Player and Score Routines
' *********************************************************************

Sub DMDUpdateBallNumber(nBallNr)



	if renderingmode = 2 Then
		PuPlayer.LabelSet pDMD,"BallValue",nBallNr,1,"{'mt':2,'fonth':13,'xalign':0,'yalign':0,'ypos':63,'xpos':90.75}"
	elseif PlatformOS <> "windows" Then
		PuPlayer.LabelSet pDMD,"BallValue",nBallNr,1,"{'mt':2,'fonth':"&nDMDFontSize&",'xalign':0,'yalign':0,'ypos':63,'xpos':"&(91.25+nOffsetX)&"}"
	elseif DMDType > 0 then
		PuPlayer.LabelSet pDMD,"BallValue",nBallNr,1,"{'mt':2,'fonth':13,'xalign':0,'yalign':0,'ypos':63,'xpos':"&(91.25+nOffsetX)&"}"
	Else
		PuPlayer.LabelSet pDMD,"BallValue",nBallNr,1,"{'mt':2,'fonth':13,'xalign':0,'yalign':0,'ypos':63,'xpos':"&(91.25-nDTOffsetX)&"}"
	End If
	pDMDLabelSetColorGradient "BallValue",  cOrange, cRed
End Sub

Sub UpdateAlbumCount



	if renderingmode = 2 then
		PuPlayer.LabelSet pDMD,"AlbumNum",(BattlesWon(CurrentPlayer) Mod 12),1,"{'mt':2,'fonth':13,'xalign':0,'yalign':0,'ypos':65,'xpos':3.6}"
	elseif PlatformOS <> "windows" Then
		PuPlayer.LabelSet pDMD,"AlbumNum",(BattlesWon(CurrentPlayer) Mod 12),1,"{'mt':2,'fonth':"&nDMDFontSize&",'xalign':0,'yalign':0,'ypos':65,'xpos':"&(4.2+nOffsetX)&"}"
	elseif DMDType > 0 then
		PuPlayer.LabelSet pDMD,"AlbumNum",(BattlesWon(CurrentPlayer) Mod 12),1,"{'mt':2,'fonth':13,'xalign':0,'yalign':0,'ypos':65,'xpos':"&(4.2+nOffsetX)&"}"
	Else
		PuPlayer.LabelSet pDMD,"AlbumNum",(BattlesWon(CurrentPlayer) Mod 12),1,"{'mt':2,'fonth':13,'xalign':0,'yalign':0,'ypos':65,'xpos':"&(4.2-nDTOffsetX)&"}"
	End If
	pDMDLabelSetColorGradient "AlbumNum",  cYellow, cRed
End Sub

Sub UpdateComboCount


	if renderingmode = 2 then
		PuPlayer.LabelSet pDMD,"ComboNum",ComboCount,1,"{'mt':2,'fonth':13,'xalign':0,'yalign':0,'ypos':43,'xpos':90.5}"
	elseif PlatformOS <> "windows" Then	
		PuPlayer.LabelSet pDMD,"ComboNum",ComboCount,1,"{'mt':2,'fonth':"&nDMDFontSize&",'xalign':0,'yalign':0,'ypos':43,'xpos':"&(91+nOffsetX)&"}"	
	elseif DMDType > 0 then
		PuPlayer.LabelSet pDMD,"ComboNum",ComboCount,1,"{'mt':2,'fonth':13,'xalign':0,'yalign':0,'ypos':43,'xpos':"&(91+nOffsetX)&"}"	
	Else
		PuPlayer.LabelSet pDMD,"ComboNum",ComboCount,1,"{'mt':2,'fonth':13,'xalign':0,'yalign':0,'ypos':43,'xpos':"&(91-nDTOffsetX)&"}"	
	End If
	pDMDLabelSetColorGradient "ComboNum",  cOrange, cRed
End Sub

Sub DMDClearPlayerName

	pDMDLabelHide "CurrName"
	PuPlayer.LabelSet pDMD,"Position1Score","",1,""
	PuPlayer.LabelSet pDMD,"Position2Score","",1,""
	PuPlayer.LabelSet pDMD,"Position3Score","",1,""
	PuPlayer.LabelSet pDMD,"Position4Score","",1,""
End Sub

Sub DMDUpdatePlayerName

	DMDClearPlayerName



'	if VRRoom > 0 Then
'		PuPlayer.LabelSet pDMD,"CurrScore"," " & FormatScore(Score(nplayer)) & " ",1,"{'mt':2,'fonth':"& VR_Score &",'xpos':50,'ypos':87.0}"
'	Elseif Score(nPlayer) > 999999999 Then 
'		PuPlayer.LabelSet pDMD,"CurrScore"," " & FormatScore(Score(nplayer)) & " ",1,"{'mt':2,'fonth':8,'xpos':50,'ypos':87.0}"
'	Elseif nPlayersinGame < 3 Then
'		PuPlayer.LabelSet pDMD,"CurrScore"," " & FormatScore(Score(nplayer)) & " ",1,"{'mt':2,'fonth':10,'xpos':50,'ypos':88.5}"
'	Else
'pDMDLabelSetColorGradient "CurrScore",  cWhite, cPurple


	if PlayersPlayingGame = 1 Then
		PuPlayer.LabelSet pDMD,"CurrScore"," " & FormatScoreDMD(Score(CurrentPlayer)) & " ",1,"{'mt':2,'fonth':17,'xpos':50,'ypos':87.5}"
	Else
		PuPlayer.LabelSet pDMD,"CurrScore"," " & FormatScoreDMD(Score(CurrentPlayer)) & " ",1,"{'mt':2,'fonth':13,'xpos':50,'ypos':87.5}"
	End If
'	End If


	if renderingmode = 2  then
		PuPlayer.LabelSet pDMD,"CurrName",CurrentPlayer,1,"{'mt':2,'fonth':13 ,'xalign':0,'yalign':0,'xpos':90.75,'ypos':82.5}"
	elseif PlatformOS <> "windows" Then
		PuPlayer.LabelSet pDMD,"CurrName",CurrentPlayer,1,"{'mt':2,'fonth':"&nDMDFontSize&", 'xalign':0,'yalign':0,'xpos':"&(91.25+nOffsetX)&",'ypos':82.5}"
	elseif DMDType > 0 then
		PuPlayer.LabelSet pDMD,"CurrName",CurrentPlayer,1,"{'mt':2,'fonth':13, 'xalign':0,'yalign':0,'xpos':"&(91.25+nOffsetX)&",'ypos':82.5}"
	else
		PuPlayer.LabelSet pDMD,"CurrName",CurrentPlayer,1,"{'mt':2,'fonth':13, 'xalign':0,'yalign':0,'xpos':"&(91.25-nDTOffsetX)&",'ypos':82.5}"
	end if

	pDMDLabelSetColorGradient "CurrName",  cPink, cRed

	Select Case CurrentPlayer
		Case 1
			
			if PlayersPlayingGame > 1 Then
'				PuPlayer.LabelSet pDMD, "Bullet2", "PuPOverlays\\P2.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':76.5,'xpos':26}"
				PuPlayer.LabelSet pDMD,"Position2Score",FormatScoreDMD(Score(2)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':0,'yalign':1,'xpos':29.25,'ypos':80.0}"  
			End If

			if PlayersPlayingGame > 2 Then
'				PuPlayer.LabelSet pDMD, "Bullet3", "PuPOverlays\\P3.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':90.5,'xpos':26}"
				PuPlayer.LabelSet pDMD,"Position3Score",FormatScoreDMD(Score(3)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':0,'yalign':1,'xpos':29.25,'ypos':94.5}"  
			End If

			if PlayersPlayingGame > 3 Then
'				PuPlayer.LabelSet pDMD, "Bullet4", "PuPOverlays\\P4.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':90.5,'xpos':70.25}"
				PuPlayer.LabelSet pDMD,"Position4Score",FormatScoreDMD(Score(4)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':2,'yalign':1,'xpos':70.75,'ypos':94.5}" 
			End If
		Case 2 

'			PuPlayer.LabelSet pDMD, "Bullet1", "PuPOverlays\\P1.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':76.5,'xpos':26}"
			PuPlayer.LabelSet pDMD,"Position1Score",FormatScoreDMD(Score(1)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':0,'yalign':1,'xpos':29.25,'ypos':80.0}"  

			if PlayersPlayingGame > 2 Then
'				PuPlayer.LabelSet pDMD, "Bullet3", "PuPOverlays\\P3.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':90.5,'xpos':26}"
				PuPlayer.LabelSet pDMD,"Position3Score",FormatScoreDMD(Score(3)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':0,'yalign':1,'xpos':29.25,'ypos':94.5}"  
			End If

			if PlayersPlayingGame > 3 Then
'				PuPlayer.LabelSet pDMD, "Bullet4", "PuPOverlays\\P4.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':90.5,'xpos':70.25}"
				PuPlayer.LabelSet pDMD,"Position4Score",FormatScoreDMD(Score(4)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':2,'yalign':1,'xpos':70.75,'ypos':94.5}"
			End If

		Case 3
'			PuPlayer.LabelSet pDMD, "Bullet1", "PuPOverlays\\P1.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':76.5,'xpos':26}"
			PuPlayer.LabelSet pDMD,"Position1Score",FormatScoreDMD(Score(1)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':0,'yalign':1,'xpos':29.25,'ypos':80.0}"  

'			PuPlayer.LabelSet pDMD, "Bullet2", "PuPOverlays\\P2.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':90.5,'xpos':26}"
			PuPlayer.LabelSet pDMD,"Position2Score",FormatScoreDMD(Score(2)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':0,'yalign':1,'xpos':29.25,'ypos':94.5}" 

			if PlayersPlayingGame > 3 Then
'				PuPlayer.LabelSet pDMD, "Bullet4", "PuPOverlays\\P4.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':90.5,'xpos':70.25}"
				PuPlayer.LabelSet pDMD,"Position4Score",FormatScoreDMD(Score(4)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':2,'yalign':1,'xpos':70.75,'ypos':94.5}"
			End If
		Case 4
'			PuPlayer.LabelSet pDMD, "Bullet1", "PuPOverlays\\P1.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':76.5,'xpos':26}"
			PuPlayer.LabelSet pDMD,"Position1Score",FormatScoreDMD(Score(1)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':0,'yalign':1,'xpos':29.25,'ypos':80.0}"  

'			PuPlayer.LabelSet pDMD, "Bullet2", "PuPOverlays\\P2.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':90.5,'xpos':26}"
			PuPlayer.LabelSet pDMD,"Position2Score",FormatScoreDMD(Score(2)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':0,'yalign':1,'xpos':29.25,'ypos':94.5}"  

'			PuPlayer.LabelSet pDMD, "Bullet3", "PuPOverlays\\P3.png",1,"{'mt':2,'width':4, 'height':6,'xalign':0,'yalign':0,'ypos':90.5,'xpos':70.25}"
			PuPlayer.LabelSet pDMD,"Position3Score",FormatScoreDMD(Score(3)),1,"{'mt':2,'color': "&cWhite&" ,'xalign':2,'yalign':1,'xpos':70.75,'ypos':94.5}" 
	End Select

	pDMDLabelSetColorGradientPercent "CurrScore",  cWhite, cPurple, 60
End Sub

Sub DMDUpdateAll
	DMDUpdatePlayerName
	DMDUpdateBallNumber Balls
	UpdateAlbumCount
	UpdateComboCount
'	DisplayBonusValue
'	DisplayPlayfieldValue
End Sub

' *********************************************************************
'              End PupDMD Supporting Player and Score Routines
' *********************************************************************

'********************* START OF PUPDMD FRAMEWORK v3.0 BETA *************************
'******************************************************************************
'*****   Create a PUPPack within PUPPackEditor for layout config!!!  **********
'******************************************************************************
'
'
'  Quick Steps:
'      1>  create a folder in PUPVideos with Starter_PuPPack.zip and call the folder "yourgame"
'      2>  above set global variable pGameName="yourgame"
'      3>  copy paste the settings section above to top of table script for user changes.
'      4>  on Table you need to create ONE timer only called pupDMDUpdate and set it to 250 ms enabled on startup.
'      5>  go to your table1_init or table first startup function and call PUPINIT function
'      6>  Go to bottom on framework here and setup game to call the appropriate events like pStartGame (call that in your game code where needed)...etc
'      7>  attractmodenext at bottom is setup for you already,  just go to each case and add/remove as many as you want and setup the messages to show.  
'      8>  Have fun and use pDMDDisplay(xxxx)  sub all over where needed.  remember its best to make a bunch of mp4 with text animations... looks the best for sure!
'
'
'Note:  for *Future Pinball* "pupDMDupdate_Timer()" timer needs to be renamed to "pupDMDupdate_expired()"  and then all is good.
'       and for future pinball you need to add the follow lines near top
'Need to use BAM and have com idll enabled.
'				Dim icom : Set icom = xBAM.Get("icom") ' "icom" is name of "icom.dll" in BAM\Plugins dir
'				if icom is Nothing then MSGBOX "Error cannot run without icom.dll plugin"
'				Function CreateObject(className)       
'   					Set CreateObject = icom.CreateObject(className)   
'				End Function


'**************************
'   PinUp Player USER Config
'**************************

Dim pGameName       : pGameName="BlizzardOfOzz"  'pupvideos foldername, probably set to cGameName in realworld


Const HasPuP = True   'dont set to false as it will break pup

'Screens
	Const pTopper	=0
	Dim pDMD
	Const pBackglass=2
	Dim pDMDFull
	Dim pDMDVideo 
	Const pPlayfield=3
	Const pMusic	=4
	Const pAudio	=7
	Const pCallouts	=8
	Const pOverVid	=11
	Const pTransp   =15
	Const pHS		=16
	Const pHS2		=17




Dim PuPlayer
dim PUPDMDObject  'for realtime mirroring.
Dim pDMDlastchk: pDMDLastchk= -1    'performance of updates
Dim pDMDCurPage: pDMDCurPage= 0     'default page is empty.
Dim pInAttract : pInAttract=false   'pAttract mode
Dim pFrameSizeX: pFrameSizeX=1920     'DO NOT CHANGE, this is pupdmd author framesize
Dim pFrameSizeY: pFrameSizeY=1080     'DO NOT CHANGE, this is pupdmd author framesize
Dim pUseFramePos : pUseFramePos=1     'DO NOT CHANGE, this is pupdmd author setting




'*************  starts PUP system,  must be called AFTER b2s/controller running so put in last line of table1_init
Sub PuPInit

Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")   
PuPlayer.B2SInit "", pGameName

PuPType = 0

	IIDMDTimer.Interval = 810

	pupPackScreenFile = PuPlayer.GetRoot & "BlizzardOfOzz\ScreenType.txt"
	Set ObjFso = CreateObject("Scripting.FileSystemObject")
	Set ObjFile = ObjFso.OpenTextFile(pupPackScreenFile)
	DMDType = ObjFile.ReadLine

	

	pDMDVideo = 11

	if PlatformOS = "windows" Then

		Select Case DMDType
			Case 0
				pDMD = 5
				pDMDFull = 5
	'			pDMDVideo = 5
				nOffsetX = 1.25
			Case 2
				pDMD = 2
				pDMDFull = 2
				nOffsetX = 1.25
			Case 3
				pDMD = 5
				pDMDFull = 5
				nOffsetX = 0
			Case 4
				pDMD = 5
				pDMDFull = 5
				nOffsetX = 0
			Case 5
				pDMD = 5
				pDMDFull = 5
				nOffsetX = 0
		End Select
	Else
		Select Case DMDType
			Case 0
				pDMD = 5
				pDMDFull = 5
	'			pDMDVideo = 5
				nOffsetX = 1.25
			Case 2
				pDMD = 2
				pDMDFull = 2
				nOffsetX = 1.25
			Case 3
				pDMD = 5
				pDMDFull = 5
				nOffsetX = 1.25
			Case 4
				pDMD = 5
				pDMDFull = 5
				nOffsetX = 1.25
			Case 5
				pDMD = 5
				pDMDFull = 5
				nOffsetX = 0
		End Select

	End If

	

	CheckPupVersion
	PuPlayer.LabelInit pDMD
	PuPlayer.LabelInit pTopper
	'PuPlayer.LabelInit pDMDVideo
	PuPlayer.LabelInit pBackglass

	ResetOverlay

pSetPageLayouts

	if DMDType = 99 Then 
		BatReminder.Visible = True
	Else
		BatReminder.Visible = False
	End If

pDMDSetPage(pDMDBlank)   'set blank text overlay page.

	if ScorbitActive Then
		pBackglassSetPage 1
		dbg "Delaying Pup startup"
		vpmtimer.addtimer 2000, "pDMDStartup '"
		'BallHandlingQueue.Add "pDMDStartUP","pDMDStartUP",24,2000,0,0,0,False
	Else
		pDMDStartUP
	End If

	if Scorbitactive then 
		if Scorbit.DoInit(4408, "PupOverlays", Version, "blizzardofoz-vpin") then 	' Prod
			tmrScorbit.Interval=2000
			tmrScorbit.UserValue = 0
			tmrScorbit.Enabled=True 
			Scorbit.UploadLog = ScorbitUploadLog
		End if 
	End if 

	bGameReady = True


	if DMDType = 0 Then 
		BallHandlingQueue.Add "loadBG","LoadBG",24,2500,0,0,0,False
	Else
		BallHandlingQueue.Add "loadBG","LoadBG",24,1000,0,0,0,False
	End If
	BallHandlingQueue.Add "UpdateModeMessages.Enabled = 1","UpdateModeMessages.Enabled = 1",24,2000,0,0,0,False

End Sub 'end PUPINIT

	Const cWhite = 	16777215
	Const cRed = 	397512
	Const cGold = 	1604786
	Const cGold2 = 46079
	Const cGreen = 32768
	Const cGrey = 	8421504
	Const cYellow = 65535
	Const cOrange = 33023
	Const cPurple = 16711808
	Const cBlue = 16711680
	Const cLightBlue = 16744448
	Const cBoltYellow = 2148582
	Const cLightGreen = 9747818
	Const cBlack = 0
	Const cPink = 12615935
	Const cSilver = 8421504


Const dmddef="Calvera Personal Use Only"
Const dmdNum="Calvera Personal Use Only"
'pages
Const pDMDBlank = 0
Const pScores = 1
Const pAttract = 2
Const pPrevScores = 3
Const pCredits = 4
Const pSlotMachine = 5
Const pBonus = 6
Const pEvent = 7
Const pHighScore = 8


Sub CheckPupVersion

	Dim strPupVersion
	strPupVersion = PuPlayer.GetVersion
	strPupVersion = Replace(strPupVersion, ".", "")
	strPupVersion = mid (strPupVersion, 1,3)
    strPupVersion = CDbl(strPupVersion)

	If strPupVersion => 150 then
		exit sub
	Else
		msgbox "This table requires PuP Player version 1.5 or greater.  Please update your pup install to play the table", 0
		Table1_Exit	
	End if

End sub

Sub pSetPageLayouts
	Dim i

	pDMDAlwaysPAD		'we pad all text with space before and after for shadow clipping/etc


	'pupCreateLabelImageBG "AttractCredits","PuPOverlays\\Credits1.png",0,0,100,100,88,0

	pupCreateLabelImageDMD "DemonGIF","GIFS\\Demon.gif",7.75,37.25,13,35,1,0


	PuPlayer.playlistadd pDMDVideo,"Background", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"AttractMode", 1 , 0
    PuPlayer.playlistadd pDMDVideo,"Skillshot", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"BallSave", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"BallLost", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"ExtraBall", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"Combos", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"GameOver", 1 , 0
    PuPlayer.playlistadd pDMDVideo,"BlizMultiball", 1 , 0
    PuPlayer.playlistadd pDMDVideo,"Misc", 1 , 0
    PuPlayer.playlistadd pDMDVideo,"Multipliers", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"Mode", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"Mystery", 1 , 0
    PuPlayer.playlistadd pDMDVideo,"PowerPulse", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"Multiball", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"Jackpot", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"BallLock", 1 , 0
	PuPlayer.playlistadd pDMDVideo,"Wizard", 1 , 0
	PuPlayer.playlistadd pTopper,"Topper", 1 , 0

'labelNew <screen#>, <Labelname>, <fontName>,<size%>,<colour>,<rotation>,<xalign>,<yalign>,<xpos>,<ypos>,<PageNum>,<visible>
'***********************************************************************'
'<screen#>, in standard we’d set this to pDMD ( or 1)
'<Labelname>, your name of the label. keep it short no spaces (like 8 chars) although you can call it anything really. When setting the label you will use this labelname to access the label.
'<fontName> Windows font name, this must be exact match of OS front name. if you are using custom TTF fonts then double check the name of font names.
'<size%>, Height as a percent of display height. 20=20% of screen height.
'<colour>, integer value of windows color.
'<rotation>, degrees in tenths   (900=90 degrees)
'<xAlign>, 0= horizontal left align, 1 = center horizontal, 2= right horizontal
'<yAlign>, 0 = top, 1 = center, 2=bottom vertical alignment
'<xpos>, this should be 0, but if you want to ‘force’ a position you can set this. it is a % of horizontal width. 20=20% of screen width.
'<ypos> same as xpos.
'<PageNum> IMPORTANT… this will assign this label to this ‘page’ or group.
'<visible> initial state of label. visible=1 show, 0 = off.

	if DMDType = 2 Then
		pupCreateLabelImage "ScorbitQRicon1","PuPOverlays\\QRcodeS.png",0,0,100,100,1,0
		pupCreateLabelImageBG "ScorbitQR1","PuPOverlays\\QRcode.png",0,0,100,100,77,0

		pupCreateLabelImage "ScorbitQRicon2","PuPOverlays\\QRcodeB.png",0,0,100,100,1,0
		pupCreateLabelImageBG "ScorbitQR2","PuPOverlays\\QRclaim.png",0,0,100,100,77,0
	Else
		pupCreateLabelImage "ScorbitQRicon1","PuPOverlays\\QRcodeS.png",0,0,100,100,1,0
		pupCreateLabelImageBG "ScorbitQR1","PuPOverlays\\QRcode.png",0,0,100,100,1,0

		pupCreateLabelImage "ScorbitQRicon2","PuPOverlays\\QRcodeB.png",0,0,100,100,1,0
		pupCreateLabelImageBG "ScorbitQR2","PuPOverlays\\QRclaim.png",0,0,100,100,1,0
	End If



	pupCreateLabelImageBG "BGlass", "PuPOverlays\\Card0.png",0,0,100,100,1,0
	pupCreateLabelImageDMD "DMDOverlay", "PuPOverlays\\DefaultDMD.png",0,0,100,100,1,0

	pupCreateLabelImageDMD "BlizzTimerImage", "PuPOverlays\\BlizzTimer.png",0,0,100,100,1,0


    pupCreateLabelImageDMD "BonusMP2x","PupOverlays\\2x.png",0,0,100,100,1,0	' 1
	pupCreateLabelImageDMD "BonusMP3x","PupOverlays\\3x.png",0,0,100,100,1,0	' 2
	pupCreateLabelImageDMD "BonusMP4x","PupOverlays\\4x.png",0,0,100,100,1,0	' 3
	pupCreateLabelImageDMD "BonusMP5x","PupOverlays\\5x.png",0,0,100,100,1,0	' 4


	' USED FOR DEBUG ONLY
	PuPlayer.LabelNew pDMD, "Line1b",dmdDef,			15,cGold	,0,1,1, 50,32,	pScores,0
	PuPlayer.LabelNew pDMD, "Line2b",dmdDef, 			15,cGold	,0,1,1, 50,55,	pScores,0



'Attract
	PuPlayer.LabelNew pDMD, "Attract2a", dmddef, 12, cWhite, 0, 1, 1, 50, 40, 1, 0
	PuPlayer.LabelNew pDMD, "Attract2b", dmddef, 12, cWhite, 0, 1, 1, 50, 56, 1, 0
	pDMDLabelSetBorder "Attract2a",cRed,6,6,1
	pDMDLabelSetBorder "Attract2b",cRed,6,6,1

'Instant Info
	PuPlayer.LabelNew pDMD, "InstantInfo2a", dmddef, 12, cWhite, 0, 1, 1, 50, 40, 1, 0
	PuPlayer.LabelNew pDMD, "InstantInfo2b", dmddef, 12, cWhite, 0, 1, 1, 50, 56, 1, 0
	pDMDLabelSetBorder "InstantInfo2a",cWhite,4,4,1
	pDMDLabelSetBorder "InstantInfo2b",cWhite,4,4,1


'page 1
	PuPlayer.LabelNew pDMD,"HSLine1",	             dmddef,16,cPurple   ,0,1,1,50,40,1,0
	PuPlayer.LabelNew pDMD,"HSLine2",	             dmddef,16,cPurple   ,0,1,1,50,55,1,0
	pDMDLabelSetBorder "HSLine1",cBlack,6,6,1
	pDMDLabelSetBorder "HSLine2",cBlack,6,6,1

	PuPlayer.LabelNew pDMD, "ComboLine2a", dmddef, 14, cWhite, 0, 1, 1, 50, 53, 1, 0
	PuPlayer.LabelNew pDMD, "ComboLine2b", dmddef, 14, cYellow, 0, 1, 1, 50, 60, 1, 0
	pDMDLabelSetBorder "ComboLine2a",cBlack,6,6,1
	pDMDLabelSetBorder "ComboLine2b",cBlack,6,6,1

	'USED FOR TEMP UPDATES
	PuPlayer.LabelNew pDMD, "Splash", dmddef, 18, cWhite, 0, 1, 1, 50, 32, pScores, 0
	PuPlayer.LabelNew pDMD, "Splash2a", dmddef, 12, cGold, 0, 1, 1, 50, 34, pScores, 0
	PuPlayer.LabelNew pDMD, "Splash2b", dmddef, 12, cGold, 0, 1, 1, 50, 50, pScores, 0
	PuPlayer.LabelNew pDMD, "Splash3a", dmddef, 10, cGold, 0, 1, 1, 50, 30, pScores, 0
	PuPlayer.LabelNew pDMD, "Splash3b", dmddef, 10, cGold, 0, 1, 1, 50, 42, pScores, 0
	PuPlayer.LabelNew pDMD, "Splash3c", dmddef, 10, cGold, 0, 1, 1, 50, 54, pScores, 0

	pDMDLabelSetBorder "Splash",cBlack,6,6,1
	pDMDLabelSetBorder "Splash2A",cBlack,6,6,1
	pDMDLabelSetBorder "Splash2B",cBlack,6,6,1
	pDMDLabelSetBorder "Splash3a",cBlack,6,6,1
	pDMDLabelSetBorder "Splash3b",cBlack,6,6,1
	pDMDLabelSetBorder "Splash3c",cBlack,6,6,1

	PuPlayer.LabelNew pDMD, "BlizzTimerValue", dmddef, 26, cBlue, 0, 1, 1, 30, 50, pScores, 0
	pDMDLabelSetBorder "BlizzTimerValue",cPurple,2,2,1

	'USED FOR LONG TERM DISPLAYS
	PuPlayer.LabelNew pDMD, "Event3A", dmddef, 10, cYellow, 0, 1, 1, 65, 30, pScores, 0
	PuPlayer.LabelNew pDMD, "Event3B", dmddef, 10, cYellow, 0, 1, 1, 65, 40, pScores, 0
	PuPlayer.LabelNew pDMD, "Event3Ca", dmddef, 10, cRed, 0, 1, 1, 52, 50, pScores, 0
	PuPlayer.LabelNew pDMD, "Event3C", dmddef, 10, cYellow, 0, 1, 1, 66, 50, pScores, 0

	pDMDLabelSetBorder "Event3A",cOrange,2,2,1
	pDMDLabelSetBorder "Event3B",cOrange,2,2,1
	pDMDLabelSetBorder "Event3Ca",cYellow,2,2,1
	pDMDLabelSetBorder "Event3C",cOrange,2,2,1


	'USED FOR JACKPOTS
	PuPlayer.LabelNew pDMD, "SplashJP2a", dmddef, 12, cRed, 0, 1, 1, 50, 34, pScores, 0
	PuPlayer.LabelNew pDMD, "SplashJP2b", dmddef, 12, cRed, 0, 1, 1, 50, 50, pScores, 0

	'USED FOR JUKEBOX
	PuPlayer.LabelNew pDMD, "JukeBox2a", dmddef, 12, cPink, 0, 1, 1, 50, 34, pScores, 0
	PuPlayer.LabelNew pDMD, "JukeBox2b", dmddef, 12, cOrange, 0, 1, 1, 50, 50, pScores, 0


	PuPlayer.LabelNew pDMD,"CurrScore",         dmddef,5,cWhite   ,0,1,1, 50,88.5,1,0
	PuPlayer.LabelNew pDMD,"CurrName",	             dmddef,5,cPink   ,0,0,0,50,82.5,1,0
	PuPlayer.LabelNew pDMD,"Position1Score",	             dmddef,5,cWhite   ,0,0,1,33,79.5,1,0
'	PuPlayer.LabelNew pDMD,"Position2Name",	             dmddef,5,cBlack   ,0,0,0,20,89,1,0
	PuPlayer.LabelNew pDMD,"Position2Score",	             dmddef,5,cWhite  ,0,0,1,20,94,1,0
'	PuPlayer.LabelNew pDMD,"Position3Name",	             dmddef,5,cBlack   ,0,0,0,50,89,1,0
	PuPlayer.LabelNew pDMD,"Position3Score",	             dmddef,5,cWhite   ,0,0,1,50,94,1,0
'	PuPlayer.LabelNew pDMD,"Position4Name",	             dmddef,5,cBlack   ,0,0,0,80,89,1,0
	PuPlayer.LabelNew pDMD,"Position4Score",	             dmddef,5,cWhite   ,0,0,1,80,94,1,0

	pDMDLabelSetBorder "CurrScore",cRed,5,5,1
	pDMDLabelSetBorder "CurrName",cRed,5,5,1

	PuPlayer.LabelNew pDMD,"AlbumNum",         dmddef,5,cYellow   ,0,0,0, 4.2,65,1,0
	PuPlayer.LabelNew pDMD,"ComboNum",	             dmddef,5,cOrange   ,0,0,0,43,91,1,0

	pDMDLabelSetBorder "AlbumNum",cRed,5,5,1
	pDMDLabelSetBorder "ComboNum",cRed,4,4,1

	pDMDLabelSetBorder "Position1Score",cRed,3,3,1
	pDMDLabelSetBorder "Position2Score",cRed,4,4,1
	pDMDLabelSetBorder "Position3Score",cRed,5,5,1
	pDMDLabelSetBorder "Position4Score",cRed,5,5,1
'	pDMDLabelSetBorder "Position1Name",cBlack,5,5,1
'	pDMDLabelSetBorder "Position2Name",cBlack,5,5,1
'	pDMDLabelSetBorder "Position3Name",cBlack,5,5,1
'	pDMDLabelSetBorder "Position4Name",cBlack,5,5,1

	' EOB Sequence
	PuPlayer.LabelNew pDMD, "Event5A", dmddef, 10, cWhite, 0, 0, 1, 25, 20, pScores, 0
	PuPlayer.LabelNew pDMD, "Event5B", dmddef, 10, cPink, 0, 0, 1, 25, 30, pScores, 0
	PuPlayer.LabelNew pDMD, "Event5C", dmddef, 10, cGold2, 0, 0, 1, 25, 40, pScores, 0
	PuPlayer.LabelNew pDMD, "Event5D", dmddef, 10, cGrey, 0, 0, 1, 25, 50, pScores, 0
	PuPlayer.LabelNew pDMD, "Event5E", dmddef, 10, cOrange, 0, 0, 1, 25, 60, pScores, 0
	PuPlayer.LabelNew pDMD, "Event5F", dmddef, 10, cYellow, 0, 1, 1, 50, 70, pScores, 0

	pDMDLabelSetBorder "Event5A",cRed,6,6,1
	pDMDLabelSetBorder "Event5B",cRed,6,6,1
	pDMDLabelSetBorder "Event5C",cRed,6,6,1
	pDMDLabelSetBorder "Event5D",cRed,6,6,1
	pDMDLabelSetBorder "Event5E",cRed,6,6,1
	pDMDLabelSetBorder "Event5F",cRed,6,6,1


	PuPlayer.LabelNew pDMD,"BallValue",	             dmddef,5,cYellow  ,0,0,0,91.25,63,1,0
	pDMDLabelSetBorder "BallValue",cRed,4,4,1

	PuPlayer.LabelNew pDMD,"MainModeTimerValue",	             dmddef,8,cRed   ,0,1,1,3.8,76,1,0
	pDMDLabelSetBorder "MainModeTimerValue",cWhite,5,5,1


	PuPlayer.LabelNew pDMD,"BMValue",	             dmddef,12,cWhite   ,0,1,1,21.75,87.5,1,0
	pDMDLabelSetBorder "BMValue",cBlack,4,4,1

	PuPlayer.LabelNew pDMD,"PFMValue",	             dmddef,12,cWhite   ,0,1,1,78.75,87.5,1,0
	pDMDLabelSetBorder "PFMValue",cBlack,4,4,1

	PuPlayer.LabelNew pDMD,"TimerValue",	             dmddef,12,cRed   ,0,1,1,78.75,87.5,1,0
	pDMDLabelSetBorder "TimerValue",cBlack,4,4,1

	PuPlayer.LabelNew pDMD,"LockValue",	             dmddef,12,cRed   ,0,1,1,78.75,87.5,1,0
	pDMDLabelSetBorder "LockValue",cBlack,4,4,1

	PuPlayer.LabelNew pDMD,"TESTQ",	             dmddef,20,cWhite  ,0,1,1,30,30.1,pScores,0
	PuPlayer.LabelNew pDMD,"TESTQTitle",	     dmddef,10,cWhite  ,0,1,1,30,30.1,pScores,0



End Sub

Sub pDMDStartUP
'do stuff fancy pants on first run
'	pInAttract = True
	pDMDSetPage(pScores)

	if ScorbitActive Then
		dbg2 "Calling SCORBIT check pairing"
		BallHandlingQueue.Add "CheckPairing","CheckPairing",24,2000,0,0,0,False
	End If
end Sub 'end DMDStartup

Dim objIEDebugWindow
Sub Dbg( myDebugText )
' Uncomment the next line to turn off debugging
Exit Sub

If Not IsObject( objIEDebugWindow ) Then
Set objIEDebugWindow = CreateObject( "InternetExplorer.Application" )
objIEDebugWindow.Navigate "about:blank"
objIEDebugWindow.Visible = True
objIEDebugWindow.ToolBar = False
objIEDebugWindow.Width = 600	
objIEDebugWindow.Height = 900
objIEDebugWindow.Left = 2100
objIEDebugWindow.Top = 100
Do While objIEDebugWindow.Busy
Loop
objIEDebugWindow.Document.Title = "VPX Debug Window"
objIEDebugWindow.Document.Body.InnerHTML = "<b>Blizzard of Ozz Debug Window -TimeStamp: " & GameTime& "</b></br>"
End If

objIEDebugWindow.Document.Body.InnerHTML = objIEDebugWindow.Document.Body.InnerHTML & myDebugText & " --TimeStamp:<b> " & GameTime & "</b><br>" & vbCrLf
End Sub

'***********************************************************PinUP Player DMD Helper Functions

Sub pTranslatePos(Byref xpos, byref ypos)  'if using uUseFramePos then all coordinates are based on framesize
   xpos=int(xpos/pFrameSizeX*10000) / 100
   ypos=int(ypos/pFrameSizeY*10000) / 100
end Sub

Sub pTranslateY(Byref ypos)           'if using uUseFramePos then all heights are based on framesize
   ypos=int(ypos/pFrameSizeY*10000) / 100
end Sub

Sub pTranslateX(Byref xpos)           'if using uUseFramePos then all heights are based on framesize
   xpos=int(xpos/pFrameSizeX*10000) / 100
end Sub

Dim PBackglassCurPage
Sub pBackglassSetPage(pagenum)    
    PuPlayer.LabelShowPage pBackglass,pagenum,0,""   'set page to blank 0 page if want off
    PBackglassCurPage=pagenum
end Sub

Sub pBackglassLabelShow(labName)
PuPlayer.LabelSet pBackglass,labName,"",1,""   
end sub

Sub pBackglassLabelHide(labName)
PuPlayer.LabelSet pBackglass,labName,"",0,""   
end sub

Sub pupCreateLabelImageDMD(lName, lFilename,xpos, ypos, Iwidth, Iheight, pagenum, lvis)
	PuPlayer.LabelNew pDMDFull,lName ,"",50,RGB(100,100,100),0,1,1,1,1,pagenum,lvis
	PuPlayer.LabelSet pDMDFull,lName,lFilename,lvis,"{'mt':2,'width':"&IWidth&",'height':"&Iheight&",'xpos':"&xpos&",'ypos':"&ypos&"}"
end Sub

sub pDMDHideAnimate(labName) 
   PuPlayer.LabelSet pDMDFull,labName,"",0,"" 
End Sub

Sub SupressModeMessages(mTime)
	debug.print "Calling Supress : " &mTime

'	if Battle(CurrentPlayer, 0) = 0 Then Exit Sub
'	debug.print "IN Supress : " &mTime
	bSupressModeMessages = True
	DMDQueue.Add "bSupressModeMessages = False","bSupressModeMessages = False",70,mTime,0,0,0,False

	if bBlizzardPrepMode Then
		HideBlizzTimer
		PuPlayer.LabelSet pDMD,"Event3A","",0,""
		PuPlayer.LabelSet pDMD,"Event3B","",0,""
		PuPlayer.LabelSet pDMD,"Event3C","",0,""
		PuPlayer.LabelSet pDMD,"Event3Ca","",0,""
		DMDQueue.Add "pDMDlabelshow ""BlizzTimerImage"" ","pDMDlabelshow ""BlizzTimerImage"" ",45,mTime,0,0,0,False
	End If

	DMDQueue.Add "UpdateModeProgress False","UpdateModeProgress False",45,mTime+300,0,0,0,False
	GeneralPupQueue.Add "PlayModeVideo "&Battle(CurrentPlayer, 0),"PlayModeVideo "&Battle(CurrentPlayer, 0),60,mTime+300,0,0,0,False

	if Battle(CurrentPlayer,1) = 2 Then 
		DMDQueue.Add "UpdateSpinProgress","UpdateSpinProgress",45,mTime+300,0,0,0,False		
	End If
End Sub


Sub pDMDSplashTwoLines(msgText,msgText2,timeSec,mColor)'Para uso normal del DMD
	PuPlayer.LabelShowPage pDMD,1,timeSec,""
	SupressModeMessages timeSec*1000

	PuPlayer.LabelSet pDMD,"Event3A","",0,""
	PuPlayer.LabelSet pDMD,"Event3B","",0,""
	PuPlayer.LabelSet pDMD,"Event3C","",0,""
	PuPlayer.LabelSet pDMD,"Event3Ca","",0,""

	PuPlayer.LabelSet pDMD,"Splash2a",msgText,1,"{'mt':2,'color': " & mColor &" }"  
	PuPlayer.LabelSet pDMD,"Splash2b",msgText2,1,"{'mt':2,'color': " & mColor &" }"  
	pDMDLabelSetBorder "Splash2a",cWhite,3,3,1
	pDMDLabelSetBorder "Splash2b",cWhite,3,3,1
end Sub  

Sub ClearTwoLines
	PuPlayer.LabelSet pDMD,"Splash2a","",0,""  
	PuPlayer.LabelSet pDMD,"Splash2b","",0,""
End Sub

Sub pDMDLabelSet(labName,LabText)
PuPlayer.LabelSet pDMD,labName,LabText,1,""   
end sub


Sub pDMDLabelHide(labName)
PuPlayer.LabelSet pDMD,labName,"`u`",0,""   
end sub

Sub pDMDLabelShow(labName)
PuPlayer.LabelSet pDMD,labName,"`u`",1,""   
end sub

Sub pDMDLabelVisible(labName, isVis)
PuPlayer.LabelSet pDMD,labName,"`u`",isVis,""   
end sub

Sub pDMDLabelSendToBack(labName)
PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'zback': 1 }"   
end sub

Sub pDMDLabelSendToFront(labName)
PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'ztop': 1 }"   
end sub

sub pDMDLabelSetPos(labName, byVal xpos, byVal ypos)
   if pUseFramePos=1 Then pTranslatePos xpos,ypos
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'xpos':"&xpos& ",'ypos':"&ypos&"}"    
end sub

sub pDMDLabelSetSizeImage(labName, byVal lWidth, byVal lHeight)
   if pUseFramePos=1 Then pTranslatePos lWidth,lHeight
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'width':"& lWidth & ",'height':"&lHeight&"}" 
end sub

sub pDMDLabelSetSizeText(labName, byVal fHeight)
   if pUseFramePos=1 Then pTranslateY fHeight
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'size':"&fHeight&"}" 
end sub

sub pDMDLabelSetAutoSize(labName, byVal lWidth, byVal lHeight)
   if pUseFramePos=1 Then pTranslatePos lWidth,lHeight
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'autow':"& lWidth & ",'autoh':"&lHeight&"}" 
end sub

sub PDMDLabelSetAlign(labName,xAlign, YAlign)  '0=left 1=center 2=right,  note you should use center as much as possible because some things like rotate/zoom/etc only look correct with center align!
    PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'xalign':"& xAlign & ",'yalign':"&yAlign&"}"     
end sub

sub pDMDLabelStopAnis(labName)    'stop any pup animations on label/image (zoom/flash/pulse).  this is not about animated gifs
     PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'stopani':1 }" 
end sub

sub pDMDLabelSetRotateText(labName, fAngle)  ' in tenths.  so 900 is 90 degrees.
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'rotate':"&fAngle&"}" 
end sub

sub pDMDLabelSetRotate(labName, fAngle)  ' in tenths.  so 900 is 90 degrees. rotate support for images too.  note images must be aligned center to rotate properly(default)
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'rotate':"&fAngle&"}" 
end sub

sub pDMDLabelSetZoom(labName, fFactor)  ' fFactor is 120 for 120% of current height, 80% etc...
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'zoom':"&fFactor&"}" 
end sub

sub pDMDLabelSetColor(labName, lCol)
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'color':"&lCol&"}" 
end sub

sub pDMDLabelSetAlpha(labName, lAlpha)  '0-255  255=full, 0=blank
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'alpha':"&lAlpha&"}" 
end sub

sub pDMDLabelSetColorGradient(labName, byVal startCol, byVal EndCol)
dim GS: GS=1
if startCol=EndCol Then GS=0  'turn grad off is same colors.
PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'color':"&startCol&" ,'gradstate':"&GS&" , 'gradcolor':"&endCol&"}" 
end sub

sub pDMDLabelSetColorGradientPercent(labName, byVal startCol, byVal EndCol, byVal StartPercent)
if startCol=EndCol Then StartPercent=0  'turn grad off is same colors.
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'color':"&startCol&" ,  'gradstate':"&StartPercent&", 'gradcolor':"&endCol&"}" 
end sub

sub pDMDLabelSetGrayScale(labName, isGray)  'only on image objects.  will show as grayscale.  1=gray filter on 0=off normal mode
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'grayscale':"&isGray&"}" 
end sub

sub pDMDLabelSetFilter(labName, fMode)  ''fmode 1-5 (invertRGB, invert,grayscale,invertalpha,clear),blur)
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'filter':"&fmode&"}" 
end sub

Sub pDMDLabelFlashFilter(LabName,byVal timeSec,fMode)   'timeSec in ms  'fmode 1-5 (invertRGB, invert,grayscale,invertalpha,clear,blur)
    if timeSec<20 Then timeSec=timeSec*1000
    PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':1,'at':9,'fq':150,'len':" & (timeSec) & ",'fm':" & fMode & "}"   
end sub



sub pDMDLabelSetShadow(labName,lCol,offsetx,offsety,isVis)  ' shadow of text
dim ST: ST=1 : if isVIS=false Then St=0
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'shadowcolor':"&lCol&",'shadowtype': "&ST&", 'xoffset': "&offsetx&", 'yoffset': "&offsety&"}"
end sub

sub pDMDLabelSetBorder(labName,lCol,offsetx,offsety,isVis)   'outline/border around text.
dim ST: ST=2 : if isVIS=false Then St=0
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'shadowcolor':"&lCol&",'shadowtype': "&ST&", 'xoffset': "&offsetx&", 'yoffset': "&offsety&"}"
end sub



'animations   'pDMDLabelPulseText "pulsetext","jackpot",4000,rgb(100,0,0)

sub pDMDLabelPulseText(LabName,LabValue,mLen,mColor)       'mlen in ms
    PuPlayer.LabelSet pDMD,labName,LabValue,1,"{'mt':1,'at':4,'hstart':80,'hend':120,'len':" & (mLen) & ",'pspeed': 0,'fc':" & mColor & ",'aa':0 }"
end Sub

sub pDMDLabelPulseText2(LabName,LabValue,mLen,mColor)       'mlen in ms
    PuPlayer.LabelSet pDMD,labName,LabValue,1,"{'mt':1,'at':4,'hstart':70,'hend':170,'len':" & (mLen) & ",'pspeed': 0,'fc':" & mColor & ",'aa':0 }"
end Sub

sub pDMDLabelPulseNumber(LabName,LabValue,mLen,mColor,pNumStart,pNumEnd,pNumformat)   'pnumformat 0 no format, 1 with thousands  mLen=ms
     PuPlayer.LabelSet pDMD,labName,LabValue,1,"{'mt':1,'at':4,'hstart':80,'hend':120,'len':" & (mLen) & ",'pspeed': 0,'fc':" & mColor & ",'numstart':"&pNumStart&",'numend' :"&pNumEnd&", 'numformat':"&pNumFormat&",'aa':0 }"    
end Sub

sub pDMDLabelPulseImage(LabName,mLen,isVis)       'mlen in ms isVis is state after animation
    PuPlayer.LabelSet pDMD,labName,"`u`",isVis,"{'mt':1,'at':4,'hstart':80,'hend':120,'len':" & (mLen) & ",'pspeed': 0 }"
end Sub

sub pDMDLabelPulseImage2(LabName,mLen,isVis)       'mlen in ms isVis is state after animation
    PuPlayer.LabelSet pDMD,labName,"`u`",isVis,"{'mt':1,'at':4,'hstart':70,'hend':180,'len':" & (mLen) & ",'pspeed': 0 }"
end Sub

sub pDMDLabelPulseTextEX(LabName,LabValue,mLen,mColor,isVis,zStart,zEnd)       'mlen in ms  same subs as above but youspecifiy zoom start and zoom end in % height of original font.
    PuPlayer.LabelSet pDMD,labName,LabValue,isVis,"{'mt':1,'at':4,'hstart':"&zStart&",'hend':"&zEnd&",'len':" & (mLen) & ",'pspeed': 0,'fc':" & mColor & ",'aa':0 }"
end Sub

sub pDMDLabelPulseNumberEX(LabName,LabValue,mLen,mColor,pNumStart,pNumEnd,pNumformat,isVis,zStart,zEnd)   'pnumformat 0 no format, 1 with thousands  mLen=ms
     PuPlayer.LabelSet pDMD,labName,LabValue,isVis,"{'mt':1,'at':4,'hstart':"&zStart&",'hend':"&zEnd&",'len':" & (mLen) & ",'pspeed': 0,'fc':" & mColor & ",'numstart':"&pNumStart&",'numend' :"&pNumEnd&", 'numformat':"&pNumFormat&",'aa':0}"    
end Sub

sub pDMDLabelPulseImageEX(LabName,mLen,isVis,zStart,zEnd)       'mlen in ms isVis is state after animation
    PuPlayer.LabelSet pDMD,labName,"`u`",isVis,"{'mt':1,'at':4,'hstart':"&zStart&",'hend':"&zEnd&",'len':" & (mLen) & ",'pspeed': 0 }"
end Sub

sub pDMDLabelWiggleText(LabName,LabValue,mLen,mColor)       'mlen in ms  zstart MUST be less than zEND.  -40 to 40 for example
    PuPlayer.LabelSet pDMD,labName,LabValue,1,"{'mt':1,'at':8,'rstart':-45,'rend':45,'len':" & (mLen) & ",'rspeed': 5,'fc':" & mColor & ",'aa':0 }"
end Sub

sub pDMDLabelWiggleTextEX(LabName,LabValue,mLen,mColor,isVis,zStart,zEnd)       'mlen in ms  zstart MUST be less than zEND.  -40 to 40 for example
    PuPlayer.LabelSet pDMD,labName,LabValue,isVis,"{'mt':1,'at':8,'rstart':"&zStart&",'rend':"&zEnd&",'len':" & (mLen) & ",'rspeed': 5,'fc':" & mColor & ",'aa':0 }"
end Sub

sub pDMDLabelWiggleImage(LabName,mLen,isVis)         'mlen in ms  zstart MUST be less than zEND.  -40 to 40 for example
    PuPlayer.LabelSet pDMD,labName,"`u`",isVis,"{'mt':1,'at':8,'rstart':-45,'rend':45,'len':" & (mLen) & ",'rspeed': 5,'fc':" & 0 & ",'aa':0 }"
end Sub

sub pDMDLabelWiggleImageEX(LabName,mLen,isVis,zStart,zEnd)       'mlen in ms  zstart MUST be less than zEND.  -40 to 40 for example
    PuPlayer.LabelSet pDMD,labName,"`u`",isVis,"{'mt':1,'at':8,'rstart':"&zStart&",'rend':"&zEnd&",'len':" & (mLen) & ",'rspeed': 5,'fc':" & 0 & ",'aa':0 }"
end Sub

sub pDMDLabelClone(LabName,LabValue,mLen,mColor,pX,pY)   'px,PY  use with temp label to repeat control.
     if pUseFramePos=1 Then pTranslatePos pX,pY
     PuPlayer.LabelSet pDMD,labName,LabValue,0,"{'mt':1,'at':10, 'len':" & (mLen) & ", 'fc':" & mColor & ",'xp':"&pX&",'yp' :"&pY&" ,'ad':1, 'dl':1000 }"    
end Sub

sub pDMDLabelCloneDelay(LabName,LabValue,mLen,mColor,pX,pY,dL)   'px,PY  use with temp label to repeat control.  dL delay ms
     if pUseFramePos=1 Then pTranslatePos pX,pY
     PuPlayer.LabelSet pDMD,labName,LabValue,0,"{'mt':1,'at':10, 'len':" & (mLen) & ", 'fc':" & mColor & ",'xp':"&pX&",'yp' :"&pY&" ,'ad':1, 'dl':"&dL&" }"    
end Sub


sub pDMDPNGAnimate(labName,cSpeed)  'speed is frame timer, 0 = stop animation  100 is 10fps for animated png and gif nextframe timer.
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'animate':"&cSpeed&"}" 
end sub

sub pDMDPNGAnimateEx(labName,startFrame,endFrame,LoopMode)  'sets up the apng/gif settings before you call animate.  if you set start/end frame same if will display that frame, set start to -1 to reset settings.
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'gifstart':"&startFrame&",'gifend':"&endFrame&",'gifloop':"&loopMode&" }"          'gifstart':3, 'gifend':10, 'gifloop': 1
end sub

sub pDMDPNGShowFrame(labName,fFrame)  'in a animated png/gif, will set it to an individual frame so you could use as an imagelist control
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'gifstart':"&fFrame&",'gifend':"&fFrame&" }"          '
end sub

sub pDMDPNGAnimateOnce(labName,cSpeed)  'will show an animated gif/png and then hide when done, overrides loop to force stop at end.
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'animate':"&cSpeed&", 'gifloop': 0 , 'aniendhide':1 }" 
end sub

sub pDMDPNGAnimateReset(labName)  'speed is frame timer, 0 = stop animation  100 is 10fps for animated png and gif nextframe timer, this will show anigif and hide at end no loop
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'animate':0, 'gifloop': 1 , 'aniendhide':0 , 'gifstart':-1}" 
end sub

sub pDMDPNGAnimateOnceAndDispose(labName,fName, cSpeed)  'speed is frame timer, 0 = stop animation  100 is 10fps for animated png and gif nextframe timer, this will show anigif and hide at end no loop
   PuPlayer.LabelSet pDMD,labName,fName,1,"{'mt':2,'animate':"&cSpeed&", 'gifloop': 0 , 'aniendhide':1, 'anidispose':1 }" 
end sub

sub pDMDLabelSetOutShadow(labName, lCol,offsetx,offsety,isOutline,isVis)
   PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'shadowcolor':"&lCol&",'shadowstate': "&isVis&", 'xoffset': "&offsetx&", 'yoffset': "&offsety&", 'outline': "&isOutline&"}"
end sub

sub pDMDLabelMoveHorz(LabName,LabValue,mLen,mColor,pMoveStart,pMoveEnd)   'pmovestart is -1= left-off 0=current pos 1=right-off    or can use % 
     PuPlayer.LabelSet pDMD,labName,LabValue,1,"{'mt':1,'at':2, 'len':" & (mLen) & ", 'fc':" & mColor & ",'xps':"&pMoveStart&",'xpe' :"&pMoveEnd&", 'tt':2,'ad':1 }"    
end Sub

sub pDMDLabelMoveVert(LabName,LabValue,mLen,mColor,pMoveStart,pMoveEnd)   'pmovestart is -1= left-off 0=current pos 1=right-off   or can use %  
     PuPlayer.LabelSet pDMD,labName,LabValue,1,"{'mt':1,'at':2, 'len':" & (mLen) & ", 'fc':" & mColor & ",'yps':"&pMoveStart&",'ype' :"&pMoveEnd&", 'tt':2,'ad':1 }"    
end Sub

sub pDMDLabelMoveTO(LabName,LabValue,mLen,mColor,byVal pStartX,byVal pStartY,byVal pEndX,byVal pEndY)   'pmovestart is -1= left-off 0=current pos 1=right-off
     if pUseFramePos=1 AND (pStartX+pStartY+pEndx+pendY)>4 Then 
                       pTranslatePos pStartX,pStartY
                       pTranslatePos pEndX,pEndY
     end IF 
     PuPlayer.LabelSet pDMD,labName,LabValue,1,"{'mt':1,'at':2, 'len':" & (mLen) & ", 'fc':" & mColor & ",'xps':"&pStartX&",'xpe' :"&pEndX& ",'yps':"&pStartY&",'ype' :"&pEndY&", 'tt':2 ,'ad':1}"    
end Sub

sub pDMDLabelMoveHorzFade(LabName,LabValue,mLen,mColor,pMoveStart,pMoveEnd)   'pmovestart is -1= left-off 0=current pos 1=right-off, or can use %
     PuPlayer.LabelSet pDMD,labName,LabValue,0,"{'mt':1,'at':2, 'len':" & (mLen) & ", 'fc':" & mColor & ",'xps':"&pMoveStart&",'xpe' :"&pMoveEnd&", 'tt':2 ,'ad':1, 'af':700}"    
end Sub

sub pDMDLabelMoveVertFade(LabName,LabValue,mLen,mColor,pMoveStart,pMoveEnd)   'pmovestart is -1= left-off 0=current pos 1=right-off  or can use %   
     PuPlayer.LabelSet pDMD,labName,LabValue,0,"{'mt':1,'at':2, 'len':" & (mLen) & ", 'fc':" & mColor & ",'yps':"&pMoveStart&",'ype' :"&pMoveEnd&", 'tt':2 ,'ad':1, 'af':700}"    
end Sub

sub pDMDLabelMoveTOFade(LabName,LabValue,mLen,mColor,byVal pStartX,byVal pStartY,byVal pEndX,byVal pEndY)   'pmovestart is -1= left-off 0=current pos 1=right-off
     if pUseFramePos=1 AND (pStartX+pStartY+pEndx+pendY)>4 Then 
                       pTranslatePos pStartX,pStartY
                       pTranslatePos pEndX,pEndY
     end IF 
     PuPlayer.LabelSet pDMD,labName,LabValue,0,"{'mt':1,'at':2, 'len':" & (mLen) & ", 'fc':" & mColor & ",'xps':"&pStartX&",'xpe' :"&pEndX& ",'yps':"&pStartY&",'ype' :"&pEndY&", 'tt':6 ,'ad':1, 'af':700}"    
end Sub





sub pDMDLabelFadeOut(LabName,mLen)   'alpha is 255 max, 0=clear.  
     PuPlayer.LabelSet pDMD,labName,"`u`",0,"{'mt':1,'at':5,'astart':255,'aend':0,'len':" & (mLen) & " }"    
end Sub

sub pDMDLabelFadeIn(LabName,mLen)    'alpha is 255 max, 0=clear. 
     PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':1,'at':5,'astart':0,'aend':255,'len':" & (mLen) & " }"    
end Sub


sub pDMDLabelFadePulse(LabName,mLen,mColor)   'alpha is 255 max, 0=clear. alpha start/end and pulsespeed of change
    PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':1,'at':6,'astart':70,'aend':255,'len':" & (mLen) & ",'pspeed': 40,'fc':" & mColor & "}" 
end Sub

Sub pDMDLabelFlash(LabName,byVal timeSec, mColor)   'timeSec in ms
    if timeSec<20 Then timeSec=timeSec*1000
    PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':1,'at':1,'fq':150,'len':" & (timeSec) & ",'fc':" & mColor & "}"   
end sub



sub pDMDScreenFadeOut(LabName,mLen)   'alpha is 255 max, 0=clear.  
     PuPlayer.LabelSet pDMD,labName,"`u`",0,"{'mt':1,'at':7,'astart':255,'aend':0,'len':" & (mLen) & " }"    
end Sub

sub pDMDScreenFadeIn(LabName,mLen)    'alpha is 255 max, 0=clear. 
     PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':1,'at':7,'astart':0,'aend':255,'len':" & (mLen) & " }"    
end Sub



Sub pDMDScrollBig(LabName,msgText,byVal timeSec,mColor) 'timeSec in MS
if timeSec<20 Then timeSec=timeSec*1000
PuPlayer.LabelSet pDMD,LabName,msgText,0,"{'mt':1,'at':2,'xps':1,'xpe':-1,'len':" & (timeSec) & ",'mlen':" & (timeSec*1) & ",'tt':0,'fc':" & mColor & "}"
end sub

Sub pDMDScrollBigV(LabName,msgText,byVal timeSec,mColor) 'timeSec in MS
if timeSec<20 Then timeSec=timeSec*1000
PuPlayer.LabelSet pDMD,LabName,msgText,0,"{'mt':1,'at':2,'yps':1,'ype':-1,'len':" & (timeSec) & ",'mlen':" & (timeSec*0.8) & ",'tt':0,'fc':" & mColor & "}"
end sub


Sub pDMDZoomBig(LabName,msgText,byVal timeSec,mColor,isVis,byVal zStart,byVal zEnd)  'timeSec in MS  zstart/end is % of screen height  notice aa antialias is 0 for big font zooms for performance.  'ns is size by %label height.
if timeSec<20 Then timeSec=timeSec*1000
PuPlayer.LabelSet pDMD,LabName,msgText,isVis,"{'mt':1,'at':3,'hstart':" & (zStart) & ",'hend':" & (zEnd) & ",'len':" & (timeSec) & ",'mlen':" & (timeSec*0.4) & ",'tt':" & 0 & ",'fc':" & mColor & ", 'ns':1, 'aa':0}"
end sub




Sub AudioDuckPuP(MasterPuPID,VolLevel)  
'will temporary volume duck all pups (not masterid) till masterid currently playing video ends.  will auto-return all pups to normal.
'VolLevel is number,  0 to mute 99 for 99%  
PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "& MasterPuPID& ", ""FN"": 42, ""DV"": "&VolLevel&" }"             
end Sub

Sub AudioDuckPuPAll(MasterPuPID,VolLevel)  
'will temporary volume duck all pups (not masterid) till masterid currently playing video ends.  will auto-return all pups to normal.
'VolLevel is number,  0 to mute 99 for 99%  
PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "& MasterPuPID& ", ""FN"": 42, ""DV"": "&VolLevel&" , ""ALL"":1 }"             
end Sub




Sub pSetAspectRatio(PuPID, arWidth, arHeight)
     PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "&PuPID& ", ""FN"": 50, ""WIDTH"": "&arWidth&", ""HEIGHT"": "&arHeight&" }"   
end Sub  

Sub pDisableLoopRefresh(PuPID)
     PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "&PuPID& ", ""FN"": 2, ""FF"":0, ""FO"":0 }"   
end Sub  

'set safeloop mode on current playing media.  Good for background videos that refresh often?  { "mt":301, "SN": XX, "FN":41 }
Sub pSafeLoopModeCurrentVideo(PuPID)
     PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "&PuPID& ", ""FN"": 41 }"   
end Sub  

Sub pSetLowQualityPc  'sets fulldmd to run in lower quality mode (slowpc mode)  AAlevel for text is removed and other performance/quality items.  default is always run quality, 
     PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 5, ""FN"":45, ""SP"":1 }"    'slow pc mode
end Sub 

Sub pDMDSetTextQuality(AALevel)  '0 to 4 aa.  4 is sloooooower.  default 1,  perhaps use 2-3 if small desktop view.  only affect text quality.  can set per label too with 'qual' settings.
     PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 5, ""FN"":52, ""SC"": "& AALevel &" }"    'slow pc mode
end Sub   

Sub pDMDLabelDispose(labName)   'not needed unless you want to want to free a heavy resource label from cache/memory.  or temp lables that you created.  performance reasons.
      PuPlayer.LabelSet pDMD,labName,"`u`",1,"{'mt':2,'dispose': 1 }"   
end Sub

Sub pDMDAlwaysPAD  'will pad all text with a space before and after to help with possible text clipping.
     PuPlayer.SendMSG "{ ""mt"":301, ""SN"": 5, ""FN"":46, ""PA"":1 }"    'slow pc mode
end Sub   


Sub pDMDSetHUD(isVis)   'show hide just the pBackGround object (HUD overlay).      
    pDMDLabelVisible "pBackGround",isVis
end Sub  




Sub pDMDSetPage(pagenum)    
    PuPlayer.LabelShowPage pDMD,pagenum,0,""   'set page to blank 0 page if want off
    PDMDCurPage=pagenum
end Sub

Sub pDMDSplashPage(pagenum, cTime)    'cTime is seconds.  3 5,  it will auto return to current page after ctime
    PuPlayer.LabelShowPage pDMD,pagenum,cTime,""   'set page to blank 0 page if want off
    PDMDCurPage=pagenum
end Sub



Sub PDMDSplashPagePlaying(pagenum)  'will hide HUD and show labepage while current media is playing. and then autoreturn.
    PuPlayer.LabelShowPage pDMD,pagenum,500,"hidehudplay"
end Sub    

Sub PDMDSplashPagePlayingHUD(pagenum)  'will show labelpage and auto return to def after current video stopped
    PuPlayer.LabelShowPage pDMD,pagenum,500,"returnplay"
end Sub    


Sub pHideOverlayDuringCurrentPlay() 'will hide pup text labels and HUD till current video stops playing.
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "& "5"& ", ""FN"": 34 }"             'hideoverlay text during next videoplay on DMD auto return
end Sub


Sub pSetVideoPosMS(mPOS)  'set position of video/audio in ms,  must be playing already or will be ignored.  { "mt":301, "SN": XX, "FN":51, "SP": 3431} 
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "& "5"& ", ""FN"": 51, ""SP"":"&mPOS&" }"
end Sub

sub pAllVisible(lvis)   '0/1 to show hide pup text overlay and HUD
    PuPlayer.SendMSG "{ ""mt"":301, ""SN"": "& "5"& ",""OT"":"&lvis&", ""FN"": 3 }"             'hideoverlay text force
end Sub


Sub pDMDSetBackFrame(fname)
  PuPlayer.playlistplayex pDMD,"PUPFrames",fname,0,1    
end Sub

Sub pDMDBackLoopStart(fPlayList,fname)
  PuPlayer.playlistplayex pDMD,fPlayList,fname,0,1
  PuPlayer.SetBackGround pDMD,1
end Sub

Sub pDMDBackLoopStop
  PuPlayer.SetBackGround pDMD,0
  PuPlayer.playstop pDMD
end Sub

'jukebox mode will auto advance to next media in playlist and you can use next/prior sub to manuall advance
'you should really have a specific pupid# display like musictrack that is only used for the playlist
'sub PUPDisplayAsJukebox(pupid) needs to be called/set prior to sending your first media to that pupdisplay.
'pupid=pupdiplay# like pMusic

Sub PUPDisplayAsJukebox(pupid)
PuPlayer.SendMSG("{'mt':301, 'SN': " & pupid & ", 'FN':30, 'PM':1 }")
End Sub

Sub PuPlayListPrior(pupid)
 PuPlayer.SendMSG("{'mt':301, 'SN': " & pupid & ", 'FN':31, 'PM':1 }")
End Sub

Sub PuPlayListNext(pupid)
 PuPlayer.SendMSG("{'mt':301, 'SN': " & pupid & ", 'FN':31, 'PM':2 }")
End Sub

Sub pDMDPause()
 PuPlayer.playpause pDMD
end Sub

Sub pDMDResume()
 PuPlayer.playresume pDMD
end Sub

Sub pDMDStop()
 PuPlayer.playstop pDMD
end Sub

Sub pDMDVolumeDef(cVol)  'sets the default volume of player, doesnt affect current playing media
 PuPlayer.setVolume pdmd,cVol
end Sub

Sub pDMDVolumeCurrent(cVol)  'sets the volume of current media (like to duck audio), doesnt affect default volume for next media.
 PuPlayer.setVolumeCurrent pdmd,cVol
end Sub

Sub pDMDSetLoop(isLoop)     'it will loop the currently playing file 0=cancel looping 1=loop
 PuPlayer.setLoop pDMD,isLoop
end Sub

Sub pDMDBackground(isBack)  'will set the currently playing file as background video and continue to loop and return to it automatically 0=turn off as background.
 PuPlayer.setBackground pDMD,isBack
end Sub


Sub PuPEvent(EventNum)
if hasPUP=false then Exit Sub
PuPlayer.B2SData "D"&EventNum,1  'send event to puppack driver  
End Sub


Sub pupCreateLabel(lName, lValue, lFont, lSize, lColor, xpos, ypos,pagenum, lvis)
PuPlayer.LabelNew pDMD,lName ,lFont,lSize,lColor,0,1,1,1,1,pagenum,lvis
if pUseFramePos=1 Then pTranslatePos xpos,ypos
if pUseFramePos=1 Then pTranslateY lSize
PuPlayer.LabelSet pDMD,lName,lValue,lvis,"{'mt':2,'xpos':"& xpos & ",'ypos':"&ypos&",'fonth':"&lsize&",'v2':1 }"
end Sub

Sub pupCreateLabelImage(lName, lFilename,xpos, ypos, Iwidth, Iheight, pagenum, lvis)
PuPlayer.LabelNew pDMD,lName ,"",50,RGB(100,100,100),0,1,1,0,1,pagenum,lvis
if pUseFramePos=1 Then pTranslatePos xpos,ypos
if pUseFramePos=1 Then pTranslatePos Iwidth,iHeight
PuPlayer.LabelSet pDMD,lName,lFilename,lvis,"{'mt':2,'width':"&IWidth&",'height':"&Iheight&",'xpos':"&xpos&",'ypos':"&ypos&",'v2':1 }"
end Sub

Sub pupCreateLabelImageBG(lName, lFilename,xpos, ypos, Iwidth, Iheight, pagenum, lvis)
PuPlayer.LabelNew pBackglass,lName ,"",50,RGB(100,100,100),0,1,1,0,1,pagenum,lvis
if pUseFramePos=1 Then pTranslatePos xpos,ypos
if pUseFramePos=1 Then pTranslatePos Iwidth,iHeight
PuPlayer.LabelSet pBackglass,lName,lFilename,lvis,"{'mt':2,'width':"&IWidth&",'height':"&Iheight&",'xpos':"&xpos&",'ypos':"&ypos&",'v2':1 }"
end Sub



Sub pDMDSplashBig(msgText,timeSec, mColor)   'note timesec is seconds( 2, 3..etc) , if timesec>1000 then its ms. (2300, 3200)
PuPlayer.LabelShowPage pDMD,2,timeSec,""
PuPlayer.LabelSet pDMD,"Splash",msgText,0,"{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) & ",'fc':" & mColor & "}"   
end sub

Sub pDMDSplashScore(msgText,timeSec, mColor)   'note timesec is seconds( 2, 3..etc) , if timesec>1000 then its ms. (2300, 3200)
PuPlayer.LabelSet pDMD,"ScoreSplash",msgText,0,"{'mt':1,'at':1,'fq':150,'len':" & (timeSec*1000) & ",'fc':" & mColor & "}"   
end sub



Sub SplashPageHUDSample
 PuPlayer.playlistplayex pDMD,"RandomScoring","Fire Missile 1.mp4",0,1
 PDMDSplashPagePlayingHUD 4
end Sub

Sub SplashPageSample
 PuPlayer.playlistplayex pDMD,"RandomScoring","Fire Missile 1.mp4",0,1
 PDMDSplashPagePlaying 4
end Sub




'************** Nailbuster TriggerScript Code v1.20
' create a timer in table named exactly pTriggerScript
' set interval on timer to 100ms (not enabled on startup)
' 
'  currently support up to 10 concurrent timers
'
'
'   usage:    TriggerScript <ms>, "vbcode to execute"
'
'   simpletest:  TriggerScript 3500,"MsgBox 1234"   'will show a dialog 1234
'



Const TriggerScriptSize=10
Dim pReset(10)                 ' TriggerScriptSize
Dim pStatement(10)             ' TriggerScriptSize - holds future scripts
Dim FX




Sub TriggersStopAll()
for fx=0 to TriggerScriptSize
    pReset(FX)=0
    pStatement(FX)=""
next
pTriggerScript.Enabled=False  'YOU MUST HAVE A TIMER NAMED pTriggerScript interval 100 not active on startup.
end Sub

TriggersStopAll



DIM pTriggerCounter:pTriggerCounter=pTriggerScript.interval    'YOU MUST HAVE A TIMER NAMED pTriggerScript interval 100 not active on startup.

Sub pTriggerScript_Timer()
    dim bMoreToRun:bMoreToRun=False
    for fx=0 to TriggerScriptSize  
        if pReset(fx)>0 Then    
            pReset(fx)=pReset(fx)-pTriggerCounter 
            if pReset(fx)<=0 Then
                pReset(fx)=0
                execute(pStatement(fx))
            end if
            bMoreToRun=True            
        End if
    next
    if bMoreToRun = False then pTriggerScript.Enabled=False    ' Disable when we dont need it 
End Sub


Sub TriggerScript(pTimeMS, pScript) ' This is used to Trigger script after the pTriggerScript Timer
    for fx=0 to TriggerScriptSize  
        if pReset(fx)=0 Then
            pReset(fx)=pTimeMS
            pStatement(fx)=pScript
            if pTriggerScript.enabled=false Then pTriggerScript.enabled=true
            Exit Sub
        End If 
    next
end Sub

'*************************************************************************************************
'                END PinUp Player USER Config
'*************************************************************************************************
Const nPupVideoVolume = 100

Function rndNumNot(min, max, notNum)
    Randomize
    Dim iNum

Do While 1
    iNum = Int((max-min+1)*Rnd+min)
    If iNum <> notNum Then
        Exit Do
    End If 
loop

rndNumNot = iNum ' return values
End Function


Sub PlayWiz1Video
	PuPlayer.playevent pDMDVideo,"Wizard","MiniWizard1.mp4",nPupVideoVolume,65,1,0,""
End Sub

Sub PlayWiz2Video
	PuPlayer.playevent pDMDVideo,"Wizard","MiniWizard2.mp4",nPupVideoVolume,65,1,0,""
End Sub

Sub PlayWiz3Video
	PuPlayer.playevent pDMDVideo,"Wizard","MiniWizard3.mp4",nPupVideoVolume,65,1,0,""
End Sub

Sub PlayWizFinalVideo
	PuPlayer.playevent pDMDVideo,"Wizard","FinalWizard.mp4",nPupVideoVolume,65,1,0,""
End Sub

Dim LastBSVideo : LastBSVideo = 0
Sub PlayBallSaveVideo
	Dim i
	'i = rndNumNot(1,1,LastBSVideo)
	LastBSVideo = LastBSVideo + 1
	i = RndNbr(2)
	debug.print "BS:" &i
    SupressModeMessages 2500
	PuPlayer.playevent pDMDVideo,"BallSave","BallSaved_"&i&".mp4",nPupVideoVolume,68,3,0,""
	'LastBSVideo = i
End Sub

Dim LastBLVideo : LastBLVideo = 0
Sub PlayBallLostVideo
	Dim i
'	i = RndNbr(5)
'	debug.print "BL:" &i
    SupressModeMessages 2500
	PuPlayer.playevent pDMDVideo,"BallLost","BallLost.mp4",nPupVideoVolume,90,0,0,""

End Sub

Dim LastACVideo : LastACVideo = 0
Sub PlayAlbumCollected
	Dim i
	i = rndNumNot(1,5,LastACVideo)
'	debug.print "AC:" &i
    SupressModeMessages 2500
	PuPlayer.playevent pDMDVideo,"Mode","albumcollected"&i&".mp4",nPupVideoVolume,77,3,0,""
	LastACVideo = i
End Sub

Sub PlayMultipliers
    SupressModeMessages 2500
	PuPlayer.playevent pDMDVideo,"Multipliers","Bonus"&BonusMultiplier(CurrentPlayer)&"x.mp4",nPupVideoVolume,70,3,0,""
End Sub

Sub PlayPlayfieldMultipliers
    SupressModeMessages 2500
	PuPlayer.playevent pDMDVideo,"Multipliers","Playfield"&PlayfieldMultiplier(CurrentPlayer)&"x.mp4",nPupVideoVolume,70,3,0,""
End Sub

'***************************************************************
' 	ZQUE: VPIN WORKSHOP ADVANCED QUEUING SYSTEM - 1.1.1
'***************************************************************
' WHAT IS IT?
' The VPin Workshop Advanced Queuing System allows table authors
' to put sub routine calls in a queue without creating a bunch
' of timers. There are many use cases for this: queuing sequences
' for light shows and DMD scenes, delaying solenoids until the
' DMD is finished playing all its sequences (such as holding a
' ball in a scoop), managing what actions take priority over
' others (e.g. an extra ball sequence is probably more important
' than a small jackpot), and many more.
'
' This system uses Scripting.Dictionary, a single timer, and the
' GameTime global to keep track of everything in the queue.
' This allows for better stability and a virtually unlimited
' number of items in the queue. It also allows for greater
' versatility, like pre-delays, queue delays, priorities, and
' even modifying items in the queue.
'
' The VPin Workshop Queuing System can replace vpmTimer as a
' proper queue system (each item depends on the previous)
' whereas vpmTimer is a collection of virtual timers that run
' in parallel. It also adds on other advanced functionality.
' However, this queue system does not have ROM support out of
' the box like vpmTimer does.
'
' I recommend reading all the comments before you implement the
' queuing system into your table.
'
' WHAT YOU NEED to use the queuing system:
' 1) Put this VBS file in your scripts folder.
' 2) Include this file via Scripting.FileSystemObject, and
'	ExecuteGlobal it.
' 3) Make one or more queues by constructing the vpwQueueManager:
'	Dim queue : Set queue = New vpwQueueManager
' 4) Create (or use) a timer that is always enabled and
'	preferably has an interval of 1 millisecond. Use a
'	higher number for less time precision but less resource
'	use. You only need one timer even if you
'	have multiple queues.
' 5) For each queue you created, call its Tick routine in
'	the timer's *_timer() routine:
'	queue.Tick
' 6) You're done! Refer to the routines in vpwQueueManager to
'	learn how to use the queuing system.
'***************************************************************

'===========================================
' vpwQueueManager
' This class manages a queue of
' vpwQueueItems and executes them.
'===========================================
Class vpwQueueManager
	Public qItems	   ' A dictionary of vpwQueueItems in the queue (do NOT use native Scripting.Dictionary.Add/Remove; use the vpwQueueManager's Add/Remove methods instead!)
	Public preQItems	' A dictionary of vpwQueueItems pending to be added to qItems
	Public debugOn	  ' Null = no debug. String = activate debug by using this unique label for the queue. REQUIRES baldgeek's error logs.
	
	'----------------------------------------------------------
	' vpwQueueManager.qCurrentItem
	' This contains a string of the key currently active / at
	' the top of the queue. An empty string means no items are
	' active right now.
	' This is an important property; it should be monitored
	' in another timer or routine whenever you Add a queue item
	' with a -1 (indefinite) preDelay or postDelay. Then, for
	' preDelay, ExecuteCurrentItem should be called to run the
	' queue item. And for postDelay, DoNextItem should be
	' called to move to the next item in the queue.
	'
	' For example, let's say you add a queue item with the
	' key "kickTheBall" and an indefinite preDelay. You want
	' to wait until another timer fires before this queue item
	' executes and kicks the ball out of a scoop. In the other
	' timer, you will monitor qCurrentItem. Once it equals
	' "kickTheBall", call ExecuteCurrentItem, which will run
	' the queue item and presumably kick out the ball.
	'
	' WARNING!: If you do not properly execute one of these
	' callback routines on an indefinite delayed item, then
	' the queue will effectively freeze / stop until you do.
	'---------------------------------------------------------
	Public qCurrentItem
	
	Public preDelayTime	 ' The GameTime the preDelay for the qCurrentItem was started
	Public postDelayTime	' The GameTime the postDelay for the qCurrentItem was started
	
	Private onQueueEmpty	' A string or object to be called every time the queue empties (use the QueueEmpty property to get/set this)
	Private queueWasEmpty   ' Boolean to determine if the queue was already empty when firing DoNextItem
	
	Private Sub Class_Initialize
		Set qItems = CreateObject("Scripting.Dictionary")
		Set preQItems = CreateObject("Scripting.Dictionary")
		qCurrentItem = ""
		onQueueEmpty = ""
		queueWasEmpty = True
		debugOn = Null
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.Tick
	' This is where all the magic happens! Call this method in
	' your timer's _timer routine to check the queue and
	' execute the necessary methods. We do not iterate over
	' every item in the queue here, which allows for superior
	' performance even if you have hundreds of items in the
	' queue.
	'----------------------------------------------------------
	Public Sub Tick()
		Dim item
		If qItems.Count > 0 Then ' Don't waste precious resources if we have nothing in the queue
			' If no items are active, or the currently active item no longer exists, move to the next item in the queue.
			' (This is also a failsafe to ensure the queue continues to work even if an item gets manually deleted from the dictionary).
			If qCurrentItem = "" Or Not qItems.Exists(qCurrentItem) Then
				DoNextItem
			Else ' We are good; do stuff as normal
				Set item = qItems.item(qCurrentItem)
				
				If item.Executed Then
					' If the current item was executed and the post delay passed, go to the next item in the queue
					If item.postDelay >= 0 And GameTime >= (postDelayTime + item.postDelay) Then
						DebugLog qCurrentItem & " - postDelay of " & item.postDelay & " passed."
						DoNextItem
					End If
				Else
					' If the current item expires before it can be executed, go to the next item in the queue
					If item.timeToLive > 0 And GameTime >= (item.queuedOn + item.timeToLive) Then
						DebugLog qCurrentItem & " - expired (Time To live). Moving To the Next queue item."
						DoNextItem
					End If
					
					' If the current item was not executed yet and the pre delay passed, then execute it
					If item.preDelay >= 0 And GameTime >= (preDelayTime + item.preDelay) Then
						DebugLog qCurrentItem & " - preDelay of " & item.preDelay & " passed. Executing callback."
						item.Execute
						preDelayTime = 0
						postDelayTime = GameTime
					End If
				End If
			End If
		End If
		
		' Loop through each item in the pre-queue to find any that is ready to be added
		If preQItems.Count > 0 Then
			Dim k, key
			k = preQItems.Keys
			For Each key In k
				Set item = preQItems.Item(key)
				
				' If a queue item was pre-queued and is ready to be considered as actually in the queue, add it
				If GameTime >= (item.queuedOn + item.preQueueDelay) Then
					DebugLog key & " (preQueue) - preQueueDelay of " & item.preQueueDelay & " passed. Item added To the main queue."
					preQItems.Remove key
					item.preQueueDelay = 0
					item.queuedOn = GameTime
					qItems.Add key, item
					queueWasEmpty = False
				End If
			Next
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.DoNextItem
	' Goes to the next item in the queue and deletes the
	' currently active one.
	'----------------------------------------------------------
	Public Sub DoNextItem()
		If Not qCurrentItem = "" Then
			If qItems.Exists(qCurrentItem) Then qItems.Remove qCurrentItem ' Remove the current item from the queue if it still exists
			qCurrentItem = ""
		End If
		
		If qItems.Count > 0 Then
			Dim k, key
			Dim nextItem
			Dim nextItemPriority
			Dim nextItemTimeToLive
			Dim nextItemQueuedOn
			Dim item
			nextItemPriority = 0
			nextItem = ""
			
			' Find which item needs to run next based on priority first, queue order second (ignore items with an active preQueueDelay)
			k = qItems.Keys
			For Each key In k
				Set item = qItems.Item(key)
				
				If item.preQueueDelay <= 0 And item.priority > nextItemPriority Then
					nextItem = key
					nextItemPriority = item.priority
					nextItemTimeToLive = item.timeToLive
					nextItemQueuedOn = item.queuedOn
				End If
			Next
			
			If qItems.Exists(nextItem) Then
				DebugLog "DoNextItem - checking " & nextItem & " (priority " & nextItemPriority & ")"
				
				' Make sure the item is not expired. If it is, remove it and re-call doNextItem
				If nextItemTimeToLive > 0 And GameTime >= (nextItemQueuedOn + nextItemTimeToLive) Then
					DebugLog "DoNextItem - " & nextItem & " expired (Time To live). Removing And going To the Next item."
					qItems.Remove nextItem
					DoNextItem
					Exit Sub
				End If
				
				' Set item as current / active, and execute if it has no pre-delay (otherwise Tick will take care of pre-delay)
				qCurrentItem = nextItem
				Set item = qItems.Item(nextItem)
				If item.preDelay = 0 Then
					DebugLog "DoNextItem - " & nextItem & " Now active. It has no preDelay, so executing callback immediately."
					item.Execute
					preDelayTime = 0
					postDelayTime = GameTime
				Else
					DebugLog "DoNextItem - " & nextItem & " Now active. Waiting For a preDelay of " & item.preDelay & " before executing."
					preDelayTime = GameTime
					postDelayTime = 0
				End If
			End If
		ElseIf queueWasEmpty = False Then
			DebugLog "DoNextItem - Queue Is Now Empty; executing queueEmpty callback."
			CallQueueEmpty() ' Call QueueEmpty if this was the last item in the queue
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.ExecuteCurrentItem
	' Helper routine that can be used when the current item is
	' on an indefinite preDelay. Call this when you are ready
	' for that item to execute.
	'----------------------------------------------------------
	Public Sub ExecuteCurrentItem()
		If Not qCurrentItem = "" And qItems.Exists(qCurrentItem) Then
			DebugLog "ExecuteCurrentItem - Executing the callback For " & qCurrentItem & "."
			Dim item
			Set item = qItems.Item(qCurrentItem)
			item.Execute
			preDelayTime = 0
			postDelayTime = GameTime
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.Add
	' REQUIRES Class vpwQueueItem
	'
	' Add an item to the queue.
	'
	' PARAMETERS:
	'
	' key (string) - Unique name for this queue item
	' (warning: Specifying a key that already exists will
	'  overwrite the item in the queue)
	'
	' qCallback (object|string) - An object to be called,
	' or string to be executed globally, when this queue item
	' runs. I highly recommend making sub routines for groups
	' of things that should be executed by the queue so that
	' your qCallback string does not get long, and you can
	' easily organize your callbacks. Also, use double
	' double-quotes when the call itself has quotes in it
	' (VBScript escaping).
	' Example: "playsound ""Plunger"""
	'
	' priority (number) - Items in the queue will be executed
	' in order from highest priority to lowest. Items with the
	' same priority will be executed in order according to
	' when they were added to the queue. Use any number
	' greater than 0. My recommendation is to make a plan for
	' your table on how you will prioritize various types of
	' queue items and what priority number each type should
	' have. Also, you should reserve priority 1 (lowest) to
	' items which should wait until everything else in the
	' queue is done (such as ejecting a ball from a scoop).
	'
	' preQueueDelay (number) - The number of
	' milliseconds before the queue actually considers this
	' item as "in the queue" (pretend you started a timer to
	' add this item into the queue after this delay; this
	' logically works in a similar way; the only difference is
	' timeToLive is still considered even when an item is
	' pre-queued.) Set to 0 to add to the queue immediately.
	' NOTE: this should be less than timeToLive.
	'
	' preDelay (number) - The number of milliseconds before
	' the qCallback executes once this item is active (top)
	' in the queue. Set this to 0 to immediately execute the
	' qCallback when this item becomes active.
	' Set this to -1 to have an indefinite delay until
	' vpwQueueManager.ExecuteCurrentItem is called (see the
	' comment for qCurrentItem for more information).
	' NOTE: this should be less than timeToLive. And, if
	' timeToLive runs out before preDelay runs out, the item
	' will be removed and will not execute.
	'
	' postDelay (number) - After the qCallback executes, the
	' number of milliseconds before moving on to the next item
	' in the queue. Set this to -1 to have an indefinite delay
	' until vpwQueueManager.DoNextItem is called (see the
	' comment for qCurrentItem for more information).
	'
	' timeToLive (number) - After this item is added to the
	' queue, the number of milliseconds before this queue item
	' expires / is removed if the qCallback is not executed by
	' then. Set to 0 to never expire. NOTE: If not 0, this should
	' be greater than preDelay + preQueueDelay or the item will
	' expire before the qCallback is executed.
	' Example use case: Maybe a player scored a jackpot, but
	' it would be awkward / irrelevant to play that jackpot
	' sequence if it hasn't played after a few seconds (e.g.
	' other items in the queue took priority).
	'
	' executeNow (boolean) - Specify true if this item
	' should interrupt the queue and run immediately. This
	' will only happen, however, if the currently active item
	' has a priority less than or equal to the item you are
	' adding. Note this does not bypass preQueueDelay nor
	' preDelay if set.
	' Example: If a player scores an extra ball, you might
	' want that to interrupt everything else going on as it
	' is an important milestone.
	'----------------------------------------------------------
	Public Sub Add(key, qCallback, priority, preQueueDelay, preDelay, postDelay, timeToLive, executeNow)
		DebugLog "Added " & key
		
		' Remove duplicate if it exists
		If preQueueDelay <= 0 And qItems.Exists(key) Then
			DebugLog key & " (Add) - Already exists In the queue. Replacing With the new one."
			qItems.Remove key
			If qCurrentItem = key Then qCurrentItem = "" ' Prevent infinite loops if this queue item has no preDelay and re-adds itself via the callback
		End If
		If preQueueDelay > 0 And preQItems.Exists(key) Then
			preQItems.Remove key
		End If
		
		' Construct the item class
		Dim newClass
		Set newClass = New vpwQueueItem
		With newClass
			.Callback = qCallback
			.priority = priority
			.preQueueDelay = preQueueDelay
			.preDelay = preDelay
			.postDelay = postDelay
			.timeToLive = timeToLive
		End With
		
		' Determine execution stuff if this item does not have a pre-queue delay
		If preQueueDelay <= 0 Then
			If executeNow = True Then
				' Make sure this item does not immediately execute if the current item has a higher priority
				If Not qCurrentItem = "" And qItems.Exists(qCurrentItem) Then
					Dim item
					Set item = qItems.Item(qCurrentItem)
					If item.priority <= priority Then
						DebugLog key & " (Add) - Execute Now was Set To True And this item's priority (" & priority & ") Is >= the active item's priority (" & item.priority & " from " & qCurrentItem & "). Making it the current active queue item."
						qItems.Remove qCurrentItem ' TODO: Do we really want to remove an item if it has not executed yet (preDelay)?
						qCurrentItem = key
						If preDelay = 0 Then
							DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
							newClass.Execute
							preDelayTime = 0
							postDelayTime = GameTime
						Else
							DebugLog key & " (Add) - Waiting For a pre-delay of " & preDelay & " before executing the callback."
							preDelayTime = GameTime
							postDelayTime = 0
						End If
					Else
						DebugLog key & " (Add) - Execute Now was Set To True, but this item's priority (" & priority & ") Is Not >= the active item's priority (" & item.priority & " from " & qCurrentItem & "). This item will Not be executed Now And will be added To the queue normally."
					End If
				Else
					DebugLog key & " (Add) - Execute Now was Set To True And no item was active In the queue. Making it the current active queue item."
					qCurrentItem = key
					If preDelay = 0 Then
						DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
						newClass.Execute
						preDelayTime = 0
						postDelayTime = GameTime
					Else
						DebugLog key & " (Add) - Waiting For a pre-delay of " & preDelay & " before executing the callback."
						preDelayTime = GameTime
						postDelayTime = 0
					End If
				End If
			ElseIf qCurrentItem = key Then
				DebugLog key & " (Add) - An item With the same key Is currently the active queue item. Making this item the active one Now."
				If preDelay = 0 Then
					DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
					newClass.Execute
					preDelayTime = 0
					postDelayTime = GameTime
				Else
					DebugLog key & " (Add) - Waiting For a pre-delay of " & preDelay & " before executing the callback."
					preDelayTime = GameTime
					postDelayTime = 0
				End If
			Else
				DebugLog key & " (Add) - Execute Now was False. This item was added To the queue."
			End If
			qItems.Add key, newClass
			queueWasEmpty = False
		Else
			DebugLog key & " (Add) - Not actually added To the queue yet. It has a pre-queue delay of " & preQueueDelay
			preQItems.Add key, newClass
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.Remove
	'
	' Removes an item from the queue. It is better to use this
	' than to remove the item from qItems directly as this sub
	' will also call DoNextItem to advance the queue if
	' the item removed was the active item.
	' NOTE: This only removes items from qItems; to remove
	' an item from preQItems, use the standard
	' Scripting.Dictionary Remove method.
	'
	' PARAMETERS:
	'
	' key (string) - Unique name of the queue item to remove.
	'----------------------------------------------------------
	Public Sub Remove(key)
		If qItems.Exists(key) Then
			DebugLog key & " (Remove)"
			qItems.Remove key
			If qCurrentItem = key Or qCurrentItem = "" Then DoNextItem ' Ensure the queue does not get stuck
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.RemoveAll
	'
	' Removes all items from the queue / clears the queue.
	' It is better to call this sub than to remove all items
	' from qItems directly because this sub cleans up the queue
	' to ensure it continues to work properly.
	'
	' PARAMETERS:
	'
	' preQueue (boolean) - Also clear the pre-queue.
	'----------------------------------------------------------
	Public Sub RemoveAll(preQueue)
		DebugLog "Queue was emptied via RemoveAll."
		
		' Loop through each item in the queue and remove it
		Dim k, key
		k = qItems.Keys
		For Each key In k
			qItems.Remove key
		Next
		qCurrentItem = ""
		
		If queueWasEmpty = False Then CallQueueEmpty() ' Queue is now empty, so call our callback if applicable
		
		If preQueue Then
			k = preQItems.Keys
			For Each key In k
				preQItems.Remove key
			Next
		End If
	End Sub
	
	'----------------------------------------------------------
	' Get vpwQueueManager.QueueEmpty
	' Get the current callback for when the queue is empty.
	'----------------------------------------------------------
	Public Property Get QueueEmpty()
		If IsObject(onQueueEmpty) Then
			Set QueueEmpty = onQueueEmpty
		Else
			QueueEmpty = onQueueEmpty
		End If
	End Property
	
	'----------------------------------------------------------
	' Let vpwQueueManager.QueueEmpty
	' Set the callback to call every time the queue empties.
	' This could be useful for setting a sub routine to be
	' called each time the queue empties for doing things such
	' as ejecting balls from scoops. Unlike using the Add
	' method, this callback is immune from getting removed by
	' higher priority items in the queue and will be called
	' every time the queue is emptied, not just once.
	'
	' PARAMETERS:
	'
	' callback (object|string) - The callback to call every
	' time the queue empties.
	'----------------------------------------------------------
	Public Property Let QueueEmpty(callback)
		If IsObject(callback) Then
			Set onQueueEmpty = callback
		ElseIf VarType(callback) = vbString Then
			onQueueEmpty = callback
		End If
	End Property
	
	'----------------------------------------------------------
	' Get vpwQueueManager.CallQueueEmpty
	' Private method that actually calls the QueueEmpty
	' callback.
	'----------------------------------------------------------
	Private Sub CallQueueEmpty()
		If queueWasEmpty = True Then Exit Sub
		
		If IsObject(onQueueEmpty) Then
			Call onQueueEmpty(0)
		ElseIf VarType(onQueueEmpty) = vbString Then
			If onQueueEmpty > "" Then ExecuteGlobal onQueueEmpty
		End If
		
		queueWasEmpty = True
	End Sub
	
	'----------------------------------------------------------
	' DebugLog
	' Log something if debugOn is not null.
	' REQUIRES / uses the WriteToLog sub from Baldgeek's
	' error log library.
	'----------------------------------------------------------
	Private Sub DebugLog(message)
		If Not IsNull(debugOn) Then
			WriteToLog "VPW Queue " & debugOn, message
		End If
	End Sub
End Class

'===========================================
' vpwQueueItem
' Represents a single item for the queue
' system. Do NOT use this class directly.
' Instead, use the vpwQueueManager.Add
' routine.

' You can, however, access an individual
' item in the queue via
' vpwQueueManager.qItems and then modify
' its properties while it is still in the
' queue.
'===========================================
Class vpwQueueItem
	Public priority		 ' The item's set priority
	Public timeToLive	   ' The item's set timeToLive milliseconds requested
	Public preQueueDelay	' The item's pre-queue milliseconds requested
	Public preDelay		 ' The item's pre delay milliseconds requested
	Public postDelay		' The item's post delay milliseconds requested
	Private qCallback	   ' The item's callback object or string (use the Callback property on the class to get/set it)
	
	Public executed		 ' Whether or not this item's qCallback was executed yet
	Public queuedOn		 ' The game time this item was added to the queue
	Public executedOn	   ' The game time this item was executed
	
	Private Sub Class_Initialize
		' Defaults
		priority = 0
		timeToLive = 0
		preQueueDelay = 0
		preDelay = 0
		postDelay = 0
		qCallback = ""
		queuedOn = GameTime
		executedOn = 0
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueItem.Execute
	' Executes the qCallback on this item if it was not yet
	' already executed.
	'----------------------------------------------------------
	Public Sub Execute()
		If executed Then Exit Sub ' Do not allow an item's qCallback to ever Execute more than one time
		
		' Execute qCallback
		If IsObject(qCallback) Then
			Call qCallback(0)
		ElseIf VarType(qCallback) = vbString Then
			If qCallback > "" Then ExecuteGlobal qCallback
		End If
		
		executedOn = GameTime
		executed = True
	End Sub
	
	Public Property Get Callback()
		If IsObject(qCallback) Then
			Set Callback = qCallback
		Else
			Callback = qCallback
		End If
	End Property
	
	Public Property Let Callback(cb)
		If IsObject(cb) Then
			Set qCallback = cb
		ElseIf VarType(cb) = vbString Then
			qCallback = cb
		End If
	End Property
End Class


'*******************************************
' Routine called each time the queue is
' emptied.
'*******************************************
Sub QueueTimer_Timer()
	BallHandlingQueue.Tick
	GeneralPupQueue.Tick
	DMDQueue.Tick
	AudioQueue.Tick
	EOBQueue.Tick
    LightQueue.Tick
End Sub

Sub WipeAllQueues
	GeneralPupQueue.RemoveAll(True)
	AudioQueue.RemoveAll(True)
	BallHandlingQueue.RemoveAll(True)
	DMDQueue.RemoveAll(True)
	EOBQueue.RemoveAll(True)
    LightQueue.RemoveAll(True)
End Sub



'***************************************************************
'***  END VPIN WORKSHOP ADVANCED QUEUING SYSTEM
'***************************************************************
Dim asModeMessagesL1
asModeMessagesL1 = Array("", "HIT THE SPINNERS", "HIT THE BUMPERS", "SHOOT THE RAMPS", "SHOOT THE ORBITS", "SHOOT THE LIGHTS" , "FOLLOW THE LIGHTS", "HIT THE TARGETS", "HIT THE TARGETS"_
, "CHASE THE LIGHTS", "SHOOT THE LOOPS", "CHASE THE LIGHTS", "HIT RAMPS & ORBITS", "HIT THE")

Dim asModeMessagesL2
asModeMessagesL2 = Array("","SPINNERS LEFT","BUMPER HITS LEFT","RAMP HITS LEFT","ORBIT HITS LEFT","LIGHTS REMAIN", "LIGHTS REMAIN", "TARGETS LEFT", "TARGETS LEFT", "HITS REMAIN", "LOOPS REMAIN"_
, "HITS REMAIN", "SHOTS LEFT", "JACKPOTS" )

' Used for number of shots for each mode to compelte
Dim anModeProgress
anModeProgress = Array("","100","25","6","6","9","9", "20", "6", "8", "6", "8", "6")

Dim anModeProgress2
anModeProgress2 = Array("","","","","","")

Sub UpdateSpinProgress
	PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(1)- SpinCount ,1,"{'mt':2,'xpos':52}"
End Sub

' Colors defined higher up in script just a reminder, can define any you want user colorpicker app to get the number

	' cWhite = 	16777215
	' cRed = 	397512
	' cGold = 	1604786
	' cGold2 = 46079
	' cGreen = 32768
	' cGrey = 	8421504
	' cYellow = 65535
	' cOrange = 33023
	' cPurple = 16711808
	' cBlue = 16711680
	' cLightBlue = 16744448
	' cBoltYellow = 2148582
	' cLightGreen = 9747818
	' cBlack = 0
	' cPink = 12615935
	' cSilver = 8421504


Sub UpdateModeMessages_Timer
'	PuPlayer.LabelSet pDMD,"TESTQTitle","Active Mode",1,"{'mt':2,'fonth':6,'xalign':0,'yalign':0,'ypos':12,'xpos':"&(0 + nOffsetX) &"}"
'	PuPlayer.LabelSet pDMD,"TESTQ",Battle(CurrentPlayer,0),1,"{'mt':2,'fonth':20,'xalign':0,'yalign':0,'ypos':15,'xpos':"&(3 + nOffsetX) &"3}"

	if bSupressModeMessages Then 
		PuPlayer.LabelSet pDMD,"Event3A","",0,""
		PuPlayer.LabelSet pDMD,"Event3B","",0,""
		PuPlayer.LabelSet pDMD,"Event3Ca","",0,""
		PuPlayer.LabelSet pDMD,"Event3C","",0,""
		PuPlayer.LabelSet pDMD,"BlizzTimerValue","",0,""
		Exit Sub
	End If

	if Battle(CurrentPlayer,0) <> 0 Then pDMDlabelHide "BlizzTimerImage"

	' Handle Pup message updates for Blizzard Modes
	if bBlizzardPrepMode AND NOT bSupressModeMessages Then UpdateBlizzPrepMessages : Exit Sub
	'if bBlizzardMode AND NOT bSupressModeMessages Then UpdateBlizzMessages : Exit Sub


	Select Case Battle(CurrentPlayer,0)

		Case 0
			PuPlayer.LabelSet pDMD,"Event3A","",0,""
			PuPlayer.LabelSet pDMD,"Event3B","",0,""
			PuPlayer.LabelSet pDMD,"Event3Ca","",0,""
			PuPlayer.LabelSet pDMD,"Event3C","",0,""
		Case 1
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(1),1,"{'mt':2,'color': " & cYellow &" }"
''			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(1)- SpinCount ,1,""
			PuPlayer.LabelSet pDMD,"Event3C"," "& asModeMessagesL2(1),1,"{'mt':2,'color': " & cYellow &" }"
		Case 2
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(2),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(2) - SuperBumperHits ,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(2),1,"{'mt':2,'color': " & cYellow &" }"
		Case 3
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(3),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(3) - RampHits3 ,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(3),1,"{'mt':2,'color': " & cYellow &" }"
		Case 4
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(4),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(4) - OrbitHits,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(4),1,"{'mt':2,'color': " & cYellow &" }"
		Case 5
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(5),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(5) - Mode5Lights(0),1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(5),1,"{'mt':2,'color': " & cYellow &" }"
		Case 6
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(6),1,"{'mt':2,'color': " & cYellow &" }"
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(6),1,"{'mt':2,'color': " & cYellow &" }"
		Case 7
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(7),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(7) - TargetHits7,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(7),1,"{'mt':2,'color': " & cYellow &" }"
		Case 8
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(8),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(8) - TargetHits8,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(8),1,"{'mt':2,'color': " & cYellow &" }"
		Case 9
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(9),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(9) - LightHits9,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(9),1,"{'mt':2,'color': " & cYellow &" }"
		Case 10
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(10),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(10) - LoopCount,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(10),1,"{'mt':2,'color': " & cYellow &" }"
		Case 11
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(11),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(11) - LightHits11,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(11),1,"{'mt':2,'color': " & cYellow &" }"
		Case 12
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(12),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(12) - RampHits12,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(12),1,"{'mt':2,'color': " & cYellow &" }"
		Case 13
			PuPlayer.LabelSet pDMD,"Event3B",asModeMessagesL1(13),1,"{'mt':2,'color': " & cYellow &" }"
'			PuPlayer.LabelSet pDMD,"Event3Ca",?????,1,""
			PuPlayer.LabelSet pDMD,"Event3C",asModeMessagesL2(13),1,"{'mt':2,'color': " & cYellow &" }"
	End Select


End Sub

Sub UpdateModeProgress(Pulse)
	if bSupressModeMessages or bSupressModeProgress then Exit Sub
	bSupressModeProgress = True

	if Pulse Then 
		Select Case Battle(CurrentPlayer,0)
			Case 1
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(1) - SpinCount ,1,"{'mt':2,'xpos':52}"
			Case 2
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(2) - SuperBumperHits, 1600, cRed
			Case 3
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(3) - RampHits3, 1600, cRed
			Case 4
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(4) - OrbitHits, 1600, cRed
			Case 5
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(5) - Mode5Lights(0), 1600, cRed
			Case 6	
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(6) - LightHits6, 1600, cRed
			Case 7
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(7) - TargetHits7, 1600, cRed
			Case 8
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(8) - TargetHits8, 1600, cRed
			Case 9
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(9) - LightHits9, 1600, cRed
			Case 10
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(10) - LoopCount, 1600, cRed
			Case 11
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(11) - LightHits11, 1600, cRed
			Case 12
				pDMDLabelPulseText2 "Event3Ca", anModeProgress(12) - RampHits12, 1600, cRed
			Case 13
				' nothing
		End Select

	Else
		Select Case Battle(CurrentPlayer,0)
			Case 1
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(1) - SpinCount ,1,"{'mt':2,'xpos':52}"
			Case 2
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(2) - SuperBumperHits ,1,"{'mt':2,'xpos':49.5,'color': " & cRed &"}"
			Case 3
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(3) - RampHits3 ,1,"{'mt':2,'xpos':52,'color': " & cRed &"}"
			Case 4
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(4) - OrbitHits,1,"{'mt':2,'xpos':52,'color': " & cRed &"}"
			Case 5
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(5) - Mode5Lights(0),1,"{'mt':2,'xpos':53,'color': " & cRed &"}"
			Case 6	
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(6) - LightHits6,1,"{'mt':2,'xpos':53,'color': " & cRed &"}"
			Case 7
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(7) - TargetHits7,1,"{'mt':2,'xpos':52.5,'color': " & cRed &"}"
			Case 8
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(8) - TargetHits8,1,"{'mt':2,'xpos':53,'color': " & cRed &"}"
			Case 9
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(9) - LightHits9,1,"{'mt':2,'xpos':55.25,'color': " & cRed &"}"
			Case 10
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(10) - LoopCount,1,"{'mt':2,'xpos':53.5,'color': " & cRed &"}"
			Case 11
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(11) - LightHits11,1,"{'mt':2,'xpos':55.25,'color': " & cRed &"}"
			Case 12
				PuPlayer.LabelSet pDMD,"Event3Ca",anModeProgress(12) - (RampHits12),1,"{'mt':2,'xpos':55,'color': " & cRed &"}"
			Case 13
				' nothing
		End Select
	End If

		DMDQueue.Add "bSupressModeProgress = False-"&Gametime,"bSupressModeProgress = False",65,1600,0,0,0,False
		DMDQueue.Add "bSupressModeProgress = False-"&Gametime,"bSupressModeProgress = False",85,2000,0,0,0,True
End Sub


Sub ClearPupAttractMessages
	PuPlayer.LabelSet pDMD, "Attract2a", "",  0, ""
	PuPlayer.LabelSet pDMD, "Attract2b", "",  0, ""
	PuPlayer.LabelSet pDMD, "JukeBox2a", "",  0, ""
	PuPlayer.LabelSet pDMD, "JukeBox2b", "",  0, ""
End Sub

Dim AttractTimerCount
Sub AttractTimer_Timer()

	Select Case AttractTimerCount
		Case 0
			AttractTimer.Interval = 2000
			PuPlayer.playevent pDMDVideo,"Background","Blank.mp4",100,3,0,0,""
			if bFreeplay Then
					PuPlayer.LabelSet pDMD,"Attract2a"," FREE ",1,"{'mt':2,'color': " & cWhite &" }"
					PuPlayer.LabelSet pDMD,"Attract2b"," PLAY ",1,"{'mt':2,'color': " & cWhite &" }"  
			Else
				If Credits> 0 Then
					PuPlayer.LabelSet pDMD,"Attract2a"," CREDITS ",1,"{'mt':2,'color': " & cWhite &" }"
					PuPlayer.LabelSet pDMD,"Attract2b"," "&credits&" ",1,"{'mt':2,'color': " & cWhite &" }"  
				Else
					PuPlayer.LabelSet pDMD,"Attract2a"," CREDITS 0 ",1,"{'mt':2,'color': " & cWhite &" }"
					PuPlayer.LabelSet pDMD,"Attract2b"," INSERT COIN ",1,"{'mt':2,'color': " & cWhite &" }"  
				End If
			End If

		Case 1
				PuPlayer.LabelSet pDMD,"Attract2a"," "&HighScoreName(0)&" ",1,"{'mt':2,'color': " & cWhite &" }"
				PuPlayer.LabelSet pDMD,"Attract2b",FormatScoreDMD(HighScore(0)),1,"{'mt':2,'color': " & cWhite &" }" 
		Case 2
				PuPlayer.LabelSet pDMD,"Attract2a"," "&HighScoreName(1)&" ",1,""
				PuPlayer.LabelSet pDMD,"Attract2b",FormatScoreDMD(HighScore(1)),1,""   
		Case 3
				PuPlayer.LabelSet pDMD,"Attract2a"," "&HighScoreName(2)&" ",1,""
				PuPlayer.LabelSet pDMD,"Attract2b",FormatScoreDMD(HighScore(2)),1,""   
		Case 4
				PuPlayer.LabelSet pDMD,"Attract2a"," "&HighScoreName(3)&" ",1,""
				PuPlayer.LabelSet pDMD,"Attract2b",FormatScoreDMD(HighScore(3)),1,""   
		Case 5
				PuPlayer.LabelSet pDMD,"Attract2a"," GAME OVER ",1,""
				PuPlayer.LabelSet pDMD,"Attract2b","",1,""   
		Case 6
			PuPlayer.playevent pDMDVideo,"Background","Blank.mp4",100,3,5,0,""
			AttractTimer.Interval = 10500
			PuPlayer.LabelSet pDMD, "Attract2a", "",  0, ""
			PuPlayer.LabelSet pDMD, "Attract2b", "",  0, ""
			PuPlayer.playevent pDMDVideo,"Background","Zandy.mp4",100,3,0,0,""

	End Select

	AttractTimerCount = AttractTimerCount + 1

	if AttractTimerCount > 6 Then AttractTimerCount = 0
End Sub


sub testRTP

newbattle = 3

StartBattle
End Sub
'******************************************************
' 	ZFLD:  FLUPPER DOMES
'******************************************************
' Based on FlupperDoms2.2

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials domebasemat, Flashermaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "dome" and "ronddome" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names
' Copy a set of 4 objects flasherbase, flasherlit, flasherlight and flasherflash from layer 7 to your table
' If you duplicate the four objects for a new flasher dome, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherbloom flashers from layer 10 to your table. you will need to make one per flasher dome that you plan to make
' Select the correct flasherbloom texture for each flasherbloom flasher, per flasher dome
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasher in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitFlasher 1, "green"
'
' Color Options: "blue", "green", "red", "purple", "yellow", "white", and "orange"

' You can use the RotateFlasher call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasher sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasher 1, 180		 'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLevel(1) = 1 : FlasherFlash1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasher below)
' - Color of the objects are set in the script, not on the primitive itself
' - Screws are optional to copy and position manually
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flashermaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlit and flasherflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlightintensity and/or flasherflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

' example script for rom based tables (non modulated):

' SolCallback(25)="FlashRed"
'
' Sub FlashRed(flstate)
'	If Flstate Then
'		ObjTargetLevel(1) = 1
'	Else
'		ObjTargetLevel(1) = 0
'	End If
'   FlasherFlash1_Timer
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
'	ObjTargetLevel(1) = level/255 : FlasherFlash1_Timer
' End Sub

Sub Flash1(Enabled)
	If Enabled Then
		ObjTargetLevel(1) = 1
	Else
		ObjTargetLevel(1) = 0
	End If
	FlasherFlash1_Timer
	Sound_Flash_Relay enabled, Flasherbase1
End Sub

Sub Flash2(Enabled)
	If Enabled Then
		ObjTargetLevel(2) = 1
	Else
		ObjTargetLevel(2) = 0
	End If
	FlasherFlash2_Timer
	Sound_Flash_Relay enabled, Flasherbase2
End Sub

Sub Flash3(Enabled)
	If Enabled Then
		ObjTargetLevel(3) = 1
	Else
		ObjTargetLevel(3) = 0
	End If
	FlasherFlash3_Timer
	Sound_Flash_Relay enabled, Flasherbase3
End Sub

Sub Flash4(Enabled)
	If Enabled Then
		ObjTargetLevel(4) = 1
	Else
		ObjTargetLevel(4) = 0
	End If
	FlasherFlash4_Timer
	Sound_Flash_Relay enabled, Flasherbase1
End Sub


Dim TableRef
Set TableRef = Table1		   ' *** change this, if your table has another name				   ***

Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

Dim TestFlashers, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0				' *** set this to 1 to check position of flasher object			 ***
FlasherLightIntensity = 0.3	 ' *** lower this, if the VPX lights are too bright (i.e. 0.1)	   ***
FlasherFlareIntensity = 0.3	 ' *** lower this, if the flares are too bright (i.e. 0.1)		   ***
FlasherBloomIntensity = 0.2	 ' *** lower this, if the blooms are too bright (i.e. 0.1)		   ***
FlasherOffBrightness = 0.5	  ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "orange"
InitFlasher 2, "orange"
InitFlasher 3, "red"
InitFlasher 4, "green"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
 '  RotateFlasher 1,90
  ' RotateFlasher 2,90
'   RotateFlasher 3,90
'   RotateFlasher 4,90

Sub InitFlasher(nr, col)
	' store all objects in an array for use in FlashFlasher subroutine
	Set objbase(nr) = Eval("Flasherbase" & nr)
	Set objlit(nr) = Eval("Flasherlit" & nr)
	Set objflasher(nr) = Eval("Flasherflash" & nr)
	Set objlight(nr) = Eval("Flasherlight" & nr)
	Set objbloom(nr) = Eval("Flasherbloom" & nr)
	
	' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
	If objbase(nr).RotY = 0 Then
		objbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbase(nr).x) / (objbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
		objflasher(nr).RotZ = objbase(nr).ObjRotZ
		objflasher(nr).height = objbase(nr).z + 40
	End If
	
	' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
	objlight(nr).IntensityScale = 0
	objlit(nr).visible = 0
	objlit(nr).material = "Flashermaterial" & nr
	objlit(nr).RotX = objbase(nr).RotX
	objlit(nr).RotY = objbase(nr).RotY
	objlit(nr).RotZ = objbase(nr).RotZ
	objlit(nr).ObjRotX = objbase(nr).ObjRotX
	objlit(nr).ObjRotY = objbase(nr).ObjRotY
	objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
	objlit(nr).x = objbase(nr).x
	objlit(nr).y = objbase(nr).y
	objlit(nr).z = objbase(nr).z
	objbase(nr).BlendDisableLighting = FlasherOffBrightness
	
	'rothbauerw
	'Adjust the position of the flasher object to align with the flasher base.
	'Comment out these lines if you want to manually adjust the flasher object
	If objbase(nr).roty > 135 Then
		objflasher(nr).y = objbase(nr).y + 50
		objflasher(nr).height = objbase(nr).z + 20
	Else
		objflasher(nr).y = objbase(nr).y + 20
		objflasher(nr).height = objbase(nr).z + 0
	End If
	objflasher(nr).x = objbase(nr).x
	
	'rothbauerw
	'Adjust the position of the light object to align with the flasher base.
	'Comment out these lines if you want to manually adjust the flasher object
	objlight(nr).x = objbase(nr).x
	objlight(nr).y = objbase(nr).y
	objlight(nr).bulbhaloheight = objbase(nr).z - 10
	
	'rothbauerw
	'Assign the appropriate bloom image basked on the location of the flasher base
	'Comment out these lines if you want to manually assign the bloom images
	Dim xthird, ythird
	xthird = tablewidth / 3
	ythird = tableheight / 3
	If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
		objbloom(nr).imageA = "flasherbloomCenter"
		objbloom(nr).imageB = "flasherbloomCenter"
	ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
		objbloom(nr).imageA = "flasherbloomUpperLeft"
		objbloom(nr).imageB = "flasherbloomUpperLeft"
	ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
		objbloom(nr).imageA = "flasherbloomUpperRight"
		objbloom(nr).imageB = "flasherbloomUpperRight"
	ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
		objbloom(nr).imageA = "flasherbloomCenterLeft"
		objbloom(nr).imageB = "flasherbloomCenterLeft"
	ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
		objbloom(nr).imageA = "flasherbloomCenterRight"
		objbloom(nr).imageB = "flasherbloomCenterRight"
	ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
		objbloom(nr).imageA = "flasherbloomLowerLeft"
		objbloom(nr).imageB = "flasherbloomLowerLeft"
	ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
		objbloom(nr).imageA = "flasherbloomLowerRight"
		objbloom(nr).imageB = "flasherbloomLowerRight"
	End If
	
	' set the texture and color of all objects
	Select Case objbase(nr).image
		Case "dome2basewhite"
			objbase(nr).image = "dome2base" & col
			objlit(nr).image = "dome2lit" & col
			
		Case "ronddomebasewhite"
			objbase(nr).image = "ronddomebase" & col
			objlit(nr).image = "ronddomelit" & col
			
		Case "domeearbasewhite"
			objbase(nr).image = "domeearbase" & col
			objlit(nr).image = "domeearlit" & col
	End Select
	If TestFlashers = 0 Then
		objflasher(nr).imageA = "domeflashwhite"
		objflasher(nr).visible = 0
	End If
	Select Case col
		Case "blue"
			objlight(nr).color = RGB(4,120,255)
			objflasher(nr).color = RGB(200,255,255)
			objbloom(nr).color = RGB(4,120,255)
			objlight(nr).intensity = 5000
			
		Case "green"
			objlight(nr).color = RGB(12,255,4)
			objflasher(nr).color = RGB(12,255,4)
			objbloom(nr).color = RGB(12,255,4)
			
		Case "red"
			objlight(nr).color = RGB(255,32,4)
			objflasher(nr).color = RGB(255,32,4)
			objbloom(nr).color = RGB(255,32,4)
			
		Case "purple"
			objlight(nr).color = RGB(230,49,255)
			objflasher(nr).color = RGB(255,64,255)
			objbloom(nr).color = RGB(230,49,255)
			
		Case "yellow"
			objlight(nr).color = RGB(200,173,25)
			objflasher(nr).color = RGB(255,200,50)
			objbloom(nr).color = RGB(200,173,25)
			
		Case "white"
			objlight(nr).color = RGB(255,240,150)
			objflasher(nr).color = RGB(100,86,59)
			objbloom(nr).color = RGB(255,240,150)
			
		Case "orange"
			objlight(nr).color = RGB(255,70,0)
			objflasher(nr).color = RGB(255,70,0)
			objbloom(nr).color = RGB(255,70,0)
	End Select
	objlight(nr).colorfull = objlight(nr).color
	If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
		objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
		ObjFlasher(nr).y = ObjFlasher(nr).y + 10
	End If
End Sub

Sub RotateFlasher(nr, angle)
	angle = ((angle + 360 - objbase(nr).ObjRotZ) Mod 180) / 30
	objbase(nr).showframe(angle)
	objlit(nr).showframe(angle)
End Sub

Sub FlashFlasher(nr)
	If Not objflasher(nr).TimerEnabled Then
		objflasher(nr).TimerEnabled = True
		objflasher(nr).visible = 1
		objbloom(nr).visible = 1
		objlit(nr).visible = 1
	End If
	objflasher(nr).opacity = 1000 * FlasherFlareIntensity * ObjLevel(nr) ^ 2.5
	objbloom(nr).opacity = 100 * FlasherBloomIntensity * ObjLevel(nr) ^ 2.5
	objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr) ^ 3
	objbase(nr).BlendDisableLighting = FlasherOffBrightness + 10 * ObjLevel(nr) ^ 3
	objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr) ^ 2
	UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
	If Round(ObjTargetLevel(nr),1) > Round(ObjLevel(nr),1) Then
		ObjLevel(nr) = ObjLevel(nr) + 0.3
		If ObjLevel(nr) > 1 Then ObjLevel(nr) = 1
	ElseIf Round(ObjTargetLevel(nr),1) < Round(ObjLevel(nr),1) Then
		ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
		If ObjLevel(nr) < 0 Then ObjLevel(nr) = 0
	Else
		ObjLevel(nr) = Round(ObjTargetLevel(nr),1)
		objflasher(nr).TimerEnabled = False
	End If
	'   ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
	If ObjLevel(nr) < 0 Then
		objflasher(nr).TimerEnabled = False
		objflasher(nr).visible = 0
		objbloom(nr).visible = 0
		objlit(nr).visible = 0
	End If
End Sub

Sub FlasherFlash1_Timer()
	FlashFlasher(1)
End Sub
Sub FlasherFlash2_Timer()
	FlashFlasher(2)
End Sub
Sub FlasherFlash3_Timer()
	FlashFlasher(3)
End Sub
Sub FlasherFlash4_Timer()
	FlashFlasher(4)
End Sub

'******************************************************
'******  END FLUPPER DOMES
'******************************************
'*****************
'triggers
'*****************

Sub BackleftFlash
    Flash1 True
    LightQueue.Add "Flash1 False"&gametime,"Flash1 False",50,150,0,0,0,False
End Sub

Sub BackRightFlash
    Flash2 True
    LightQueue.Add "Flash2 False"&gametime,"Flash2 False",50,150,0,0,0,False
End Sub


Sub RightFlash
    Flash4 True
    LightQueue.Add "Flash4 False"&gametime,"Flash4 False",50,150,0,0,0,False
End Sub

Sub LeftFlash
    Flash3 True
    LightQueue.Add "Flash3 False"&gametime,"Flash3 False",50,150,0,0,0,False
End Sub

Dim JPCount
Sub JackpotFlashTimer_Timer()
    JPCount = JPCount + 1
    if JPCount > 5 Then JackpotFlashTimer.Enabled = False
    LeftFlash
    RightFlash
    BackLeftFlash
    BackRightFlash
End Sub

Sub HighScoreHelper(lOne,lTwo,lTime)
	ClearHighScoreTwoLine

	PuPlayer.LabelSet pDMD,"HSLine1",lOne,1,"{'mt':2,'color':" & cPurple & "}"
	PuPlayer.LabelSet pDMD,"HSLine2",lTwo,1,"{'mt':2,'color':" & cPurple & "}"

End Sub

sub ClearHighScoreTwoLine
	PuPlayer.LabelSet pDMD,"HSLine1","",0,""
	PuPlayer.LabelSet pDMD,"HSLine2","",0,""
End Sub

Sub PupGameOver
	bGameReady = False
	PuPlayer.playevent pDMDVideo,"GameOver","GAMEOVER.mp4",100,99,0,0,""
	DMDQueue.Add "bGameReady = True","bGameReady = True",95,6000,0,0,0,True
End Sub

Sub GameOverMesssages
	pDMDSplashTwoLines "GAME", "OVER", 4, cPurple
	DMDQueue.Add "ClearTwoLines","ClearTwoLines",95,4100,0,0,0,True
End Sub

Dim IconStep

Sub IconTimer_Timer
	IconStep = IconStep + 1

	Select Case IconStep
		Case 1
			PuPlayer.LabelSet pDMD, "BonusMP2x","PupOverlays\\2x.png",1,""
		Case 2
			PuPlayer.LabelSet pDMD, "BonusMP3x","PupOverlays\\3x.png",1,""
		Case 3
			PuPlayer.LabelSet pDMD, "BonusMP4x","PupOverlays\\4x.png",1,""
		Case 4
			PuPlayer.LabelSet pDMD, "BonusMP5x","PupOverlays\\5x.png",1,""
'		Case 5
'			PuPlayer.LabelSet pDMD, "BonusMP5x","PupOverlays\\5x.png",0,""
'		Case 6
'			PuPlayer.LabelSet pDMD, "BonusMP4x","PupOverlays\\4x.png",0,""
'		Case 7
'			PuPlayer.LabelSet pDMD, "BonusMP3x","PupOverlays\\3x.png",0,""
		Case 5
'			pDMDLabelPulseImage(LabName,mLen,isVis)
			pDMDLabelPulseImage "BonusMP2x", 2000, 0
			pDMDLabelPulseImage "BonusMP3x", 2000, 0
			pDMDLabelPulseImage "BonusMP4x", 2000, 0
			pDMDLabelPulseImage "BonusMP5x", 2000, 0
		Case 11
			HideAllBonusIcons
			IconTimer.Enabled = False
	End Select

End Sub

Sub Dbg2( myDebugText )
' Uncomment the next line to turn off debugging
Exit Sub

If Not IsObject( objIEDebugWindow ) Then
Set objIEDebugWindow = CreateObject( "InternetExplorer.Application" )
objIEDebugWindow.Navigate "about:blank"
objIEDebugWindow.Visible = True
objIEDebugWindow.ToolBar = False
objIEDebugWindow.Width = 600	
objIEDebugWindow.Height = 900
objIEDebugWindow.Left = 2100
objIEDebugWindow.Top = 100
Do While objIEDebugWindow.Busy
Loop
objIEDebugWindow.Document.Title = "My Debug Window"
objIEDebugWindow.Document.Body.InnerHTML = "<b>Ozzy Debug Window -TimeStamp: " & GameTime& "</b></br>"
End If

objIEDebugWindow.Document.Body.InnerHTML = objIEDebugWindow.Document.Body.InnerHTML & myDebugText & " --TimeStamp:<b> " & GameTime & "</b><br>" & vbCrLf
End Sub

Sub Gate4_Hit()
	if ScorbitActive Then
		HideScorbit
	End If
End Sub

'**************************
'   SCORBIT
'**************************
'==================================================================================================================
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  SCORBIT Interface
' To Use:
' 1) Define a timer tmrScorbit
' 2) Call DoInit at the end of PupInit or in Table Init if you are nto using pup with the appropriate parameters
'     Replace 389 with your TableID from Scorbit 
'     Replace GRWvz-MP37P from your table on OPDB - eg: https://opdb.org/machines/2103
'		if Scorbit.DoInit(389, "PupOverlays", "1.0.0", "GRWvz-MP37P") then 
'			tmrScorbit.Interval=2000
'			tmrScorbit.UserValue = 0
'			tmrScorbit.Enabled=True 
'		End if 
' 3) Customize helper functions below for different events if you want or make your own 
' 4) Call 
'		DoInit - After Pup/Screen is setup (PuPInit)
'		StartSession - When a game starts (ResetForNewGame)
'		StopSession - When the game is over (Table1_Exit, EndOfGame)
'		SendUpdate - called when Score Changes (AddScore)
'			SendUpdate(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers)
'			Example:  Scorbit.SendUpdate Score(0), Score(1), Score(2), Score(3), Balls, CurrentPlayer+1, PlayersPlayingGame
'		SetGameMode - When different game events happen like starting a mode, MB etc.  (ScorbitBuildGameModes helper function shows you how)
' 5) Drop the binaries sQRCode.exe and sToken.exe in your Pup Root so we can create session tokens and QRCodes.
'	- Drop QRCode Images (QRCodeS.png, QRcodeB.png) in yur pup PuPOverlays if you want to use those 
' 6) Callbacks 
'		Scorbit_Paired   	- Called when machine is successfully paired.  Hide QRCode and play a sound 
'		Scorbit_PlayerClaimed	- Called when player is claimed.  Hide QRCode, play a sound and display name 
'		ScorbitClaimQR		- Call before/after plunge (swPlungerRest_Hit, swPlungerRest_UnHit)
' 7) Other 
'		Set Pair QR Code	- During Attract
'			if (Scorbit.bNeedsPairing) then 
'				PuPlayer.LabelSet pDMDFull, "ScorbitQR_a", "PuPOverlays\\QRcode.png",1,"{'mt':2,'width':32, 'height':64,'xalign':0,'yalign':0,'ypos':5,'xpos':5}"
'				PuPlayer.LabelSet pDMDFull, "ScorbitQRIcon_a", "PuPOverlays\\QRcodeS.png",1,"{'mt':2,'width':36, 'height':85,'xalign':0,'yalign':0,'ypos':3,'xpos':3,'zback':1}"
'			End if 
'		Set Player Names 	- Wherever it makes sense but I do it here: (pPupdateScores)
'		   if ScorbitActive then 
'			if Scorbit.bSessionActive then
'				PlayerName=Scorbit.GetName(CurrentPlayer+1)
'				if PlayerName="" then PlayerName= "Player " & CurrentPlayer+1 
'			End if 
'		   End if 
'
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' TABLE CUSTOMIZATION START HERE 

Sub Scorbit_Paired()								' Scorbit callback when new machine is paired 
dbg2 "Scorbit PAIRED"
	PlaySound "scorbit_login"

	' Only run if first time pairing
	if bUnpaired Then
		ResetOverlay
		GeneralPupQueue.Add "ResetOverlay","ResetOverlay",25,100,0,0,0,False
		pbackglasslabelhide "ScorbitQR1"
	End If
End Sub 


Sub Scorbit_PlayerClaimed(PlayerNum, PlayerName)	' Scorbit callback when QR Is Claimed 
dbg2 "Scorbit LOGIN"
	PlaySound "scorbit_login"
	ScorbitClaimQR(False)
End Sub 


Sub ScorbitClaimQR(bShow)	
dbg2 "In ScorbitClaimQR: " &bShow					'  Show QRCode on first ball for users to claim this position
dbg2 "Session Active: " &Scorbit.bSessionActive
dbg2 "bNeedsPairing:" &Scorbit.bNeedsPairing
	if Scorbit.bSessionActive=False then Exit Sub 
	if ScorbitShowClaimQR=False then Exit Sub
	if Scorbit.bNeedsPairing then exit sub 


dbg2 "BShow: " & bShow
dbg2 "First Ball: " &bOnTheFirstBallScorbit
'dbg2 "bGameInPlay: " & bGameStarted
dbg2 "GetName: " &Scorbit.GetName(CurrentPlayer)

	if bShow and bOnTheFirstBallScorbit and Scorbit.GetName(CurrentPlayer)="" then 
		if DMDType = 2 Then 
		pdmdsetpage 77
		PuPlayer.playlistplayex pDMD,"PuPOverlays","Scorbit_Claim.png",0,1
		PuPlayer.LabelSet pDMD, "ScorbitQR2", "PuPOverlays\\QRclaim.png",1,"{'mt':2,'width':21.25, 'height':37.5,'xalign':0,'yalign':0,'ypos':31.,'xpos':73.8}"
		Else
			PuPlayer.playlistplayex pBackglass,"PuPOverlays","Scorbit_Claim.png",0,1
			PuPlayer.LabelSet pBackglass, "ScorbitQR2", "PuPOverlays\\QRclaim.png",1,"{'mt':2,'width':21.25, 'height':37.5,'xalign':0,'yalign':0,'ypos':31.,'xpos':73.8}"
		End If
	Else 
		dbg2 "Hiding QR claim & Overlay"
		HideScorbit
	End if 
End Sub 

Sub StopScorbit
	Scorbit.StopSession Score(1), Score(2), Score(3), Score(4), PlayersPlayingGame   ' Stop updateing scores
End Sub


Sub ScorbitBuildGameModes()		' Custom function to build the game modes for better stats 
	dim GameModeStr
	if Scorbit.bSessionActive=False then Exit Sub 
	Dbg2 " Should be adding Game String:" &GameModeStr
	Scorbit.SetGameMode(GameModeStr)
End Sub 

Sub DelayQRClaim_Timer()
	if bOnTheFirstBall AND bBallInPlungerLane then 
		ScorbitClaimQR(True)
		Dbg2 " Should be calling Show Claim"
	End If
'	 ScorbitClaimQR(True)
'	DelayQRClaim.Enabled=False
End Sub

sub CheckPairing
dbg2 "Inside SCORBIT check pairing"
	if (Scorbit.bNeedsPairing) then 
		dbg2 "Should be displaying pairing info"
		if DMDType = 2 Then 
			pdmdsetpage 77
			PuPlayer.playlistplayex pDMD,"PuPOverlays","Scorbit_Pair.png",0,1
			PuPlayer.LabelSet pDMD, "ScorbitQR1", "PuPOverlays\\QRcode.png",1,"{'mt':2,'width':21.25, 'height':37.5,'xalign':0,'yalign':0,'ypos':31.,'xpos':73.8}"
		Else
			PuPlayer.playlistplayex pBackglass,"PuPOverlays","Scorbit_Pair.png",0,1
			PuPlayer.LabelSet pBackglass, "ScorbitQR1", "PuPOverlays\\QRcode.png",1,"{'mt':2,'width':21.25, 'height':37.5,'xalign':0,'yalign':0,'ypos':31.,'xpos':73.8}"
		End If

		DelayQRClaim.Interval=6000
		DelayQRClaim.Enabled=True
	Else
dbg2 "Already Paired"
		DelayQRClaim.Interval=6000
		DelayQRClaim.Enabled=True
	end if
End sub

Sub HideScorbit
	pDMDsetpage pScores
	DelayQRClaim.Enabled = False
	if DMDType = 2 Then
		PuPlayer.playlistplayex pBackglass,"PuPOverlays","defaultDMD.png",0,1
	Elseif PlatformOS <> "windows" Then

	Else
'		PuPlayer.playevent pBackglass,"Backglass","Blank.mp4",0,20,6,0,""
		PuPlayer.playlistplayex pBackglass,"PuPOverlays","Card0.png",0,1
		'if renderingmode = 2 then PinCab_Backglass.image = "card1"
	End If
	pBackglasslabelhide "ScorbitQR1"
	pBackglasslabelhide "ScorbitQRIcon1"
	pBackglasslabelhide "ScorbitQR2"
	pBackglasslabelhide "ScorbitQRIcon2"
End Sub




' END ----------

Sub Scorbit_LOGUpload(state)	' Callback during the log creation process.  0=Creating Log, 1=Uploading Log, 2=Done 
	Select Case state 
		case 0:
			dbg "CREATING LOG"
		case 1:
			dbg "Uploading LOG"
		case 2:
			dbg "LOG Complete"
	End Select 
End Sub 
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' TABLE CUSTOMIZATION END HERE - NO NEED TO EDIT BELOW THIS LINE




' Workaround - Call get a reference to Member Function
Sub tmrScorbit_Timer()								' Timer to send heartbeat 
	Scorbit.DoTimer(tmrScorbit.UserValue)
	tmrScorbit.UserValue=tmrScorbit.UserValue+1
	if tmrScorbit.UserValue>5 then tmrScorbit.UserValue=0
End Sub 
Function ScorbitIF_Callback()
	Scorbit.Callback()
End Function 
Class ScorbitIF

	Public bSessionActive
	Public bNeedsPairing
	Private bUploadLog
	Private bActive
	Private LOGFILE(10000000)
	Private LogIdx

	Private bProduction

	Private TypeLib
	Private MyMac
	Private Serial
	Private MyUUID
	Private TableVersion

	Private SessionUUID
	Private SessionSeq
	Private SessionTimeStart
	Private bRunAsynch
	Private bWaitResp
	Private GameMode
	Private GameModeOrig		' Non escaped version for log
	Private VenueMachineID
	Private CachedPlayerNames(4)
	Private SaveCurrentPlayer

	Public bEnabled
	Private sToken
	Private machineID
	Private dirQRCode
	Private opdbID
	Private wsh

	Private objXmlHttpMain
	Private objXmlHttpMainAsync
	Private fso
	Private Domain

	Public Sub Class_Initialize()
		bActive="false"
		bSessionActive=False
		bEnabled=False 
	End Sub 

	Property Let UploadLog(bValue)
		bUploadLog = bValue
	End Property

	Sub DoTimer(bInterval)	' 2 second interval
		dim holdScores(4)
		dim i
		if bInterval=0 then 
			SendHeartbeat()
		elseif bRunAsynch And bSessionActive = True then ' Game in play (Updated for TNA to resolve stutter in CoopMode)
			Scorbit.SendUpdate Score(1), Score(2), Score(3), Score(4), balls, CurrentPlayer, PlayersPlayingGame
		End if 
	End Sub 

	Function GetName(PlayerNum)	' Return Parsed Players name  
		if PlayerNum<1 or PlayerNum>4 then 
			GetName=""
		else 
			GetName=CachedPlayerNames(PlayerNum-1)
		End if 
	End Function 

	Function DoInit(MyMachineID, Directory_PupQRCode, Version, opdb)
		dim Nad
		Dim EndPoint
		Dim resultStr 
		Dim UUIDParts 
		Dim UUIDFile

		bProduction=1
'		bProduction=0
		SaveCurrentPlayer=0
		VenueMachineID=""
		bWaitResp=False 
		bRunAsynch=False 
		DoInit=False 
		opdbID=opdb
		dirQrCode=Directory_PupQRCode
		MachineID=MyMachineID
		TableVersion=version
		bNeedsPairing=False
		if bProduction then 
			domain = "api.scorbit.io"
		else 
			domain = "staging.scorbit.io"
			domain = "scorbit-api-staging.herokuapp.com"
		End if 
		Set fso = CreateObject("Scripting.FileSystemObject")
		dim objLocator:Set objLocator = CreateObject("WbemScripting.SWbemLocator")
		Dim objService:Set objService = objLocator.ConnectServer(".", "root\cimv2")
		Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
		Set objXmlHttpMainAsync = CreateObject("Microsoft.XMLHTTP")
		objXmlHttpMain.onreadystatechange = GetRef("ScorbitIF_Callback")
		Set wsh = CreateObject("WScript.Shell")

		' Get Mac for Serial Number 
		dim Nads: set Nads = objService.ExecQuery("Select * from Win32_NetworkAdapter where physicaladapter=true")
		for each Nad in Nads
			if not isnull(Nad.MACAddress) then
				if left(Nad.MACAddress, 6)<>"00090F" then ' Skip over forticlient MAC
dbg2 "Using MAC Addresses:" & Nad.MACAddress & " From Adapter:" & Nad.description   
					MyMac=replace(Nad.MACAddress, ":", "")
					Exit For 
				End if 
			End if 
		Next
		Serial=eval("&H" & mid(MyMac, 5))
		if Serial<0 then Serial=eval("&H" & mid(MyMac, 6))		' Mac Address Overflow Special Case 
		if MyMachineID<>2108 then 			' GOTG did it wrong but MachineID should be added to serial number also
			Serial=Serial+MyMachineID
		End if 
'		Serial=123456
		dbg2 "Serial:" & Serial

		' Get System UUID
		set Nads = objService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
		for each Nad in Nads
			dbg2 "Using UUID:" & Nad.UUID   
			MyUUID=Nad.UUID
			Exit For 
		Next

		if MyUUID="" then 
			MsgBox "SCORBIT - Can get UUID, Disabling."
			Exit Function
		elseif MyUUID="03000200-0400-0500-0006-000700080009" or ScorbitAlternateUUID then
			If fso.FolderExists(UserDirectory) then 
				If fso.FileExists(UserDirectory & "ScorbitUUID.dat") then
					Set UUIDFile = fso.OpenTextFile(UserDirectory & "ScorbitUUID.dat",1)
					MyUUID = UUIDFile.ReadLine()
					UUIDFile.Close
					Set UUIDFile = Nothing
				Else 
					MyUUID=GUID()
					Set UUIDFile=fso.CreateTextFile(UserDirectory & "ScorbitUUID.dat",True)
					UUIDFile.WriteLine MyUUID
					UUIDFile.Close
					Set UUIDFile=Nothing
				End if
			End if 
		End if

		' Clean UUID
		UUIDParts=split(MyUUID, "-")
		MyUUID=LCASE(Hex(eval("&h" & UUIDParts(0))+MyMachineID) & UUIDParts(1) &  UUIDParts(2) &  UUIDParts(3) & UUIDParts(4))		 ' Add MachineID to UUID
		MyUUID=LPad(MyUUID, 32, "0")
'		MyUUID=Replace(MyUUID, "-",  "")
		dbg2 "MyUUID:" & MyUUID 


		' Authenticate and get our token 
		if getStoken() then 
			bEnabled=True 
'			SendHeartbeat
			DoInit=True
		End if 
	End Function 

	Sub Callback()
		Dim ResponseStr
		Dim i 
		Dim Parts
		Dim Parts2
		Dim Parts3
		if bEnabled=False then Exit Sub 

		if bWaitResp and objXmlHttpMain.readystate=4 then 
'			dbg2 "CALLBACK: " & objXmlHttpMain.Status & " " & objXmlHttpMain.readystate
			if objXmlHttpMain.Status=200 and objXmlHttpMain.readystate = 4 then 
				ResponseStr=objXmlHttpMain.responseText
				'debug3 "RESPONSE: " & ResponseStr

				' Parse Name 
				If bSessionActive = True Then
					if CachedPlayerNames(SaveCurrentPlayer-1)="" then  ' Player doesnt have a name
						if instr(1, ResponseStr, "cached_display_name") <> 0 Then	' There are names in the result
							Parts=Split(ResponseStr,",{")							' split it 
							if ubound(Parts)>=SaveCurrentPlayer-1 then 				' Make sure they are enough avail
								if instr(1, Parts(SaveCurrentPlayer-1), "cached_display_name")<>0 then 	' See if mine has a name 
									CachedPlayerNames(SaveCurrentPlayer-1)=GetJSONValue(Parts(SaveCurrentPlayer-1), "cached_display_name")		' Get my name
									CachedPlayerNames(SaveCurrentPlayer-1)=Replace(CachedPlayerNames(SaveCurrentPlayer-1), """", "")
									Scorbit_PlayerClaimed SaveCurrentPlayer, CachedPlayerNames(SaveCurrentPlayer-1)
	'								dbg2 "Player Claim:" & SaveCurrentPlayer & " " & CachedPlayerNames(SaveCurrentPlayer-1)
								End if 
							End if
						End if 
					else												    ' Check for unclaim 
						if instr(1, ResponseStr, """player"":null")<>0 Then	' Someone doesnt have a name
							Parts=Split(ResponseStr,"[")						' split it 
	'dbg2 "Parts:" & Parts(1)
							Parts2=Split(Parts(1),"}")							' split it 
							for i = 0 to Ubound(Parts2)
	'dbg2 "Parts2:" & Parts2(i)
								if instr(1, Parts2(i), """player"":null")<>0 Then
									CachedPlayerNames(i)=""
								End if 
							Next 
						End if 
					End if
				End If

				'Check heartbeat
				HandleHeartbeatResp ResponseStr
			End if 
			bWaitResp=False
		End if 
	End Sub

	Public Sub StartSession()
		if bEnabled=False then Exit Sub 
		dbg2 "Scorbit Start Session" 
		CachedPlayerNames(0)=""
		CachedPlayerNames(1)=""
		CachedPlayerNames(2)=""
		CachedPlayerNames(3)=""
		bRunAsynch=True 
		bActive="true"
		bSessionActive=True
		SessionSeq=0
		SessionUUID=GUID()
		SessionTimeStart=GameTime
		LogIdx=0
		SendUpdate 0, 0, 0, 0, 1, 1, 1
	End Sub

	' Custom method for TNA to work around coop mode stuttering
	Public Sub ForceAsynch(enabled)
		if bEnabled=False then Exit Sub
		if bSessionActive=True then Exit Sub 'Sessions should always control asynch when active
		bRunAsynch=enabled
	End Sub

	Public Sub StopSession(P1Score, P2Score, P3Score, P4Score, NumberPlayers)
		StopSession2 P1Score, P2Score, P3Score, P4Score, NumberPlayers, False
	End Sub 

	Public Sub StopSession2(P1Score, P2Score, P3Score, P4Score, NumberPlayers, bCancel)
		Dim i
		dim objFile
		if bEnabled=False then Exit Sub 
		bRunAsynch=False 'Asynch might have been forced on in TNA to prevent coop mode stutter
		if bSessionActive=False then Exit Sub 
dbg2 "Scorbit Stop Session" 

		bActive="false" 
		SendUpdate P1Score, P2Score, P3Score, P4Score, -1, -1, NumberPlayers
		bSessionActive=False
'		SendHeartbeat

		if bUploadLog and LogIdx<>0 and bCancel=False then 
			dbg2 "Creating Scorbit Log: Size" & LogIdx
			Scorbit_LOGUpload(0)
			Set objFile = fso.CreateTextFile(puplayer.getroot & pgamename & "\sGameLog.csv")
			For i = 0 to LogIdx-1 
				objFile.Writeline LOGFILE(i)
			Next 
			objFile.Close
			LogIdx=0
			Scorbit_LOGUpload(1)
			pvPostFile "https://" & domain & "/api/session_log/", puplayer.getroot & pgamename & "\sGameLog.csv", False
			Scorbit_LOGUpload(2)
			on error resume next
			fso.DeleteFile(puplayer.getroot & pgamename & "\sGameLog.csv")
			on error goto 0
		End if 

	End Sub 

	Public Sub SetGameMode(GameModeStr)
		GameModeOrig=GameModeStr
		GameMode=GameModeStr
		GameMode=Replace(GameMode, ":", "%3a")
		GameMode=Replace(GameMode, ";", "%3b")
		GameMode=Replace(GameMode, " ", "%20")
		GameMode=Replace(GameMode, "{", "%7B")
		GameMode=Replace(GameMode, "}", "%7D")
	End sub 

	Public Sub SendUpdate(P1Score, P2Score, P3Score, P4Score, CurrentBall, nPlayer, NumberPlayers)
		SendUpdateAsynch P1Score, P2Score, P3Score, P4Score, CurrentBall, nPlayer, NumberPlayers, bRunAsynch
	End Sub 

	Public Sub SendUpdateAsynch(P1Score, P2Score, P3Score, P4Score, CurrentBall, nPlayer, NumberPlayers, bAsynch)
		dim i
		Dim PostData
		Dim resultStr
		dim LogScores(4)

		if bUploadLog then 
			if NumberPlayers>=1 then LogScores(0)=P1Score
			if NumberPlayers>=2 then LogScores(1)=P2Score
			if NumberPlayers>=3 then LogScores(2)=P3Score
			if NumberPlayers>=4 then LogScores(3)=P4Score
			LOGFILE(LogIdx)=DateDiff("S", "1/1/1970", Now()) & "," & LogScores(0) & "," & LogScores(1) & "," & LogScores(2) & "," & LogScores(3) & ",,," &  nPlayer & "," & CurrentBall & ",""" & GameModeOrig & """"
			LogIdx=LogIdx+1
		End if


		if bSessionActive=False then Exit Sub 

		if bEnabled=False then Exit Sub 

		if bWaitResp then exit sub ' Drop message until we get our next response 

		SaveCurrentPlayer=nPlayer
		PostData = "session_uuid=" & SessionUUID & "&session_time=" & GameTime-SessionTimeStart+1 & _
					"&session_sequence=" & SessionSeq & "&active=" & bActive

		SessionSeq=SessionSeq+1
		if NumberPlayers > 0 then 
			for i = 0 to NumberPlayers-1
				PostData = PostData & "&current_p" & i+1 & "_score="
				if i <= NumberPlayers-1 then 
					if i = 0 then PostData = PostData & P1Score
					if i = 1 then PostData = PostData & P2Score
					if i = 2 then PostData = PostData & P3Score
					if i = 3 then PostData = PostData & P4Score
				else 
					PostData = PostData & "-1"
				End if 
			Next 
'Dbg2 "Score:" &P1Score &" XXX"
			PostData = PostData & "&current_ball=" & CurrentBall & "&current_player=" & nPlayer
			if GameMode<>"" then PostData=PostData & "&game_modes=" & GameMode

		End if 
		resultStr = PostMsg("https://" & domain, "/api/entry/", PostData, bAsynch)
		'if resultStr<>"" then debug3 "SendUpdate Resp:" & resultStr    			'rtp12
	End Sub 

' PRIVATE BELOW 
	Private Function LPad(StringToPad, Length, CharacterToPad)
	  Dim x : x = 0
	  If Length > Len(StringToPad) Then x = Length - len(StringToPad)
	  LPad = String(x, CharacterToPad) & StringToPad
	End Function

	Private Function GUID()		
		Dim TypeLib
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
		GUID = Mid(TypeLib.Guid, 2, 36)
	End Function

	Private Function GetJSONValue(JSONStr, key)
		dim i 
		Dim tmpStrs,tmpStrs2
		if Instr(1, JSONStr, key)<>0 then 
			tmpStrs=split(JSONStr,",")
			for i = 0 to ubound(tmpStrs)
				if instr(1, tmpStrs(i), key)<>0 then 
					tmpStrs2=split(tmpStrs(i),":")
					GetJSONValue=tmpStrs2(1)
					exit for
				End if 
			Next 
		End if 
	End Function

	Private Sub SendHeartbeat()
		Dim resultStr
		if bEnabled=False then Exit Sub 
		resultStr = GetMsgHdr("https://" & domain, "/api/heartbeat/", "Authorization", "SToken " & sToken)
		
		'Customized for TNA
		If bRunAsynch = False Then 
			dbg2 "Heartbeat Resp:" & resultStr
			HandleHeartbeatResp ResultStr
		End If
	End Sub 

	'TNA custom method
	Private Sub HandleHeartbeatResp(resultStr)
		dim TmpStr
		Dim Command
		Dim rc
		'Dim QRFile:QRFile=puplayer.getroot&"\" & pgamename & "\" & dirQrCode
		Dim QRFile:QRFile=puplayer.getroot & pgamename & "\" & dirQrCode
'dbg2 "QRFile: " &QRFile
		If VenueMachineID="" then
			If resultStr<>"" And Not InStr(resultStr, """machine_id"":" & machineID)=0 Then 'We Paired
				bNeedsPairing=False
				dbg2 "Scorbit: Paired"
				Scorbit_Paired()
			ElseIf resultStr<>"" And Not InStr(resultStr, """unpaired"":true")=0 Then 'We Did not Pair
				dbg2 "Scorbit: NOT Paired"
				bNeedsPairing=True
				bUnpaired = True
			Else
				' Error (or not a heartbeat); do nothing
			End If

			TmpStr=GetJSONValue(resultStr, "venuemachine_id")
			if TmpStr<>"" then 
				VenueMachineID=TmpStr
'dbg2 "VenueMachineID=" & VenueMachineID			
				'Command = """" & puplayer.getroot&"\" & pgamename & "\sQRCode.exe"" " & VenueMachineID & " " & opdbID & " """ & QRFile & """"
				Command = """" & puplayer.getroot & pgamename & "\sQRCode.exe"" " & VenueMachineID & " " & opdbID & " """ & QRFile & """"
				rc = wsh.Run(Command, 0, False)
			End if 
		End if
	End Sub

	Private Function getStoken()
		Dim result
		Dim results
'		dim wsh
		Dim tmpUUID:tmpUUID="adc12b19a3504453a7414e722f58736b"
		Dim tmpVendor:tmpVendor="vscorbitron"
		Dim tmpSerial:tmpSerial="999990104"
		'Dim QRFile:QRFile=puplayer.getroot&"\" & pgamename & "\" & dirQrCode
		Dim QRFile:QRFile=puplayer.getroot & pgamename & "\" & dirQrCode
		'Dim sTokenFile:sTokenFile=puplayer.getroot&"\" & pgamename & "\sToken.dat"
		Dim sTokenFile:sTokenFile=puplayer.getroot & pgamename & "\sToken.dat"

		' Set everything up
		tmpUUID=MyUUID
		tmpVendor="vpin"
		tmpSerial=Serial
		
		on error resume next
		fso.DeleteFile(sTokenFile)
		On error goto 0 

		' get sToken and generate QRCode
'		Set wsh = CreateObject("WScript.Shell")
		Dim waitOnReturn: waitOnReturn = True
		Dim windowStyle: windowStyle = 0
		Dim Command 
		Dim rc
		Dim objFileToRead

		'Command = """" & puplayer.getroot&"\" & pgamename & "\sToken.exe"" " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " """ & QRFile & """ """ & sTokenFile & """ " & domain
		Command = """" & puplayer.getroot & pgamename & "\sToken.exe"" " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " """ & QRFile & """ """ & sTokenFile & """ " & domain
dbg2 "RUNNING Command:" & Command
		rc = wsh.Run(Command, windowStyle, waitOnReturn)
dbg2 "Return:" & rc
		if FileExists(puplayer.getroot&"\" & pgamename & "\sToken.dat") and rc=0 then
			Set objFileToRead = fso.OpenTextFile(puplayer.getroot&"\" & pgamename & "\sToken.dat",1)
			result = objFileToRead.ReadLine()
			objFileToRead.Close
			Set objFileToRead = Nothing

			if Instr(1, result, "Invalid timestamp")<> 0 then 
				MsgBox "Scorbit Timestamp Error: Please make sure the time on your system is exact"
				getStoken=False
			elseif Instr(1, result, ":")<>0 then 
				results=split(result, ":")
				sToken=results(1)
				sToken=mid(sToken, 3, len(sToken)-4)
dbg2 "Got TOKEN:" & sToken
				getStoken=True
			Else 
dbg2 "ERROR:" & result
				getStoken=False
			End if 
		else 
dbg2 "ERROR No File:" & rc
		End if 

	End Function 

	private Function FileExists(FilePath)
		If fso.FileExists(FilePath) Then
			FileExists=CBool(1)
		Else
			FileExists=CBool(0)
		End If
	End Function

	Private Function GetMsg(URLBase, endpoint)
		GetMsg = GetMsgHdr(URLBase, endpoint, "", "")
	End Function

	Private Function GetMsgHdr(URLBase, endpoint, Hdr1, Hdr1Val)
		Dim Url
		Url = URLBase + endpoint & "?session_active=" & bActive
'dbg2 "Url:" & Url  & "  Async=" & bRunAsynch
		objXmlHttpMain.open "GET", Url, bRunAsynch
'		objXmlHttpMain.setRequestHeader "Content-Type", "text/xml"
		objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
		if Hdr1<> "" then objXmlHttpMain.setRequestHeader Hdr1, Hdr1Val

'		on error resume next
			err.clear
			objXmlHttpMain.send ""
			if err.number=-2147012867 then 
				MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
				bEnabled=False
			elseif err.number <> 0 then 
				debug3 "Server error: (" & err.number & ") " & Err.Description
			End if 
			if bRunAsynch=False then 
dbg2 "Status: " & objXmlHttpMain.status
				If objXmlHttpMain.status = 200 Then
					GetMsgHdr = objXmlHttpMain.responseText
				Else 
					GetMsgHdr=""
				End if 
			Else 
				bWaitResp=True
				GetMsgHdr=""
			End if 
'		On error goto 0

	End Function

	Private Function PostMsg(URLBase, endpoint, PostData, bAsynch)
		Dim Url

		Url = URLBase + endpoint
'dbg2 "PostMSg:" & Url & " " & PostData			'rtp12

		objXmlHttpMain.open "POST",Url, bAsynch
		objXmlHttpMain.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXmlHttpMain.setRequestHeader "Content-Length", Len(PostData)
		objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
		objXmlHttpMain.setRequestHeader "Authorization", "SToken " & sToken
		if bAsynch then bWaitResp=True 

		on error resume next
			objXmlHttpMain.send PostData
			if err.number=-2147012867 then 
				MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
				bEnabled=False
			elseif err.number <> 0 then 
				'debug3 "Multiplayer Server error (" & err.number & ") " & Err.Description
			End if 
			If objXmlHttpMain.status = 200 Then
				PostMsg = objXmlHttpMain.responseText
			else 
				PostMsg="ERROR: " & objXmlHttpMain.status & " >" & objXmlHttpMain.responseText & "<"
			End if 
		On error goto 0
	End Function

	Private Function pvPostFile(sUrl, sFileName, bAsync)
'dbg2 "Posting File " & sUrl & " " & sFileName & " " & bAsync & " File:" & Mid(sFileName, InStrRev(sFileName, "\") + 1)
		Dim STR_BOUNDARY:STR_BOUNDARY  = GUID()
		Dim nFile  
		Dim baBuffer()
		Dim sPostData
		Dim Response

		'--- read file
		Set nFile = fso.GetFile(sFileName)
		With nFile.OpenAsTextStream()
			sPostData = .Read(nFile.Size)
			.Close
		End With


		'--- prepare body
		sPostData = "--" & STR_BOUNDARY & vbCrLf & _
			"Content-Disposition: form-data; name=""uuid""" & vbCrLf & vbCrLf & _
			SessionUUID & vbcrlf & _
			"--" & STR_BOUNDARY & vbCrLf & _
			"Content-Disposition: form-data; name=""log_file""; filename=""" & SessionUUID & ".csv""" & vbCrLf & _
			"Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
			sPostData & vbCrLf & _
			"--" & STR_BOUNDARY & "--"


		'--- post
		With objXmlHttpMain
			.Open "POST", sUrl, bAsync
			.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
			.SetRequestHeader "Authorization", "SToken " & sToken
			.Send sPostData ' pvToByteArray(sPostData)
			If Not bAsync Then
				Response= .ResponseText
				pvPostFile = Response
dbg2 "Upload Response: " & Response
			End If
		End With

	End Function

	Private Function pvToByteArray(sText)
		pvToByteArray = StrConv(sText, 128)		' vbFromUnicode
	End Function

End Class 

'  END SCORBIT 
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub TargetResetTimer_Timer()
	RaiseTargetTimer_Timer
End Sub

'===========================================================
' TWO FLASHER COUNTDOWN (TRIGGERED BY target009 DROP)
'===========================================================

Dim CountdownValue
Dim CountdownRunning

'----------------------------------------
' TIMER — make a timer named "CountdownTimer" in VPX
'----------------------------------------
Sub CountdownTimer_Timer()
    ' Hide instantly if a Battle or Blizzard mode begins
'    If Battle(CurrentPlayer, 0) <> 0 Or bBLIZZARDMode Then
'        ResetCountdown
'        Exit Sub
'    End If


    If Not CountdownRunning Then Exit Sub

    CountdownValue = CountdownValue - 1
    If CountdownValue < 0 Then CountdownValue = 0 : debug.print "SHOULD RESET"

    UpdateCountdownFlashers

    If CountdownValue = 0 Then
		bBlizzardPrepMode = False
		ResetCountdown
    End If
End Sub

'----------------------------------------
' Start and Reset
'----------------------------------------
Sub StartCountdown()
     ' Skip countdown if in a Battle, Blizzard, or Wizard mode
    If Battle(CurrentPlayer, 0) <> 0 _
        Or bBLIZZARDMode _
        Or bWizMode1Active _
        Or bWizMode2Active _
        Or bWizMode3Active Then Exit Sub

	pDMDLabelSetBorder "Event3A",cWhite,3,3,1
	pDMDLabelSetBorder "Event3B",cWhite,3,3,1
	pDMDLabelSetBorder "Event3Ca",cBlue,3,3,1
	pDMDLabelSetBorder "Event3C",cWhite,3,3,1

	EnableBlizzardTargetLights
    CountdownValue = 60
    CountdownRunning = True
    CountdownTimer.Interval = 1000
    CountdownTimer.Enabled = True
    FlasherTens.Visible = True
    FlasherOnes.Visible = True
    UpdateCountdownFlashers
End Sub

Sub ResetCountdown()
	nBlizzTargetCount = 0
    CountdownRunning = False
    CountdownTimer.Enabled = False
    CountdownValue = 60
    FlasherTens.Visible = False
    FlasherOnes.Visible = False
    UpdateCountdownFlashers
	bBlizzardPrepMode = False ' make sure mode is off

	Dim i
	for i = 1 to 8
		blizzardletters(i) = 0
	Next

	HideBlizzTimer
	PuPlayer.LabelSet pDMD,"Event3A","",0,""
	PuPlayer.LabelSet pDMD,"Event3B","",0,""
	PuPlayer.LabelSet pDMD,"Event3C","",0,""
	PuPlayer.LabelSet pDMD,"Event3Ca","",0,""

End Sub



Sub EnableBlizzardTargetLights
	Li001.State = 2
    Li021.State = 2
	Li022.State = 2
	Li023.State = 2
	Li024.State = 2
	Li025.State = 2
	Li026.State = 2
	Li027.State = 2
	Li028.State = 2
End Sub

Sub DisableBlizzardTargetLights
debug.print "Disable Blizz lights"
	Li001.State = 0
    Li021.State = 0
	Li022.State = 0
	Li023.State = 0
	Li024.State = 0
	Li025.State = 0
	Li026.State = 0
	Li027.State = 0
	Li028.State = 0
End Sub

'----------------------------------------
' Update Flasher Images
'----------------------------------------
Sub UpdateCountdownFlashers()
    Dim tens, ones
    tens = Int(CountdownValue / 10)
    ones = CountdownValue Mod 10

    FlasherTens.ImageA = "d_" & tens
    FlasherOnes.ImageA = "d_" & ones
End Sub

'**************************
' FIRE ANIMATION (3 FLASHERS, FADE ON/OFF)
'**************************

Dim Fire1Pos, Fire2Pos, Fire3Pos
Dim FireFrames(80), Fire3Frames(80)
Dim i
Dim FireFadeDir    ' 1 = fade in, -1 = fade out, 0 = none
Dim FireOpacity

' Build frame list for SW_ images (Fire1 & Fire2)
For i = 1 To 81
    FireFrames(i - 1) = "SW_" & i
Next

' Build frame list for BW_ images (Fire3)
For i = 1 To 81
    Fire3Frames(i - 1) = "BW_" & i
Next

'**************************
' START / STOP SUBS
'**************************

Sub StartFire()
    Fire1Pos = 0
    Fire2Pos = 40
    Fire3Pos = 0

    FireOpacity = 0
    FireFadeDir = 1   ' start fade in

    Fire1.Visible = True
    Fire2.Visible = True
    Fire3.Visible = True

    FireTimer.Enabled = True
    FireFadeTimer.Enabled = True
End Sub

Sub StopFire()
    FireFadeDir = -1  ' start fade out
    FireFadeTimer.Enabled = True
End Sub

'**************************
' FIRE ANIMATION LOOP
'**************************

Sub FireTimer_Timer()
    Fire1.ImageA = FireFrames(Fire1Pos)
    Fire2.ImageA = FireFrames(Fire2Pos)
    Fire3.ImageA = Fire3Frames(Fire3Pos)

    Fire1Pos = (Fire1Pos + 1) Mod 81
    Fire2Pos = (Fire2Pos + 1) Mod 81
    Fire3Pos = (Fire3Pos + 1) Mod 81
End Sub

'**************************
' FIRE FADE TIMER
'**************************

Sub FireFadeTimer_Timer()
    If FireFadeDir = 1 Then
        FireOpacity = FireOpacity + 0.05
        If FireOpacity >= 1 Then
            FireOpacity = 1
            FireFadeDir = 0
            FireFadeTimer.Enabled = False
        End If
    ElseIf FireFadeDir = -1 Then
        FireOpacity = FireOpacity - 0.05
        If FireOpacity <= 0 Then
            FireOpacity = 0
            FireFadeDir = 0
            FireFadeTimer.Enabled = False
            FireTimer.Enabled = False
            Fire1.Visible = False
            Fire2.Visible = False
            Fire3.Visible = False
        End If
    End If

    ' Apply opacity
    Fire1.IntensityScale = FireOpacity
    Fire2.IntensityScale = FireOpacity
    Fire3.IntensityScale = FireOpacity
End Sub

'**************************
' DEMFLAME ANIMATION (48 FRAMES, 2.5 SECONDS, FADE IN/OUT, NO LOOP)
'**************************

Dim DemFlamePos
Dim DemFlameFrames(47)
Dim DemFlameOpacity, DemFlameFadeDir  ' 1 = fade in, -1 = fade out, 0 = idle

' Build frame list
Dim j
For j = 0 To 47
    DemFlameFrames(j) = "Fl_" & (j + 1)
Next

'**************************
' START / STOP SUBS
'**************************

Sub StartDemFlame()
    Playsound "wiz_start"
    DemFlamePos = 0
    DemFlameOpacity = 0
    DemFlameFadeDir = 1       ' fade in
    demflame.Visible = True
    demflame.IntensityScale = 0

    tDemFlame.Interval = 52   ' 48 frames * 52ms ≈ 2.5 seconds
    tDemFlame.Enabled = True
    tDemFlameFade.Enabled = True

    ' Run for 2.5 seconds total, then stop
    tDemFlameStop.Interval = 2500
    tDemFlameStop.Enabled = True
End Sub

Sub StopDemFlame()
    DemFlameFadeDir = -1      ' fade out
    tDemFlameFade.Enabled = True
End Sub

'**************************
' DEMFLAME FRAME LOOP
'**************************

Sub tDemFlame_Timer()
    demflame.ImageA = DemFlameFrames(DemFlamePos)
    If DemFlamePos < 47 Then
        DemFlamePos = DemFlamePos + 1
    Else
        ' Hold last frame until fade-out
        tDemFlame.Enabled = False
    End If
End Sub

'**************************
' DEMFLAME FADE TIMER
'**************************

Sub tDemFlameFade_Timer()
    If DemFlameFadeDir = 1 Then
        DemFlameOpacity = DemFlameOpacity + 0.05
        If DemFlameOpacity >= 1 Then
            DemFlameOpacity = 1
            DemFlameFadeDir = 0
            tDemFlameFade.Enabled = False
        End If
    ElseIf DemFlameFadeDir = -1 Then
        DemFlameOpacity = DemFlameOpacity - 0.05
        If DemFlameOpacity <= 0 Then
            DemFlameOpacity = 0
            DemFlameFadeDir = 0
            tDemFlameFade.Enabled = False
            demflame.Visible = False
        End If
    End If

    demflame.IntensityScale = DemFlameOpacity
End Sub

'**************************
' DEMFLAME STOP TIMER
'**************************

Sub tDemFlameStop_Timer()
    StopDemFlame
    tDemFlameStop.Enabled = False
End Sub


'==========================================
' LIGHTNING BURST EFFECT (F10 / F003 / F004)
'==========================================

Dim LightningBurstCount

Sub LightningStrike()
    LightningBurstCount = 0
    LightningBurst.Enabled = True
End Sub

Sub LightningBurst_Timer()
    ' hide all between flashes
    F10.Visible = False
    F003.Visible = False
    F004.Visible = False

    ' pick one of the three flashers randomly
    Dim pick
    pick = Int(Rnd * 3)

    Select Case pick
        Case 0: FlashForMs F10, 80 + Rnd*40, 50, 0
        Case 1: FlashForMs F003, 80 + Rnd*40, 50, 0
        Case 2: FlashForMs F004, 80 + Rnd*40, 50, 0
    End Select

    LightningBurstCount = LightningBurstCount + 1

    ' each strike will flicker 4 to 6 times
    If LightningBurstCount > 3 + Int(Rnd*3) Then
        LightningBurst.Enabled = False
    Else
        ' short random delay before next flash (50–200ms)
        LightningBurst.Interval = 40 + Int(Rnd * 250)
    End If
End Sub

'----------------------------------------
' ANIMATED MINIONS POP-UP FOLLOW WALLS
'----------------------------------------

Dim MinionSwap1, MinionSwap2
Dim MinionHurryUpMode
Dim MinionModeTimeLeft


'=== ANIMATION TIMERS (120ms) ===
Sub Minions1timer_Timer()
    ' Follows Minion1Wall height and alternates Min001 / Min002 visibility
    If MinionSwap1 = False Then
        Min001.z = Minion1wall.z
        Min001.visible = True
        Min002.visible = False
        MinionSwap1 = True
    Else
        Min002.z = Minion1wall.z
        Min002.visible = True
        Min001.visible = False
        MinionSwap1 = False
    End If
End Sub

Sub Minions2timer_Timer()
    ' Follows Minion2Wall height and alternates Min003 / Min004 visibility
    If MinionSwap2 = False Then
        Min003.z = Minion2wall.z
        Min003.visible = True
        Min004.visible = False
        MinionSwap2 = True
    Else
        Min004.z = Minion2wall.z
        Min004.visible = True
        Min003.visible = False
        MinionSwap2 = False
    End If
End Sub


'----------------------------------------
' MINION MODE CONTROL
'----------------------------------------
Sub StartMinionMode()
    MinionHurryUpMode = True
	ResetMinion1
	ResetMinion2
End Sub

Sub StopMinionMode()
    MinionHurryUpMode = False
    Minions1timer.Enabled = False
    Minions2timer.Enabled = False
    Minion1wallUp.Enabled = False
    Minion2wallUp.Enabled = False

    ' --- Force walls to retract fully ---
    Minion1wall.z = -110
    Minion2wall.z = -110
    Minion1wall.collidable = False
    Minion2wall.collidable = False

    ' --- Make sure down timers are off (since we manually forced them) ---
    Minion1wallDown.Enabled = False
    Minion2wallDown.Enabled = False

    ' --- Hide Minion primitives ---
    Min001.visible = False
    Min002.visible = False
    Min003.visible = False
    Min004.visible = False
    mag1.visible=False
    mag3.visible=False
    ' --- Reset positions so they don’t hover next mode ---
    Min001.z = -200
    Min002.z = -200
    Min003.z = -200
    Min004.z = -200
End Sub


Sub ResetMinion1
    Minions1timer.Enabled = True
	Minion1wallDown.enabled = False
	Minion1wallUp.enabled = True
End Sub

Sub ResetMinion2
    Minions2timer.Enabled = True
	Minion2wallDown.enabled = False
	Minion2wallUp.enabled = True
End Sub


'----------------------------------------
' POP-UP WALL CONTROL
'----------------------------------------

Sub Minion1wall_Hit()
     ' Stop Up timer if still running
    Minion1wallUp.Enabled = False
'    If Minion1wall.transz <> 0 Then Exit Sub
    PlaySound "DEMONHIT"
    mag3.visible=False
    Minion1wallDown.Enabled = True
    AddScore 50000

    ' --- Force reset primitives ---
    Min001.z = 0
    Min002.z = 0
End Sub

Sub Minion2wall_Hit()
     ' Stop Up timer if still running
    Minion2wallUp.Enabled = False
'    If Minion2wall.transz <> 0 Then Exit Sub
    PlaySound "DEMONHIT"
    mag1.visible=False
    Minion2wallDown.Enabled = True
    AddScore 50000

    ' --- Force reset primitives ---
    Min003.z = 0
    Min004.z = 0
End Sub


'=== WALL 1 UP/DOWN ===
Sub Minion1wallUp_Timer()
    If Minion1wall.z >= 0 Then
       PlaySound "DEMONPOPUP"
        mag3.visible=True
        StartDemFlame001   ' <<< ADD THIS
        Minion1wallUp.Enabled = False
        Exit Sub
    End If
 '   Minion1wall.visible = True
    Minion1wall.collidable = True
    Minion1wall.z = Minion1wall.z + 10
End Sub

Sub Minion1wallDown_Timer()
    If Minion1wall.z <= -110 Then
        Minion1wallDown.Enabled = False
 '       Minion1wall.visible = False
        Minion1wall.collidable = False
        Exit Sub
    End If
    Minion1wall.z = Minion1wall.z - 10
End Sub


'=== WALL 2 UP/DOWN ===
Sub Minion2wallUp_Timer()
    If Minion2wall.z >= 0 Then
       PlaySound "DEMONPOPUP"
        mag1.visible=True
        StartDemFlame002   ' <<< ADD THIS
        Minion2wallUp.Enabled = False
        Exit Sub
    End If
 '   Minion2wall.visible = True
   Minion2wall.collidable = True
    Minion2wall.z = Minion2wall.z + 10
	debug.print "MT@: " &Minion2wall.z
End Sub

Sub Minion2wallDown_Timer()
    If Minion2wall.z <= -110 Then
        Minion2wallDown.Enabled = False
 '       Minion2wall.visible = False
        Minion2wall.collidable = False
        Exit Sub
    End If
    Minion2wall.z = Minion2wall.z - 10
End Sub


'==========================================
' DEMON FLAME 001
'==========================================

Dim DemFlame001Pos, DemFlame001Opacity, DemFlame001FadeDir

Sub StartDemFlame001()
    DemFlame001Pos = 0
    DemFlame001Opacity = 0
    DemFlame001FadeDir = 1
    demflame001.Visible = True
    demflame001.IntensityScale = 0

    tDemFlame001.Interval = 52
    tDemFlame001.Enabled = True
    tDemFlame001Fade.Enabled = True

    tDemFlame001Stop.Interval = 300
    tDemFlame001Stop.Enabled = True
End Sub

Sub StopDemFlame001()
    DemFlame001FadeDir = -1
    tDemFlame001Fade.Enabled = True
End Sub

Sub tDemFlame001_Timer()
    demflame001.ImageA = DemFlameFrames(DemFlame001Pos)
    If DemFlame001Pos < 47 Then
        DemFlame001Pos = DemFlame001Pos + 1
    Else
        tDemFlame001.Enabled = False
    End If
End Sub

Sub tDemFlame001Fade_Timer()
    Dim easeStep
    If DemFlame001FadeDir = 1 Then
        ' Smooth fade in (ease in)
        easeStep = 0.05 * (1 - DemFlame001Opacity)
        DemFlame001Opacity = DemFlame001Opacity + easeStep
        If DemFlame001Opacity >= 0.99 Then
            DemFlame001Opacity = 1
            DemFlame001FadeDir = 0
            tDemFlame001Fade.Enabled = False
        End If
    ElseIf DemFlame001FadeDir = -1 Then
        ' Smooth fade out (ease out)
        easeStep = 0.05 * (DemFlame001Opacity)
        DemFlame001Opacity = DemFlame001Opacity - easeStep
        If DemFlame001Opacity <= 0.01 Then
            DemFlame001Opacity = 0
            DemFlame001FadeDir = 0
            tDemFlame001Fade.Enabled = False
            demflame001.Visible = False
        End If
    End If

    demflame001.IntensityScale = DemFlame001Opacity
End Sub


Sub tDemFlame001Stop_Timer()
    StopDemFlame001
    tDemFlame001Stop.Enabled = False
End Sub

'==========================================
' DEMON FLAME 002
'==========================================

Dim DemFlame002Pos, DemFlame002Opacity, DemFlame002FadeDir

Sub StartDemFlame002()
    DemFlame002Pos = 0
    DemFlame002Opacity = 0
    DemFlame002FadeDir = 1
    demflame002.Visible = True
    demflame002.IntensityScale = 0

    tDemFlame002.Interval = 52
    tDemFlame002.Enabled = True
    tDemFlame002Fade.Enabled = True

    tDemFlame002Stop.Interval = 300
    tDemFlame002Stop.Enabled = True
End Sub

Sub StopDemFlame002()
    DemFlame002FadeDir = -1
    tDemFlame002Fade.Enabled = True
End Sub

Sub tDemFlame002_Timer()
    demflame002.ImageA = DemFlameFrames(DemFlame002Pos)
    If DemFlame002Pos < 47 Then
        DemFlame002Pos = DemFlame002Pos + 1
    Else
        tDemFlame002.Enabled = False
    End If
End Sub

Sub tDemFlame002Fade_Timer()
    Dim easeStep
    If DemFlame002FadeDir = 1 Then
        ' Smooth fade in (ease in)
        easeStep = 0.05 * (1 - DemFlame002Opacity)
        DemFlame002Opacity = DemFlame002Opacity + easeStep
        If DemFlame002Opacity >= 0.99 Then
            DemFlame002Opacity = 1
            DemFlame002FadeDir = 0
            tDemFlame002Fade.Enabled = False
        End If
    ElseIf DemFlame002FadeDir = -1 Then
        ' Smooth fade out (ease out)
        easeStep = 0.05 * (DemFlame002Opacity)
        DemFlame002Opacity = DemFlame002Opacity - easeStep
        If DemFlame002Opacity <= 0.01 Then
            DemFlame002Opacity = 0
            DemFlame002FadeDir = 0
            tDemFlame002Fade.Enabled = False
            demflame002.Visible = False
        End If
    End If

    demflame002.IntensityScale = DemFlame002Opacity
End Sub


Sub tDemFlame002Stop_Timer()
    StopDemFlame002
    tDemFlame002Stop.Enabled = False
End Sub

'==========================
' BLIZ ICE (uses Flasher.ImageA)
'==========================
Dim BlizFrame, BlizDir
Dim BlizPulseDir, BlizPulseOn

Sub InitBliz()
    BlizFrame = 1
    BlizDir = 0
    BlizPulseOn = False
    ' set initial image (match exact image names in image manager)
    Bliz.ImageA = "FR_1"
    Bliz.Visible = False
    Bliz.Opacity = 20    ' base opacity (you can change)
End Sub

' main timer drives both build and melt
Sub tBliz_Timer()
    If BlizDir = 1 Then               ' BUILD / FADE IN
        Bliz.Visible = True
        BlizFrame = BlizFrame + 1
        If BlizFrame >= 56 Then
            BlizFrame = 56
            tBliz.Enabled = False    ' stop main timer, hold frozen
            ' start shimmer pulse (optional)
            BlizPulseOn = True
            tBlizPulse.Interval = 100
            tBlizPulse.Enabled = True
        End If
        Bliz.ImageA = "FR_" & BlizFrame

    ElseIf BlizDir = -1 Then          ' MELT / FADE OUT
        ' ensure pulse stops when melting
        If BlizPulseOn Then
            BlizPulseOn = False
            tBlizPulse.Enabled = False
            Bliz.Opacity = 20
        End If

        BlizFrame = BlizFrame + 1     ' note: melting frames are >56 up to 75 (as requested)
        If BlizFrame >= 75 Then
            BlizFrame = 75
            tBliz.Enabled = False
            Bliz.Visible = False      ' fully melted, hide flasher
        End If
        Bliz.ImageA = "FR_" & BlizFrame
    End If
End Sub

' optional shimmer pulse while frozen at FR_56
Sub tBlizPulse_Timer()
    If Not BlizPulseOn Then
        tBlizPulse.Enabled = False
        Exit Sub
    End If

    ' simple up/down opacity pulse between 20 and 30
    If BlizPulseDir = 1 Then
        Bliz.Opacity = Bliz.Opacity + 1
        If Bliz.Opacity >= 30 Then BlizPulseDir = -1
    Else
        Bliz.Opacity = Bliz.Opacity - 1
        If Bliz.Opacity <= 20 Then BlizPulseDir = 1
    End If
End Sub

'===========================
' start / end controls
'===========================
Sub StartBliz()
    BlizDir = 1
    BlizFrame = 1
    Bliz.Visible = True
    Bliz.ImageA = "FR_1"
    tBliz.Interval = 45       ' tweak build speed
    tBliz.Enabled = True
End Sub

Sub EndBliz()
    ' stop main build timer if still running, then start melt from frame 56
    BlizDir = -1
    BlizFrame = 56
    tBliz.Interval = 45       ' tweak melt speed
    tBliz.Enabled = True
End Sub

Sub StopBlizzardMode
	PuPlayer.playevent pDMDVideo,"Misc","BlizzardCollect.mp4",nPupVideoVolume,60,5,0,""
	bBLIZZARDMode = False
	EndSnow
	EndBliz
    EndIceGrow
    Playtaunt2
	DisableBlizzardTargetLights
	RaiseTargetTimer.Enabled = True
	nBlizzTargetCount = 0

End Sub



Sub UpdateBLizzPrepMessages
	PuPlayer.LabelSet pDMD,"BlizzTimerValue",CountdownValue,1,"{'mt':2,'fonth':12.5,'xpos':28.75, 'ypos':46.5}"
	PuPlayer.LabelSet pDMD,"Event3B"," COLLECT ALL ",1,"{'mt':2,'color': " & cBlue &" }"
	PuPlayer.LabelSet pDMD,"Event3Ca",8 - nBlizzTargetCount,1,"{'mt':2,'xpos':" & (50-nDTOffsetX*4) & ", 'color': " & cWhite &"}"
	PuPlayer.LabelSet pDMD,"Event3C"," BLIZZARD TARGETS ",1,"{'mt':2,'color': " & cBlue &" }"
	pDMDLabelSetColorGradient "BlizzTimerValue", cWhite, cBlue
	pDMDLabelSetColorGradient "Event3B", cBlue, cPurple
	pDMDLabelSetColorGradient "Event3C",  cBlue, cPurple
End Sub

Sub HideBlizzTimer
	pDMDlabelHide "BlizzTimerImage"
	PuPlayer.LabelSet pDMD,"BlizzTimerValue","",0,""
	GeneralPupQueue.Add "BackupHideBlizzardTimerImage","BackupHideBlizzardTimerImage",24,1000,0,0,0,False
End Sub

Sub BackupHideBlizzardTimerImage
	if Target009.IsDropped = False Then pDMDlabelHide "BlizzTimerImage"
End Sub

Sub UpdateBLizzMessages
'	PuPlayer.LabelSet pDMD,"Event3A"," BLIZZARD ",1,"{'mt':2,'color': " & cBlue &" }"
	PuPlayer.LabelSet pDMD,"Event3B"," COLLECT JACKPOTS ",1,"{'mt':2,'color': " & cBlue &" }"
	'PuPlayer.LabelSet pDMD,"Event3Ca",8 - nBlizzTargetCount,1,"{'mt':2,'xpos':50, 'color': " & cWhite &"}"
	PuPlayer.LabelSet pDMD,"Event3C"," AT THE RAMPS ",1,"{'mt':2,'color': " & cBlue &" }"
	pDMDLabelSetColorGradient "Event3A", cBlue, cPurple
	pDMDLabelSetColorGradient "Event3B", cBlue, cPurple
	pDMDLabelSetColorGradient "Event3C",  cBlue, cPurple
End Sub

'==========================================================
' SNOW LOOP EFFECT - plays while Blizzard Mode is active
'==========================================================

Dim SnowFrame, SnowPlaying

Sub InitSnow()
    SnowFrame = 1
    SnowPlaying = False
    Snow.ImageA = "SN_1"
    Snow.Visible = False
    Snow.Opacity = 20
End Sub

'===========================
' TIMER
'===========================
Sub tSnow_Timer()
    If Not SnowPlaying Then
        tSnow.Enabled = False
        Exit Sub
    End If

    ' advance frame
    SnowFrame = SnowFrame + 1
    If SnowFrame > 45 Then SnowFrame = 1  ' loop back

    ' set image frame
    Snow.ImageA = "SN_" & SnowFrame
End Sub

'===========================
' CONTROL SUBS
'===========================
Sub StartSnow()
    SnowFrame = 1
    Snow.ImageA = "SN_1"
    Snow.Visible = True
    Snow.Opacity = 25         ' tweak for brightness
    SnowPlaying = True
    tSnow.Interval = 60       ' adjust for loop speed
    tSnow.Enabled = True
End Sub

Sub EndSnow()
    SnowPlaying = False
    tSnow.Enabled = False
    Snow.Visible = False
End Sub

'========================================================
'BALLS ON FIRE
'========================================================
Const FlameFollowerCount = 9
Const FlameFrameStart = 1     ' first usable frame (BFL_1)
Const FlameFrameEnd   = 48    ' last usable frame (BFL_48)
Const FlameZOffset    = 45
Const FlameXOffset    = 0
Const FlameYOffset    = -40

Dim FlameFollowers()
Dim FlameFrames()
Dim FlameAnimPos
Dim FlameActive
Dim TotalFlameFrames
Dim LastX()
Dim LastY()

'--------------------------------------------------------
' Initialize system (call once in Table_Init)
'--------------------------------------------------------
Sub InitFlameFollowers()
    Dim i, frameNum, idx

    ReDim FlameFollowers(FlameFollowerCount)
    ReDim FlameFrames(0)
    ReDim LastX(FlameFollowerCount)
    ReDim LastY(FlameFollowerCount)

    For i = 1 To FlameFollowerCount
        On Error Resume Next
        Set FlameFollowers(i) = Eval("FlameFollower" & i)
        On Error GoTo 0
        If Not FlameFollowers(i) Is Nothing Then
            FlameFollowers(i).Visible = False
            FlameFollowers(i).IntensityScale = 0
            LastX(i) = 0
            LastY(i) = 0
        End If
    Next

    ' Build frame list (BFL_1 → BFL_48)
    idx = 0
    For frameNum = FlameFrameStart To FlameFrameEnd
        ReDim Preserve FlameFrames(idx)
        FlameFrames(idx) = "BFL_" & frameNum
        idx = idx + 1
    Next
    TotalFlameFrames = idx

    ' Ensure editor timer exists and is configured
    tFlameFollow.Interval = 40
    tFlameFollow.Enabled = False

    FlameActive = False
End Sub

'--------------------------------------------------------
' Start / Stop control for flames
'--------------------------------------------------------
Sub StartFlameFollowers()
    FlameActive = True
    FlameAnimPos = 0
    tFlameFollow.Enabled = True
End Sub

Sub StopFlameFollowers()
    Dim i
    FlameActive = False
    tFlameFollow.Enabled = False
    For i = 1 To FlameFollowerCount
        If Not FlameFollowers(i) Is Nothing Then
            FlameFollowers(i).Visible = False
            FlameFollowers(i).IntensityScale = 0
        End If
    Next
End Sub

'--------------------------------------------------------
' Atan2 helper (returns radians)
'--------------------------------------------------------
Function Atan2(y, x)
    Const PI = 3.14159265358979
    If x = 0 Then
        If y > 0 Then
            Atan2 = PI / 2: Exit Function
        ElseIf y < 0 Then
            Atan2 = -PI / 2: Exit Function
        Else
            Atan2 = 0: Exit Function
        End If
    End If
    Atan2 = Atn(y / x)
    If x < 0 Then
        If y >= 0 Then Atan2 = Atan2 + PI Else Atan2 = Atan2 - PI
    End If
End Function

'--------------------------------------------------------
' Timer updates animation, rotation + ball tracking
'--------------------------------------------------------
Sub tFlameFollow_Timer()
    On Error Resume Next
    If Not FlameActive Then Exit Sub

    ' Advance animation frame
    If TotalFlameFrames > 0 Then
        FlameAnimPos = (FlameAnimPos + 1) Mod TotalFlameFrames
    End If

    Dim b, i, dx, dy, angRad, angDeg
    i = 1

    For Each b In GetBalls()
        If Not b Is Nothing Then
            ' only for rolling balls (avoid captive)
            If b.Z < 35 Then
                If i <= FlameFollowerCount Then
                    If Not FlameFollowers(i) Is Nothing Then
                        dx = b.X - LastX(i)
                        dy = b.Y - LastY(i)

                        If Abs(dx) + Abs(dy) > 0.5 Then
                            angRad = Atan2(dy, dx)
                            angDeg = angRad * 57.2957795130823
                            FlameFollowers(i).RotZ = angDeg + 270
                        Else
    ' When ball is nearly still, point flame upward (toward back wall)
    FlameFollowers(i).RotZ = 0 
                       
                        End If

                        LastX(i) = b.X
                        LastY(i) = b.Y

                        ' dynamic trail distance by speed
                        Dim speed, trailDist
                        speed = Sqr((dx * dx) + (dy * dy))
                        trailDist = -75 - (speed * 0.4)

                        ' clamp trail
                        If trailDist < -80 Then trailDist = -80
                        If trailDist > -10 Then trailDist = -10

                        With FlameFollowers(i)
                            .Visible = True
                            .X = b.X + (trailDist * Cos((.RotZ + 90) * 3.1416 / 180))
                            .Y = b.Y + (trailDist * Sin((.RotZ + 90) * 3.1416 / 180))
                            .Z = b.Z + FlameZOffset
                            .ImageA = FlameFrames(FlameAnimPos)
                            .IntensityScale = 1
                        End With
                    End If
                    i = i + 1
                End If
            End If
        End If
    Next

    ' Hide any unused followers
    For i = i To FlameFollowerCount
        If Not FlameFollowers(i) Is Nothing Then FlameFollowers(i).Visible = False
    Next

    On Error GoTo 0
End Sub

Sub RuleCardTrigger_Hit()
	Rules.visible = False
	bGameReady = False	' can no longer add players
End Sub


Sub DbgTracker( myDebugText )
' Uncomment the next line to turn off debugging
Exit Sub

If Not IsObject( objIEDebugWindow ) Then
Set objIEDebugWindow = CreateObject( "InternetExplorer.Application" )
objIEDebugWindow.Navigate "about:blank"
objIEDebugWindow.Visible = True
objIEDebugWindow.ToolBar = False
objIEDebugWindow.Width = 600	
objIEDebugWindow.Height = 900
objIEDebugWindow.Left = 2100
objIEDebugWindow.Top = 100
Do While objIEDebugWindow.Busy
Loop
objIEDebugWindow.Document.Title = "My Debug Window"
objIEDebugWindow.Document.Body.InnerHTML = "<b>Blizzard of Ozz Debug Window -TimeStamp: " & GameTime& "</b></br>"
End If

objIEDebugWindow.Document.Body.InnerHTML = objIEDebugWindow.Document.Body.InnerHTML & myDebugText & "</b><br>" & vbCrLf
End Sub

'====================================================
'   ICE GROWTH EFFECT (RICE001–RICE004)
'   For Blizzard Multiball – updated logic
'====================================================

Dim RiceStep

Sub InitIce()
    RICE001.Visible = False
    RICE002.Visible = False
    RICE003.Visible = False
    RICE004.Visible = False
    RiceStep = 0
    RiceTimer.Enabled = False
    RiceTimer.Interval = 150   ' ~1 sec faster
End Sub


'-----------------------------------------------------
' Fade In (Grow Ice)
'-----------------------------------------------------
Sub StartIceGrow()
    ' Skip during prep flicker — but allow once Blizzard starts
    If bBlizzardPrepMode Then Exit Sub

    ' If not already visible, start the growth
    If Not RiceTimer.Enabled Then
        RiceStep = 1
        RiceTimer.Enabled = True
    End If
End Sub


'-----------------------------------------------------
' Fade Out (Melt Ice)
'-----------------------------------------------------
Sub EndIceGrow()
    ' Only melt when Blizzard is actually done
    If bBLIZZARDMode Or bBlizzardPrepMode Then Exit Sub

    RiceStep = -1
    RiceTimer.Enabled = True
End Sub


'-----------------------------------------------------
' Timer Animation Control
'-----------------------------------------------------
Sub RiceTimer_Timer()
    Select Case RiceStep
        ' ---------------- Fade In (grow)
        Case 1
            RICE001.Visible = True
            RICE002.Visible = False
            RICE003.Visible = False
            RICE004.Visible = False
            RiceStep = 2
        Case 2
            RICE002.Visible = True
            RICE001.Visible = False
            RiceStep = 3
        Case 3
            RICE003.Visible = True
            RICE002.Visible = False
            RiceStep = 4
        Case 4
            RICE004.Visible = True
            RICE003.Visible = False
            RiceStep = 5        ' hold final ice frame
        ' ---------------- Hold final ice look
        Case 5
            RiceTimer.Enabled = False
        ' ---------------- Fade Out (melt)
        Case -1
            RICE003.Visible = True
            RICE004.Visible = False
            RiceStep = -2
        Case -2
            RICE002.Visible = True
            RICE003.Visible = False
            RiceStep = -3
        Case -3
            RICE001.Visible = True
            RICE002.Visible = False
            RiceStep = -4
        Case -4
            RICE001.Visible = False
            RiceTimer.Enabled = False
    End Select
End Sub

'==========================================
' FINAL MODE BATS
'==========================================

Dim BatFrames(35)
Dim BatIdx
Dim bBatActive

'------------------------------------------
' INIT (call once in Table_Init)
'------------------------------------------
Sub InitBatFlaps()
    '--- BAT 1 (bat001-006) ---
    Set BatFrames(0) = bat001
    Set BatFrames(1) = bat002
    Set BatFrames(2) = bat003
    Set BatFrames(3) = bat004
    Set BatFrames(4) = bat005
    Set BatFrames(5) = bat006

    '--- BAT 2 (bat007-012) ---
    Set BatFrames(6) = bat007
    Set BatFrames(7) = bat008
    Set BatFrames(8) = bat009
    Set BatFrames(9) = bat010
    Set BatFrames(10) = bat011
    Set BatFrames(11) = bat012

    '--- BAT 3 (bat013-018) ---
    Set BatFrames(12) = bat013
    Set BatFrames(13) = bat014
    Set BatFrames(14) = bat015
    Set BatFrames(15) = bat016
    Set BatFrames(16) = bat017
    Set BatFrames(17) = bat018

    '--- BAT 4 (bat019-024) ---
    Set BatFrames(18) = bat019
    Set BatFrames(19) = bat020
    Set BatFrames(20) = bat021
    Set BatFrames(21) = bat022
    Set BatFrames(22) = bat023
    Set BatFrames(23) = bat024

    '--- BAT 5 (bat025-030) ---
    Set BatFrames(24) = bat025
    Set BatFrames(25) = bat026
    Set BatFrames(26) = bat027
    Set BatFrames(27) = bat028
    Set BatFrames(28) = bat029
    Set BatFrames(29) = bat030

    '--- BAT 6 (bat031-036) ---
    Set BatFrames(30) = bat031
    Set BatFrames(31) = bat032
    Set BatFrames(32) = bat033
    Set BatFrames(33) = bat034
    Set BatFrames(34) = bat035
    Set BatFrames(35) = bat036

    BatIdx = 0
    bBatActive = False

    BatFlapTimer.Interval = 70   ' Flap speed (lower = faster)
    BatFlapTimer.Enabled = False
End Sub

'------------------------------------------
' START / STOP
'------------------------------------------
Sub StartBats()
    bBatActive = True
    BatFlapTimer.Enabled = True
End Sub

Sub StopBats()
    bBatActive = False
    HideAllBats()
    BatFlapTimer.Enabled = False
End Sub

'------------------------------------------
' TIMER LOOP
'------------------------------------------
Sub BatFlapTimer_Timer()
    If bBatActive Then
        HideAllBats()
        ' Show matching frame for all 6 bat sets
        BatFrames(BatIdx).Visible = True
        BatFrames(BatIdx + 6).Visible = True
        BatFrames(BatIdx + 12).Visible = True
        BatFrames(BatIdx + 18).Visible = True
        BatFrames(BatIdx + 24).Visible = True
        BatFrames(BatIdx + 30).Visible = True
        BatIdx = (BatIdx + 1) Mod 6
    End If
End Sub

'------------------------------------------
' HIDE ALL FRAMES
'------------------------------------------
Sub HideAllBats()
    Dim i
    For i = 0 To 35
        BatFrames(i).Visible = False
    Next
End Sub

Sub SphereTimer_Timer()
    mag1.rotz = mag1.rotz + 1   ' rotates clockwise
    mag3.rotz = mag3.rotz + 1   ' rotates clockwise
    bat001.rotz = bat001.rotz + 6   ' rotates clockwise
    bat002.rotz = bat002.rotz + 6   ' rotates clockwise
    bat003.rotz = bat003.rotz + 6   ' rotates clockwise
    bat004.rotz = bat004.rotz + 6   ' rotates clockwise
    bat005.rotz = bat005.rotz + 6   ' rotates clockwise
    bat006.rotz = bat006.rotz + 6   ' rotates clockwise
    bat007.rotz = bat007.rotz + 12   ' rotates clockwise
    bat008.rotz = bat008.rotz + 12   ' rotates clockwise
    bat009.rotz = bat009.rotz + 12   ' rotates clockwise
    bat010.rotz = bat010.rotz + 12   ' rotates clockwise
    bat011.rotz = bat011.rotz + 12   ' rotates clockwise
    bat012.rotz = bat012.rotz + 12   ' rotates clockwise
    bat013.rotz = bat013.rotz - 9   ' rotates clockwise
    bat014.rotz = bat014.rotz - 9   ' rotates clockwise
    bat015.rotz = bat015.rotz - 9   ' rotates clockwise
    bat016.rotz = bat016.rotz - 9   ' rotates clockwise
    bat017.rotz = bat017.rotz - 9   ' rotates clockwise
    bat018.rotz = bat018.rotz - 9   ' rotates clockwise
    bat019.rotz = bat019.rotz - 8   ' rotates clockwise
    bat020.rotz = bat020.rotz - 8   ' rotates clockwise
    bat021.rotz = bat021.rotz - 8   ' rotates clockwise
    bat022.rotz = bat022.rotz - 8   ' rotates clockwise
    bat023.rotz = bat023.rotz - 8   ' rotates clockwise
    bat024.rotz = bat024.rotz - 8   ' rotates clockwise
    bat025.rotz = bat025.rotz + 6   ' rotates clockwise
    bat026.rotz = bat026.rotz + 6   ' rotates clockwise
    bat027.rotz = bat027.rotz + 6   ' rotates clockwise
    bat028.rotz = bat028.rotz + 6   ' rotates clockwise
    bat029.rotz = bat029.rotz + 6   ' rotates clockwise
    bat030.rotz = bat030.rotz + 6   ' rotates clockwise
    bat031.rotz = bat031.rotz - 9   ' rotates clockwise
    bat032.rotz = bat032.rotz - 9   ' rotates clockwise
    bat033.rotz = bat033.rotz - 9   ' rotates clockwise
    bat034.rotz = bat034.rotz - 9   ' rotates clockwise
    bat035.rotz = bat035.rotz - 9   ' rotates clockwise
    bat036.rotz = bat036.rotz - 9   ' rotates clockwise
End Sub