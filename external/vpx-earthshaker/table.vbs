Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="esha_la3",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

Const BallSize = 50
'Const Ballmass = 1.7

' flex DMD options
Dim UseFlexDMD:UseFlexDMD = False	' True = on, False = off (replacement for external DMD on real DMDs. Inlcludes jackpot values)
									' Set UseFlexDMD to False for VR room to use Alphanumeric display

Dim VR_Room
If RenderingMode = 2 Then VR_Room=True Else VR_Room=False      'VRRoom set based on RenderingMode in version 10.72

LoadVPM "01120100", "S11.VBS", 3.22

NoUpperLeftFlipper
NoUpperRightFlipper
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True And Not VR_Room Then 'Show Desktop components
	Ramp16.visible=1
	Ramp15.visible=1
	Primitive13.visible=1
Else
	Ramp16.visible=0
	Ramp15.visible=0
	Primitive13.visible=0
End if

 '*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(01) = "bsTrough.SolIn"
SolCallback(02) = "bsTrough.SolOut"
SolCallback(03) = "SolDropReset"
SolCallback(04) = "SolFault"
SolCallback(05) = "bsLSaucer.SolOut"
SolCallback(06) = "BotPop"
SolCallback(07) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(09) = "InstituteDrop"
SolCallback(10) = "PFGI2"
'SolCallBack(11) = ""
SolCallback(13) = "VukTopPop"
SolCallBack(14) = "Setlamp 114,"
SolCallBack(15) = "PFGI"
SolCallBack(16) = "Setlamp 116,"
SolCallback(22) = "ShakerMotor"
SolCallBack(25) = "Setlamp 125,"
SolCallBack(26) = "Setlamp 126,"
SolCallBack(27) = "Setlamp 127,"
SolCallBack(28) = "Setlamp 128,"
SolCallBack(29) = "Setlamp 129,"
SolCallBack(30) = "Setlamp 130,"
SolCallBack(31) = "Setlamp 131,"
SolCallBack(32) = "Setlamp 132,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("Flipperup",DOFContactors), LeftFlipper:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAt SoundFX("Flipperdown",DOFContactors), LeftFlipper:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("Flipperup",DOFContactors), RightFlipper:RightFlipper.RotateToEnd
     Else
         PlaySoundAt SoundFX("Flipperdown",DOFContactors), RightFlipper:RightFlipper.RotateToStart
     End If
End Sub


'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolDropReset(Enabled)
	dtL.SolDropUp Enabled
	UpdateTargetShadows
end sub

Sub ShakerMotor(enabled)
	If enabled Then
	    ShakeTimer.Enabled = 1
		playsound SoundFX("Motor",DOFContactors)
	Else
    	ShakeTimer.Enabled = 0
	End If
End Sub

Sub ShakeTimer_Timer()
	Nudge 0,1
	Nudge 90,1
	Nudge 180,1
	Nudge 270,1
End Sub   

'Playfield GI
Sub PFGI(Enabled)
	If Enabled Then
		dim xx
		For each xx in GI:xx.State = 0: Next
		SetLamp 101, 0
        PlaySound "RelayOff"
	Else
		For each xx in GI:xx.State = 1: Next
		SetLamp 101, 1
        PlaySound "RelayOn"
	End If
End Sub

' Upper Playfield GI
Sub PFGI2(Enabled)
	If Enabled Then
		dim xxxx
		For each xxxx in GIU:xxxx.State = 0: Next
		SetLamp 102, 0
        PlaySound "RelayOff"
	Else
		For each xxxx in GIU:xxxx.State = 1: Next
		SetLamp 102, 1
        PlaySound "RelayOn"
	End If
	UpdateTargetShadows
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsLSaucer, dtL

Sub Table1_Init
	PFGI False: PFGI2 False

	' initalise the FlexDMD display
    If UseFlexDMD Then FlexDMD_Init

	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "EARTHSHAKER - WILLIAMS 1989"&chr(13)&"SHAKE N BAKE"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.hidden = 1
        .Games(cGameName).Settings.Value("sound")=1
		If UseFlexDMD Then ExternalEnabled = .Games(cGameName).Settings.Value("showpindmd")
		If UseFlexDMD Then .Games(cGameName).Settings.Value("showpindmd") = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0
 
	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch  = 9
	vpmNudge.Sensitivity = 2
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

	Set bsTrough = New cvpmBallStack
		bsTrough.Initsw 10,11,12,13,0,0,0,0
		bsTrough.InitKick BallRelease,150,7
        bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("SolOn",DOFContactors)
		bsTrough.Balls = 3 

	Set bsLSaucer = New cvpmBallStack
		bsLSaucer.InitSaucer sw20, 20, 250, 10
		bsLSaucer.InitExitSnd SoundFX("",DOFContactors), SoundFX("SolOn",DOFContactors)

	Set dtL=New cvpmDropTarget
		dtL.InitDrop Array(sw27,sw28,sw29),Array(27,28,29)
		dtL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	HRampDX.Collidable=0
 
  ' Create Captive Ball
	CapKicker.CreateSizedBallWithMass BSize,BMass
	CapKicker.Kick 0,5

	'VR stuff
	Dim objVR
	If VR_Room Then
		GIon_Flash1.visible = False
		GIon_Flash2.visible = False
		setup_backglass	
		if UseFlexDMD then
			For Each objVR in VRAlphanumericLeds
				objVR.visible = False
			Next
			primary_grill.image = "VR_Grill_FlexDMD"
			vr_dmd.visible = True
		Else
			For Each objVR in VRAlphanumericLeds
					objVR.visible = True
			Next
			Primary_grill.image = "VR_Grill"
			vr_dmd.visible = False
		End if
		
	Else
		'Hide all for non VR
		For Each objVR in VRStuff
				objVR.visible = false
		Next

	End If
 End Sub


'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

	If keycode = PlungerKey Then Plunger.Pullback::PlaySoundAt "plungerpull", Plunger : TimerVRPlunger.enabled = true :TimerVRPlunger2.enabled = False
	If keycode = LeftFlipperKey Then Controller.Switch(58)  = 1 : VRFlipperButtonLeft.X = VRFlipperButtonLeft.X +8 
	If keycode = RightFlipperKey Then Controller.Switch(57) = 1 : VRFlipperButtonRight.X = VRFlipperButtonRight.X -8
	If keycode = StartGameKey then Primary_StartButton.y = Primary_StartButton.y -6
	If KeyDownHandler(keycode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal KeyCode)

	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger", Plunger : TimerVRPlunger.enabled = False : TimerVRPlunger2.enabled = True
	If keycode = LeftFlipperKey Then Controller.Switch(58)  = 0 : VRFlipperButtonLeft.X = VRFlipperButtonLeft.X -8
	If keycode = RightFlipperKey Then Controller.Switch(57) = 0 :VRFlipperButtonRight.X = VRFlipperButtonRight.X +8
If keycode = StartGameKey then Primary_StartButton.y = Primary_StartButton.y +6
	If KeyUpHandler(keycode) Then Exit Sub
End Sub

Sub UpdateTargetShadows
	' If the upper GI is off, shadows should be off.
	if LampState(102) = 0 then
		SetLamp 105, 0
		SetLamp 106, 0
		SetLamp 107, 0
	Else
		SetLamp 105, 1+Controller.Switch(27)
		SetLamp 106, 1+Controller.Switch(28)
		SetLamp 107, 1+Controller.Switch(29)
	end if
end sub 
	
			

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : PlaySoundAt "DrainLong", Drain : End Sub
Sub sw20_Hit:bsLSaucer.addball 0 : PlaySoundAtVol "KickerEnter", sw20, 0.25 : End Sub


'Stand Up Targets
Sub sw19_hit:vpmTimer.pulseSw 19 : End Sub
Sub sw21_hit:vpmTimer.pulseSw 21 : End Sub
Sub sw22_hit:vpmTimer.pulseSw 22 : End Sub
Sub sw23_hit:vpmTimer.pulseSw 23 : End Sub
Sub sw24_hit:vpmTimer.pulseSw 24 : End Sub
Sub sw30_hit:vpmTimer.pulseSw 30 : End Sub

'Drop Targets
 Sub Sw27_Dropped:dtL.Hit 1 :UpdateTargetShadows:End Sub  
 Sub Sw28_Dropped:dtL.Hit 2 :UpdateTargetShadows:End Sub  
 Sub Sw29_Dropped:dtL.Hit 3 :UpdateTargetShadows:End Sub

'Wire Triggers
Sub sw14_Hit:Controller.Switch(14)=1 : PlaySoundAtBall "Sensor" : End Sub
Sub sw14_UnHit:Controller.Switch(14)=0:End Sub
Sub sw15_Hit:Controller.Switch(15)=1 : PlaySoundAtBall "Sensor" : End Sub
Sub sw15_UnHit:Controller.Switch(15)=0:End Sub
Sub sw16_Hit:Controller.Switch(16)=1 : PlaySoundAtBall "Sensor" : End Sub
Sub sw16_UnHit:Controller.Switch(16)=0:End Sub
Sub sw17_Hit:Controller.Switch(17)=1 : PlaySoundAtBall "Sensor" : End Sub
Sub sw17_UnHit:Controller.Switch(17)=0:End Sub
Sub sw18_Hit:Controller.Switch(18)=1 : PlaySoundAtBall "Sensor" : End Sub
Sub sw18_UnHit:Controller.Switch(18)=0:End Sub
Sub sw33_Hit:controller.switch (33)=1 : PlaySoundAtBall "Sensor" : End Sub  
Sub sw33_unHit:controller.switch (33)=0:End Sub
Sub sw34_Hit:controller.switch (34)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw34_unHit:controller.switch (34)=0:End Sub
Sub sw35_Hit:controller.switch (35)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw35_unHit:controller.switch (35)=0:End Sub
Sub sw36_Hit:controller.switch (36)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw36_unHit:controller.switch (36)=0:End Sub

'Gate Trigger
Sub sw31_hit:vpmTimer.pulseSw 31 : End Sub
Sub sw32_hit:vpmTimer.pulseSw 32 : End Sub
Sub sw43_hit:vpmTimer.pulseSw 43 : End Sub
Sub sw44_hit:vpmTimer.pulseSw 44 : End Sub

'Subway
Sub sw38_Hit:Controller.Switch(38)=1 : playsound"Subway" : End Sub
Sub sw38_UnHit:Controller.Switch(38)=0:End Sub
Sub sw39_Hit:Controller.Switch(39)=1 : End Sub 
Sub sw39_unHit:Controller.Switch(39)=0:End Sub


'Spinners
Sub sw41_Spin:vpmTimer.PulseSw 41 : playsound"fx_spinner" : End Sub

'Left Ramp Wire Triggers
Sub sw45_Hit: controller.switch (45)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw45_unHit: controller.switch (45)=0:End Sub
Sub sw46_Hit: controller.switch (46)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw46_unHit: controller.switch (46)=0:End Sub

' Shooter lane
Sub sw50_Hit: controller.switch (50)=1 : End Sub 
Sub sw50_unHit: controller.switch (50)=0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(52) : PlaySoundAt SoundFX("bumper1",DOFContactors), Bumper1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(53) : PlaySoundAt SoundFX("bumper2",DOFContactors), Bumper2: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(54) : PlaySoundAt SoundFX("bumper3",DOFContactors), Bumper3: End Sub

'Generic Sounds
Sub Trigger1_Hit : PlaySound "fx_ballrampdrop": End Sub
Sub Trigger2_Hit : StopSound "Wire Ramp" : PlaySound "fx_ballrampdrop": End Sub
Sub Trigger3_Hit : StopSound "Wire Ramp" : End Sub
Sub Trigger4_Hit : PlaySound "Wire Ramp": End Sub
Sub Trigger5_Hit : StopSound "Wire Ramp" : PlaySound "popper_ball": End Sub
'Sub Trigger6_Hit : PlaySound "popper_ball": End Sub

Sub PlungerCradle_Hit
	'Reset ball spin parameters, they don't "dissipate" at rest and they causes inconsistent shooting behavior
	ActiveBall.AngMomX = 0
	ActiveBall.AngMomY = 0
	ActiveBall.AngMomZ = 0
	ActiveBall.AngVelX = 0
	ActiveBall.AngVelY = 0
	ActiveBall.AngVelZ = 0
end sub

 '***********************************
 'Top Raising VUK 
 '***********************************
 'Variables used for VUK 
 Dim raiseballsw, raiseball 

 Sub TopVUK_Hit() 
 	TopVUK.Enabled=FALSE
	Controller.switch (37) = True
	PlaySoundAtVol "popper_ball", TopVUK, 0.25
 End Sub
 
 Sub VukTopPop(enabled)
	if(enabled and Controller.switch (37)) then
		PlaySoundAt SoundFX("Popper_ball",DOFContactors), TopVUK
		TopVUK.DestroyBall
 		Set raiseball = TopVUK.CreateBall
 		raiseballsw = True
 		TopVukraiseballtimer.Enabled = True 'Added by Rascal
		TopVUK.Enabled=TRUE	
 		Controller.switch (37) = False
	else
		'PlaySound "Popper"
	end if
End Sub
 
 Sub TopVukraiseballtimer_Timer()
 	If raiseballsw = True then
 		raiseball.z = raiseball.z + 10
 		If raiseball.z > 190 then
 			'msgbox ("Over")
 			TopVUK.Kick 270, 10 
 			Set raiseball = Nothing
 			TopVukraiseballtimer.Enabled = False
 			raiseballsw = False
			PlaySound "Wire Ramp"
 		End If
 	End If
 End Sub
 
 '********************
 ' Bottom raising VUK 
 '********************
 Dim braiseballsw, braiseball 

 Sub BottomVuk_Hit() 
	playsound"popper_ball"
 	BottomVUK.Enabled=FALSE
	Controller.switch (40) = True
 End Sub

 Sub BotPop(enabled)
	if(enabled and Controller.switch (40)) then
		playsound SoundFX("Popper",DOFContactors)
		BottomVUK.DestroyBall
 		Set braiseball = BottomVUK.CreateBall
 		braiseballsw = True
 		BottomVukraiseballtimer.Enabled = True 'Added by Rascal
		BottomVUK.Enabled=TRUE	
 		Controller.switch (40) = False
	else
		'PlaySound "Popper"
	end if
End Sub

 Sub BottomVukraiseballtimer_Timer()
 	If braiseballsw = True then
 		braiseball.z = braiseball.z + 10
 		If braiseball.z > 90 then
 			'msgbox ("Over")
 			BottomVUK.Kick 210, 9
 			Set braiseball = Nothing
 			BottomVukraiseballtimer.Enabled = False
 			braiseballsw = False
			PlaySound "Wire Ramp"
 		End If
 	End If
 End Sub


 '*************************************************
 ' California Nevada DIVERTER
 '*************************************************
dim FaultOpen, FaultChange : FaultOpen = False : FaultChange = False

Sub SolFault(enabled) 
	If enabled Then
		FaultOpen = Not FaultOpen
		If FaultOpen Then 
 			OpenFaultTimer.enabled=1
			HRampSX.Collidable= 0
			HRampDX.Collidable= 1
			FaultChange = 1
		Else
			CloseFaultTimer.enabled=1
 			HRampSX.Collidable= 1
			HRampDX.Collidable= 0 
			FaultChange = 1
		End If
	End If
End Sub

Dim BridgeStep: BridgeStep = 0

Sub OpenFaultTimer_Timer
    Select Case BridgeStep
		Case 0: CalPrim.TransX  =0 : NevadaPrim.TransX  =0   : BridgeStep =1
		Case 1: CalPrim.TransX  =-3: NevadaPrim.TransX  =3   : BridgeStep =2
		Case 2: CalPrim.TransX  =-6: NevadaPrim.TransX  =6   : BridgeStep =3
		Case 3: CalPrim.TransX  =-11: NevadaPrim.TransX  =11 : BridgeStep =4
        Case 4: CalPrim.TransX  =-16: NevadaPrim.TransX  =16 : BridgeStep =5
        Case 5: CalPrim.TransX  =-23: NevadaPrim.TransX  =26 : BridgeStep =0 : controller.switch(42) = 1 : OpenFaultTimer.enabled = 0
    End Select
End Sub

Sub CloseFaultTimer_Timer
    Select Case BridgeStep
		Case 0: CalPrim.TransX  =-23: NevadaPrim.TransX  =26 : BridgeStep =1
        Case 1: CalPrim.TransX  =-17: NevadaPrim.TransX  =17 : BridgeStep =2
		Case 2: CalPrim.TransX  =-11: NevadaPrim.TransX  =11 : BridgeStep =3
		Case 3: CalPrim.TransX  =-6:  NevadaPrim.TransX  =6  : BridgeStep =4
        Case 4: CalPrim.TransX  =-3:  NevadaPrim.TransX  =3  : BridgeStep =5
        Case 5: CalPrim.TransX  =0:   NevadaPrim.TransX  =-0 : BridgeStep =0 : controller.switch(42) = 0 : CloseFaultTimer.enabled = 0
    End Select
End Sub


'**************************************
'
' Earthquake Institute Animation
'
'**************************************

' In the real production Earthshakers, the Institute building is just a
' static toy.  All it does is flash its lights.  But in the pre-production 
' prototypes, the building had a sliding mechanism and a motor that made it 
' sink into the playfield and rise back up on certain game events.  The
' ROM code that controls the building movement is still present in the
' production ROMs, so some Earthshaker owners have modded their tables to
' restore the moving building.  This code implements the building animation
' in our simulated version.
'
' The Institute building is implemented using a bunch of little primitive 
' objects: one for each light window, and one for the outer frame with the 
' back wall and top sign.  This array is a list of all of these primitve 
' components.  To animate the building, we simply move all of the
' primitives in unison.  The animation is extremely simple in that all the
' building does is move straight up and down.  
'
' Windows are at the start of the array to make it easy to find the object
' for a given window number.  Note that array indices start' at 0, but window 
' numbers start at 1 - so InstPrim(0) is window 1, etc.
Dim InstPrim : InstPrim = Array( _
	InstituteWindow1, _
	InstituteWindow2, _
	InstituteWindow3, _
	InstituteWindow4, _
	InstituteWindow5, _
	InstituteWindow6, _
	InstituteWindow7, _
	InstituteWindow8, _
	InstituteWindow9, _
    Primitive034, _
	InstituteBackWall)

' Building travel parameters.  The range of motion is about 3/4 of the building
' height, and it takes about 3 seconds to cover that distance.
dim InstZ, InstDZ, InstDZUp, InstDZDown, InstZMin, InstZMax, InstZRange, InstTravelTime, InstMotorOn, InstZone
InstZMin = -150 : InstZMax = 0   ' travel limits, as Z Translations for the building primitives
InstTravelTime = 3000            ' one-way travel time in milliseconds
InstMotorOn = 0                  ' start with motor turned off

' figure the distance per timer events (total distance divided by total time,
' multiplied by the time interval per event)
InstZRange = InstZMax - InstZMin
InstDZUp = InstZRange / InstTravelTime * InstituteTimer.Interval
InstDZDown = -InstDZUp           ' use the same speed up and down

' start fully up (set Z to max, zone to top (zone 0=bottom, 1=between top and bottom, 2=top))
InstZ = InstZMax
InstZone = 2

' Institute light states.  Each array entry corresponds to the window
' primitive at the same index in InstPrim().
'
' InstLamp(i) is the ROM lamp number for the nth building light
' InstLampBri is the current brightness: 0=off, 1=33%, 2=66%, 3=100%.
' On each timer event, we increase or decrease the brightness of a light
' according to its current on/off state.
Dim InstLamp : InstLamp = Array(23, 24, 25, 20, 21, 22, 17, 18, 19)
Dim InstLampBri : InstLampBri = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)

Sub InstituteTimer_Timer
	' Handle the motor
	if InstMotorOn then
		' move up or down in the current direction
		InstZ = InstZ + InstDZ

		' check for limits
		if InstZ >= InstZMax then 
			' all the the way up - stop here and reverse directions
			InstZ = InstZMax
			InstDZ = InstDZDown

			' trip switches to tell the ROM we're in the UP position
			controller.switch(25) = false
			controller.switch(26) = true
			InstZone = 3
		elseif InstZ <= InstZMin then
			' all the way down - stop here and reverse directions
			InstZ = InstZMin
			InstDZ = InstDZUp

			' trip switches to tell the ROM we're in the DOWN position
			controller.switch(25) = true
			controller.switch(26) = true
			InstZone = 0
		elseif InstZone = 3 and InstZ < InstZMax - InstZRange/3 then
			' we were at the top, and now we're 1/3 of the way down - the
			' original script set both switches off at this position
			controller.switch(25) = false
			controller.switch(26) = false
			InstZone = 2
		elseif InstZone = 2 and InstZ < InstZMin + InstZRange/3 then
			' crossing 2/3 of the way down - the original script set 25 ON
			' and 26 OFF at this position
			controller.switch(25) = true
			controller.switch(26) = false
		end if
		For i = 0 to UBound(InstPrim)
			InstPrim(i).TransZ = InstZ
		next
	end if

    ' Handle the lights
	Dim i, b
	for i = 0 to UBound(InstLamp)
		b = InstLampBri(i)
		if LampState(InstLamp(i)) = 5 then  ' controller turned lamp ON
			if b < 3 then
				b = b + 1
				InstLampBri(i) = b 
				InstPrim(i).Image = "Institute Map Brightness " & b
			end if
		else
			if b > 0 then
				b = b - 1
				InstLampBri(i) = b
				InstPrim(i).Image = "Institute Map Brightness " & b
			end if
		end if
	next
End Sub


' ROM motor control interface 
sub InstituteDrop(enabled)
	InstMotorOn = enabled
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 17 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps()
	NFadeL 1, l1
 	NFadeL 2, l2
 	NFadeL 3, l3
 	NFadeL 4, l4
 	NFadeL 5, l5
 	NFadeL 6, l6
 	NFadeL 7, l7
 	NFadeL 8, l8
 	NFadeL 9, l9
 	NFadeL 10, l10
 	NFadeL 11, l11
	If VR_Room Then FlashVR 12, FL_S12
 	NFadeL 12, l12
 	NFadeL 13, l13
 	NFadeL 14, l14
 	NFadeL 15, l15
 	NFadeL 16, l16
	FadeDisableLighting 17, InstituteWindow7, 1
	FadeDisableLighting 18, InstituteWindow8, 1
	FadeDisableLighting 19, InstituteWindow9, 1
	FadeDisableLighting 20, InstituteWindow4, 1
	FadeDisableLighting 21, InstituteWindow5, 1
	FadeDisableLighting 22, InstituteWindow6, 1
	FadeDisableLighting 23, InstituteWindow1, 1
	FadeDisableLighting 24, InstituteWindow2, 1
	FadeDisableLighting 25, InstituteWindow3, 1
 	if VR_Room Then
		FlashVR 26, FL_S26
		FlashVR 27, FL_S27
		FlashVR 28, FL_S28
		NFadeL 29, l29
		FlashVR 30, FL_S30
		FlashVR 31, FL_S31
		FlashVR 32, FL_S32
	Else
		NFadeL 26, l26
		NFadeL 27, l27
		NFadeL 28, l28
		NFadeL 29, l29
		NFadeL 30, l30
		NFadeL 31, l31
		NFadeL 32, l32
	end if
 	NFadeL 33, l33
 	NFadeL 34, l34
 	NFadeL 35, l35
 	NFadeL 36, l36
 	NFadeL 37, l37
 	NFadeL 38, l38
 	NFadeL 39, l39
	NFadeL 40, l40
    NFadeLm 41, l41 'Bumper 1
    NFadeL 41, l41a
    NFadeLm 42, l42 'Bumper 2
    NFadeL 42, l42a
    NFadeLm 43, l43 'Bumper 3
    NFadeL 43, l43a
    NFadeL 44, l44
 	NFadeL 45, l45
 	NFadeL 46, l46
 	NFadeL 47, l47
 	NFadeL 48, l48
    NFadeObjm 49, l49, "bulbcover1_redOn", "bulbcover1_red"       'Ramp Entrence Red    LED
 	NFadeL 49, l49a
 	NFadeL 50, l50
 	NFadeL 51, l51
 	NFadeL 52, l52	
 	NFadeL 53, l53
 	NFadeL 54, l54
 	NFadeL 55, l55
 	NFadeL 56, l56
    NFadeObjm 57, l57, "bulbcover1_redOn", "bulbcover1_red"       'Ramp Entrence Red    LED
    NFadeL 57, l57a

If DesktopMode = True And Not VR_Room Then 'Show Desktop components
	NFadeL 58, l58 'Backglass
	NFadeL 59, l59 'Backglass
	NFadeL 60, l60 'Backglass
	NFadeL 61, l61 'Backglass
	NFadeL 62, l62 'Backglass
	NFadeL 63, l63 'Backglass
	NFadeL 64, l64 'Backglass
ElseIf VR_Room and not USeFlexDMD Then 'Show in VR if FlexDMD is not used
	flashVR 58, FL_58	
	flashVR 59, FL_59
	flashVR 60, FL_60	
	flashVR 61, FL_61
	flashVR 62, FL_62	
	flashVR 63, FL_63	
	flashVR 64, FL_64	
End if

	' GI
	IF VR_room Then
		Flash 101, vr_gion_flash1
		Flash 102, vr_gion_flash2
	Else
		Flash 101, gion_flash1
		Flash 102, gion_flash2
	end If
	' Drop target shadows
	Flash 105, drop1
	Flash 106, drop2
	Flash 107, drop3

 'Solenoid Controlled Flashers

    NFadeLm 114, f114
    NFadeL 114, f114a
NFadeLm 116, f116
    NFadeL 116, f116a
NFadeLm 125, f125
    NFadeL 125, f125a
NFadeLm 126, f126
    NFadeL 126, f126a
NFadeLm 127, f127
    NFadeL 127, f127a
NFadeLm 128, f128
    NFadeL 128, f128a
NFadeLm 129, f129
    NFadeL 129, f129a
NFadeLm 130, f130
    NFadeL 130, f130a
NFadeLm 131, f131
    NFadeL 131, f131a
NFadeLm 132, f132
    NFadeLm 132, f132a
NFadeLm 132, f132c
    NFadeL 132, f132d

End Sub


' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

Sub FadeDisableLighting(nr, a, alvl)
	Select Case FadingLevel(nr)
		Case 4
			a.UserValue = a.UserValue - 0.5
			If a.UserValue < 0 Then 
				a.UserValue = 0
				FadingLevel(nr) = 0
			end If
			a.BlendDisableLighting = alvl * a.UserValue 'brightness
		Case 5
			a.UserValue = a.UserValue + 0.50
			If a.UserValue > 1 Then 
				a.UserValue = 1
				FadingLevel(nr) = 1
			end If
			a.BlendDisableLighting = alvl * a.UserValue 'brightness
	End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                  'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
				Object.visible = false
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
			Object.visible = true
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

Sub FlashVR(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
				Object.visible = False
        Case 5 ' on
			Object.visible = True
    End Select
End Sub

 
'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(32)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
 Digits(14)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
 Digits(15)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
 
 Digits(16)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
 Digits(17)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
 Digits(18)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
 Digits(19)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
 Digits(20)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
 Digits(21)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
 Digits(22)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
 Digits(23)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
 Digits(24)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
 Digits(25)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
 Digits(26)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
 Digits(27)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
 Digits(28)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
 Digits(29)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
 Digits(30)=Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
 Digits(31)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)
 
 Sub DisplayTimer_Timer

    Dim ChgLED, ii, num, chg, stat, obj
	If UseFlexDMD then FlexDMD.LockRenderThread
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
		If DesktopMode = True Or UseFlexDMD Or VR_Room Then
			For ii=0 To UBound(chgLED)
				num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
				If UseFlexDMD then UpdateFlexChar num, stat					'flexdmd update scores
				If DesktopMode = True And UseFlexDMD = False And Not VR_Room Then
					if (num < 32) then
						For Each obj In Digits(num)
							If chg And 1 Then obj.State=stat And 1
							chg=chg\2 : stat=stat\2
						Next
					Else
					end if
				Elseif VR_Room and Not UseflexDMD  then
					'VR display without FlexDMD
					If (num < 32) then
						For Each obj In VRDigits(num)
							If chg And 1 Then FadeDisplay obj, stat And 1 ,num	
							chg=chg\2 : stat=stat\2
						Next
					End If
			
				End If
			Next
	   end if
    End If

	If UseFlexDMD then
		'flex dmd lamp/solenoid code
		With FlexDMDScene
			'jackpot indicators
			.GetImage("JP58").Visible = Controller.Lamp(58)
			.GetImage("JP59").Visible = Controller.Lamp(59)
			.GetImage("JP60").Visible = Controller.Lamp(60)
			.GetImage("JP61").Visible = Controller.Lamp(61)
			.GetImage("JP62").Visible = Controller.Lamp(62)
			.GetImage("JP63").Visible = Controller.Lamp(63)
			.GetImage("JP64").Visible = Controller.Lamp(64)

			'shaker show/hide noise
			If Controller.Solenoid(22) Then
				.GetImage("Shaker").Visible = True
				Select Case Int(Rnd*5)+1
				Case 1
					.GetImage("Shaker").Bitmap = FlexDMD.NewImage("Shaker","VPX.DMD_Noise1").Bitmap
				Case 2
					.GetImage("Shaker").Bitmap = FlexDMD.NewImage("Shaker","VPX.DMD_Noise2").Bitmap
				Case 3
					.GetImage("Shaker").Bitmap = FlexDMD.NewImage("Shaker","VPX.DMD_Noise3").Bitmap
				Case 4
					.GetImage("Shaker").Bitmap = FlexDMD.NewImage("Shaker","VPX.DMD_Noise4").Bitmap
				Case 5
					.GetImage("Shaker").Bitmap = FlexDMD.NewImage("Shaker","VPX.DMD_Noise5").Bitmap
				End Select
			Else
				.GetImage("Shaker").Visible = False
			End If

		End With

		FlexDMD.UnlockRenderThread

	End If
 End Sub


'**********************************************************************************************************
'**********************************************************************************************************
'	Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 56
    PlaySoundAt SoundFX("SlingshotRight",DOFContactors), SLING1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 55
    PlaySoundAt SoundFX("SlingshotLeft",DOFContactors), SLING2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

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

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
		PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set a Looping sound at object.
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
	PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
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
    Vol = Csng(BallVel(ball) ^2 / 400)
End Function

Function RollVol(ball) ' Calculates the Volume of the sound based on the ball speed
    RollVol = Csng(BallVel(ball) ^2 / 3000)
End Function


Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function RndNum(min,max)
 RndNum = Int(Rnd()*(max-min+1))+min     ' Sets a random number between min and max
End Function



'===========================
' Ramp Skill Shot
'===========================

Sub Skill33a_Hit() 
Dim Velx,Vely
Velx = ActiveBall.Velx : Vely = ActiveBall.Vely
 'msgbox("33"  & Velx & "-" & Vely)
If Velx < 0 Then 
	Skill33a.DestroyBall : PlaySoundAt "BallBounce", Skill35b
	Skill33b.CreateBall : Skill33b.Kick 180,1
Else
	Skill33a.Kick 90,Velx
End If 
End Sub

' Handle Skill Shot Hole 25k
Sub Skill34a_Hit() 
Dim Velx,Vely
Velx = ActiveBall.Velx : Vely = ActiveBall.Vely
 'msgbox("34" & Velx & "-" & Vely)
If Vely > 0 Then 
	Skill34a.DestroyBall : PlaySoundAt "BallBounce", Skill35b
	Skill34b.CreateBall : Skill34b.Kick 180,1
Else
	Skill34a.Kick 180,Vely
End If 
End Sub

' Handle Skill Shot Hole 100k
Sub Skill35a_Hit() 
Dim Velx,Vely
Velx = ActiveBall.Velx : Vely = ActiveBall.Vely
 'msgbox("35" & Velx & "-" & Vely)

debug.print "VelX " & velx
' We can't rely on the ball rolling backwards into this kicker, so we need to trip if it's moving slow enough to the left 

If Velx >= -4 Then 
	Skill35a.DestroyBall : PlaySoundAt "BallBounce", Skill35b
	Skill35b.CreateBall : Skill35b.Kick 180,1
Else
	Skill35a.Kick 90,Velx
End If 
End Sub

'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )/7, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )/12, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
 ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	FlipperLSh1.RotZ = LeftFlipper1.currentangle
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
		BallShadow(b).X = BOT(b).X
		ballShadow(b).Y = BOT(b).Y + 10                       
        If BOT(b).Z > 20 and BOT(b).Z < 200 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
if BOT(b).z > 30 Then 
ballShadow(b).height = BOT(b).Z - 20
ballShadow(b).opacity = 90
Else
ballShadow(b).height = BOT(b).Z - 24
ballShadow(b).opacity = 80
End If
    Next	
End Sub





Sub Pins_Hit (idx)
	PlaySound "fx_PinHit", 0, Vol(ActiveBall)/4, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Switches_Hit (idx)
	PlaySound "Sensor", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)/2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "GateWire", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)/4, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)/4, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)/4, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)/4, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)/4, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub RandomSoundHole()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "Hole1", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "Hole2", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "Hole4", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub LeftFlipper1_Collide(parm)
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


'************************RAMP STUFF************************
'Set invisible Trigger on surface of ramp. Add these commands for each one.


Sub Trigger6_Hit:PlaySound "popper_ball": PlaySound "MetalRolling", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall): End Sub

Sub LRD_Hit()
PlaySoundAt "BallBounce", LRD
End Sub

Sub RRD_Hit()
PlaySoundAt "BallBounce", RRD
End Sub

Sub PRD_Hit: PlaySoundAt "RampDrop", PRD: End Sub
Sub PRD1_Hit: PlaySoundAt "RampDrop", PRD1: End Sub

'RAMP BUMPS

Sub RampBumps_Hit (idx)
	RandomSoundRamps
End Sub

Sub RandomSoundRamps()
	Select Case Int(Rnd*7)+1
		Case 1: Playsound "RB1", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2: Playsound "RB2", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3: Playsound "RB3", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 4: Playsound "RB4", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 5: Playsound "RB5", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 6: Playsound "RB6", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 7: Playsound "RB7", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub




Sub Table1_Exit()
	Controller.Pause = False
	Controller.Stop
	If UseFlexDMD then
		If Not FlexDMD is Nothing Then 
			FlexDMD.Show = False
			FlexDMD.Run = False
			FlexDMD = NULL
		End if
		Controller.Games(cGameName).Settings.Value("showpindmd") = ExternalEnabled
	End if
End Sub

'**********************************************************************************************************
' FlexDMD code - scutters
'**********************************************************************************************************
Dim FlexDMD
DIm FlexDMDDict
Dim FlexDMDScene
Dim ExternalEnabled

Sub FlexDMD_Init() 'default/startup values

	'setup flex dmd
	
	Dim i

	' populate the lookup dictionary for mapping display characters
	FlexDictionary_Init

	Set FlexDMD = CreateObject("FlexDMD.FlexDMD")

	If Not FlexDMD is Nothing Then
	
		FlexDMD.GameName = cGameName
 		FlexDMD.TableFile = Table1.Filename & ".vpx"
		FlexDMD.RenderMode = 2
		FlexDMD.Width = 128
		FlexDMD.Height = 32
		FlexDMD.Clear = True
		FlexDMD.Run = True

		FlexDMD.LockRenderThread

		Set FlexDMDScene = FlexDMD.NewGroup("Scene")
		
		With FlexDMDScene
			'populate blank display
			.AddActor FlexDMD.NewImage("BackG", "FlexDMD.Resources.dmds.black.png")
			.AddActor FlexDMD.NewImage("BackC", "VPX.DMD_BackColour")
			.GetImage("BackC").SetAlignedPosition 0,1,0

			.AddActor FlexDMD.NewImage("Frame", "VPX.DMD_Frame")
			'32 segment display holders
			for i = 0 to 15
				'first line
				.AddActor FlexDMD.NewImage("Seg" & i, "VPX.DMD_Space")
				.GetImage("Seg" & i).SetAlignedPosition i * 8,1,0
				'second line
				.AddActor FlexDMD.NewImage("Seg" & i+16, "VPX.DMD_Space")
				.GetImage("Seg" & i+16).SetAlignedPosition i * 8,10,0
			next

			'jackpot indicators (by lamp numbers)
			.AddActor FlexDMD.NewImage("JP58", "VPX.DMD_JP_L58")
			.GetImage("JP58").SetAlignedPosition 0,19,0
			.GetImage("JP58").Visible = False
			.AddActor FlexDMD.NewImage("JP59", "VPX.DMD_JP_L59")
			.GetImage("JP59").SetAlignedPosition 19,19,0
			.GetImage("JP59").Visible = False
			.AddActor FlexDMD.NewImage("JP60", "VPX.DMD_JP_L60")
			.GetImage("JP60").SetAlignedPosition 37,19,0
			.GetImage("JP60").Visible = False
			.AddActor FlexDMD.NewImage("JP61", "VPX.DMD_JP_L61")
			.GetImage("JP61").SetAlignedPosition 56,19,0
			.GetImage("JP61").Visible = False
			.AddActor FlexDMD.NewImage("JP62", "VPX.DMD_JP_L62")
			.GetImage("JP62").SetAlignedPosition 74,19,0
			.GetImage("JP62").Visible = False
			.AddActor FlexDMD.NewImage("JP63", "VPX.DMD_JP_L63")
			.GetImage("JP63").SetAlignedPosition 92,19,0
			.GetImage("JP63").Visible = False
			.AddActor FlexDMD.NewImage("JP64", "VPX.DMD_JP_L64")
			.GetImage("JP64").SetAlignedPosition 110,19,0
			.GetImage("JP64").Visible = False

			'shaker 'noise' overlay
			.AddActor FlexDMD.NewImage("Shaker", "VPX.DMD_Noise1")
			.GetImage("Shaker").Visible = False

		End With


		FlexDMD.Stage.AddActor FlexDMDScene
		
		FlexDMD.Show = True
		FlexDMD.UnlockRenderThread


	Else
		
		UseFlexDMD = False

	End If

End Sub

Sub FlexDictionary_Init

	Set FlexDMDDict = CreateObject("Scripting.Dictionary")

	FlexDMDDict.Add 0, "VPX.DMD_Space"
	FlexDMDDict.Add 63, "VPX.DMD_O"
	FlexDMDDict.Add 6, "VPX.DMD_1"
	FlexDMDDict.Add 2139, "VPX.DMD_2"
	FlexDMDDict.Add 2127, "VPX.DMD_3"
	FlexDMDDict.Add 2150, "VPX.DMD_4"
	FlexDMDDict.Add 2157, "VPX.DMD_S"
	FlexDMDDict.Add 2173, "VPX.DMD_6"
	FlexDMDDict.Add 7, "VPX.DMD_7"
	FlexDMDDict.Add 2175,"VPX.DMD_8"
	FlexDMDDict.Add 2159,"VPX.DMD_9"
	
	FlexDMDDict.Add 32959,"VPX.DMD_Oc"
	FlexDMDDict.Add 32902, "VPX.DMD_1c"
	FlexDMDDict.Add 35035, "VPX.DMD_2c"
	FlexDMDDict.Add 35023, "VPX.DMD_3c"
	FlexDMDDict.Add 35046, "VPX.DMD_4c"
	FlexDMDDict.Add 35053, "VPX.DMD_Sc"
	FlexDMDDict.Add 35069, "VPX.DMD_6c"
	FlexDMDDict.Add 32903, "VPX.DMD_7c"
	FlexDMDDict.Add 35071, "VPX.DMD_8c"
	FlexDMDDict.Add 35055, "VPX.DMD_9c"
	
	FlexDMDDict.Add 2167, "VPX.DMD_A"
	FlexDMDDict.Add 10767, "VPX.DMD_B"
	FlexDMDDict.Add 57, "VPX.DMD_C"
	FlexDMDDict.Add 8719, "VPX.DMD_D"
	FlexDMDDict.Add 121, "VPX.DMD_E"
	FlexDMDDict.Add 113, "VPX.DMD_F"
	FlexDMDDict.Add 2109, "VPX.DMD_G"
	FlexDMDDict.Add 2166, "VPX.DMD_H"
	FlexDMDDict.Add 8713, "VPX.DMD_I"
	FlexDMDDict.Add 30, "VPX.DMD_J"
	FlexDMDDict.Add 5232, "VPX.DMD_K"
	FlexDMDDict.Add 56, "VPX.DMD_L"
	FlexDMDDict.Add 1334, "VPX.DMD_M"
	FlexDMDDict.Add 4406, "VPX.DMD_N"
	' "O" = 0
	FlexDMDDict.Add 2163, "VPX.DMD_P"
	FlexDMDDict.Add 4159, "VPX.DMD_Q"
	FlexDMDDict.Add 6259, "VPX.DMD_R"
	 ' "S" = 5
	FlexDMDDict.Add 8705, "VPX.DMD_T"
	FlexDMDDict.Add 62, "VPX.DMD_U"
	FlexDMDDict.Add 17456, "VPX.DMD_V"
	FlexDMDDict.Add 20534, "VPX.DMD_W"
	FlexDMDDict.Add 21760, "VPX.DMD_X"
	FlexDMDDict.Add 9472, "VPX.DMD_Y"
	FlexDMDDict.Add 17417, "VPX.DMD_Z"
	
	FlexDMDDict.Add &h400,"VPX.DMD_SingleQuote"
	FlexDMDDict.Add 16640, "VPX.DMD_CloseBracket"
	FlexDMDDict.Add 5120, "VPX.DMD_OpenBracket"
	'FlexDMDDict.Add 2120, "VPX.DMD_Equals"
	FlexDMDDict.Add 10275, "VPX.DMD_Question"
	FlexDMDDict.Add 2112, "VPX.DMD_Minus"
	FlexDMDDict.Add 10861, "VPX.DMD_Dollar"
	FlexDMDDict.Add 6144, "VPX.DMD_GreaterThan"
	'FlexDMDDict.Add 65535, "VPX.DMD_Hash"
	FlexDMDDict.Add 32576, "VPX.DMD_Asterick"
	'FlexDMDDict.Add 10816, "VPX.DMD_Plus"
	'FlexDMDDict.Add 544,"VPX.DMD_Quote"
	FlexDMDDict.Add 17408,"VPX.DMD_FSlash"
	'FlexDMDDict.Add 16384,"VPX.DMD_Comma"
	'FlexDMDDict.Add 32768,"VPX.DMD_Fullstop"
	FlexDMDDict.Add 2140,"VPX.DMD_0BL"
	FlexDMDDict.Add 2147,"VPX.DMD_0TR"
	FlexDMDDict.Add 8,"VPX.DMD_Under"
	FlexDMDDict.Add 3,"VPX.DMD_Over"

	'chars with dots
	FlexDMDDict.Add 32831,"VPX.DMD_Od"
	FlexDMDDict.Add 32774, "VPX.DMD_1d"
	FlexDMDDict.Add 34907, "VPX.DMD_2d"
	FlexDMDDict.Add 34895, "VPX.DMD_3d"
	FlexDMDDict.Add 34918, "VPX.DMD_4d"
	FlexDMDDict.Add 34925, "VPX.DMD_Sd"
	FlexDMDDict.Add 34941, "VPX.DMD_6d"
	FlexDMDDict.Add 32775, "VPX.DMD_7d"
	FlexDMDDict.Add 34943, "VPX.DMD_8d"
	FlexDMDDict.Add 34927, "VPX.DMD_9d"
	
	FlexDMDDict.Add 34935, "VPX.DMD_Ad"
	FlexDMDDict.Add 43535, "VPX.DMD_Bd"
	FlexDMDDict.Add 32825, "VPX.DMD_Cd"
	FlexDMDDict.Add 41487, "VPX.DMD_Dd"
	FlexDMDDict.Add 32889, "VPX.DMD_Ed"
	FlexDMDDict.Add 32881, "VPX.DMD_Fd"
	FlexDMDDict.Add 34877, "VPX.DMD_Gd"
	FlexDMDDict.Add 34934, "VPX.DMD_Hd"
	FlexDMDDict.Add 41481, "VPX.DMD_Id"
	FlexDMDDict.Add 32798, "VPX.DMD_Jd"
	FlexDMDDict.Add 38000, "VPX.DMD_Kd"
	FlexDMDDict.Add 32824, "VPX.DMD_Ld"
	FlexDMDDict.Add 34102, "VPX.DMD_Md"
	FlexDMDDict.Add 37174, "VPX.DMD_Nd"
	FlexDMDDict.Add 34931, "VPX.DMD_Pd"
	FlexDMDDict.Add 36927, "VPX.DMD_Qd"
	FlexDMDDict.Add 39027, "VPX.DMD_Rd"
	FlexDMDDict.Add 41473, "VPX.DMD_Td"
	FlexDMDDict.Add 32830, "VPX.DMD_Ud"
	FlexDMDDict.Add 50224, "VPX.DMD_Vd"
	FlexDMDDict.Add 53302, "VPX.DMD_Wd"
	FlexDMDDict.Add 54528, "VPX.DMD_Xd"
	FlexDMDDict.Add 42240, "VPX.DMD_Yd"
	FlexDMDDict.Add 50185, "VPX.DMD_Zd"
	
	
End sub

Sub UpdateFlexChar(id, value)
	
	If id < 32 Then
		if FlexDMDDict.Exists (value) then
			FlexDMDScene.GetImage("Seg" & id).Bitmap = FlexDMD.NewImage("", FlexDMDDict.Item (value)).Bitmap
		Else
			FlexDMDScene.GetImage("Seg" & id).Bitmap = FlexDMD.NewImage("", "VPX.DMD_Under").Bitmap 'use underscore rather than space for part drawn led chars in this game
		end if
	End If

End Sub

'**********************************************************************************************************

'VR STuff


'***************************************************************************
'Beer Bubble Code - Rawd
'***************************************************************************
Sub BeerTimer_Timer()

	Randomize(21)
	BeerBubble1.z = BeerBubble1.z + Rnd(1)*0.5
	if BeerBubble1.z > -771 then BeerBubble1.z = -955
	BeerBubble2.z = BeerBubble2.z + Rnd(1)*1
	if BeerBubble2.z > -768 then BeerBubble2.z = -955
	BeerBubble3.z = BeerBubble3.z + Rnd(1)*1
	if BeerBubble3.z > -768 then BeerBubble3.z = -955
	BeerBubble4.z = BeerBubble4.z + Rnd(1)*0.75
	if BeerBubble4.z > -774 then BeerBubble4.z = -955
	BeerBubble5.z = BeerBubble5.z + Rnd(1)*1
	if BeerBubble5.z > -771 then BeerBubble5.z = -955
	BeerBubble6.z = BeerBubble6.z + Rnd(1)*1
	if BeerBubble6.z > -774 then BeerBubble6.z = -955
	BeerBubble7.z = BeerBubble7.z + Rnd(1)*0.8
	if BeerBubble7.z > -768 then BeerBubble7.z = -955
	BeerBubble8.z = BeerBubble8.z + Rnd(1)*1
	if BeerBubble8.z > -771 then BeerBubble8.z = -955
End Sub


'***************************************************************************
'VR Clock code below - THANKS RASCAL
'***************************************************************************


' VR Clock code below....
Sub ClockTimer_Timer()
	dim CurrentMinute: Currentminute = Minute(Now())
	VRClockMinutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
	VRClockhours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
    VRClockseconds.RotAndTra2 = (Second(Now()))*6
	CurrentMinute=Minute(Now())
End Sub

'***************************************************************************
' VR Plunger Animataion Code
'***************************************************************************


Dim VRPlungerystart: VRPlungerystart = VRPlunger.Y 
Dim Plungerystart : Plungerystart= Plunger.Y

Sub TimerVRPlunger_Timer
	If VRPlunger.Y < (VRPlungerystart + 100) then VRPlunger.Y = VRPlunger.y + 4  
End Sub

Sub TimerVRPlunger2_Timer
	VRPlunger.Y = Plungerystart - Plunger.y + VRPlungerystart + (5* Plunger.Position) 
end sub



'**********************************************************************************************************
'Digital VR Display
'**********************************************************************************************************
 Dim VRDigits(31)
 VRDigits(0)=Array(ax00, ax05, ax0c, ax0d, ax08, ax01, ax06, ax0f, ax02, ax03, ax04, ax07, ax0b, ax0a, ax09, ax0e)
 VRDigits(1)=Array(ax10, ax15, ax1c, ax1d, ax18, ax11, ax16, ax1f, ax12, ax13, ax14, ax17, ax1b, ax1a, ax19, ax1e)
 VRDigits(2)=Array(ax20, ax25, ax2c, ax2d, ax28, ax21, ax26, ax2f, ax22, ax23, ax24, ax27, ax2b, ax2a, ax29, ax2e)
 VRDigits(3)=Array(ax30, ax35, ax3c, ax3d, ax38, ax31, ax36, ax3f, ax32, ax33, ax34, ax37, ax3b, ax3a, ax39, ax3e)
 VRDigits(4)=Array(ax40, ax45, ax4c, ax4d, ax48, ax41, ax46, ax4f, ax42, ax43, ax44, ax47, ax4b, ax4a, ax49, ax4e)
 VRDigits(5)=Array(ax50, ax55, ax5c, ax5d, ax58, ax51, ax56, ax5f, ax52, ax53, ax54, ax57, ax5b, ax5a, ax59, ax5e)
 VRDigits(6)=Array(ax60, ax65, ax6c, ax6d, ax68, ax61, ax66, ax6f, ax62, ax63, ax64, ax67, ax6b, ax6a, ax69, ax6e)
 VRDigits(7)=Array(ax70, ax75, ax7c, ax7d, ax78, ax71, ax76, ax7f, ax72, ax73, ax74, ax77, ax7b, ax7a, ax79, ax7e)
 VRDigits(8)=Array(ax80, ax85, ax8c, ax8d, ax88, ax81, ax86, ax8f, ax82, ax83, ax84, ax87, ax8b, ax8a, ax89, ax8e)
 VRDigits(9)=Array(ax90, ax95, ax9c, ax9d, ax98, ax91, ax96, ax9f, ax92, ax93, ax94, ax97, ax9b, ax9a, ax99, ax9e)
 VRDigits(10)=Array(axa0, axa5, axac, axad, axa8, axa1, axa6, axaf, axa2, axa3, axa4, axa7, axab, axaa, axa9, axae)
 VRDigits(11)=Array(axb0, axb5, axbc, axbd, axb8, axb1, axb6, axbf, axb2, axb3, axb4, axb7, axbb, axba, axb9, axbe)
 VRDigits(12)=Array(axc0, axc5, axcc, axcd, axc8, axc1, axc6, axcf, axc2, axc3, axc4, axc7, axcb, axca, axc9, axce)
 VRDigits(13)=Array(axd0, axd5, axdc, axdd, axd8, axd1, axd6, axdf, axd2, axd3, axd4, axd7, axdb, axda, axd9, axde)
 VRDigits(14)=Array(axe0, axe5, axec, axed, axe8, axe1, axe6, axef, axe2, axe3, axe4, axe7, axeb, axea, axc9, axee)
 VRDigits(15)=Array(axf0, axf5, axfc, axfd, axf8, axf1, axf6, axff, axf2, axf3, axf4, axf7, axfb, axfa, axf9, axfe)
 VRDigits(16)=Array(bx00, bx05, bx0c, bx0d, bx08, bx01, bx06, bx0f, bx02, bx03, bx04, bx07, bx0b, bx0a, bx09, bx0e)
 VRDigits(17)=Array(bx10, bx15, bx1c, bx1d, bx18, bx11, bx16, bx1f, bx12, bx13, bx14, bx17, bx1b, bx1a, bx19, bx1e)
 VRDigits(18)=Array(bx20, bx25, bx2c, bx2d, bx28, bx21, bx26, bx2f, bx22, bx23, bx24, bx27, bx2b, bx2a, bx29, bx2e)
 VRDigits(19)=Array(bx30, bx35, bx3c, bx3d, bx38, bx31, bx36, bx3f, bx32, bx33, bx34, bx37, bx3b, bx3a, bx39, bx3e)
 VRDigits(20)=Array(bx40, bx45, bx4c, bx4d, bx48, bx41, bx46, bx4f, bx42, bx43, bx44, bx47, bx4b, bx4a, bx49, bx4e)
 VRDigits(21)=Array(bx50, bx55, bx5c, bx5d, bx58, bx51, bx56, bx5f, bx52, bx53, bx54, bx57, bx5b, bx5a, bx59, bx5e)
 VRDigits(22)=Array(bx60, bx65, bx6c, bx6d, bx68, bx61, bx66, bx6f, bx62, bx63, bx64, bx67, bx6b, bx6a, bx69, bx6e)
 VRDigits(23)=Array(bx70, bx75, bx7c, bx7d, bx78, bx71, bx76, bx7f, bx72, bx73, bx74, bx77, bx7b, bx7a, bx79, bx7e)
 VRDigits(24)=Array(bx80, bx85, bx8c, bx8d, bx88, bx81, bx86, bx8f, bx82, bx83, bx84, bx87, bx8b, bx8a, bx89, bx8e)
 VRDigits(25)=Array(bx90, bx95, bx9c, bx9d, bx98, bx91, bx96, bx9f, bx92, bx93, bx94, bx97, bx9b, bx9a, bx99, bx9e)
 VRDigits(26)=Array(bxa0, bxa5, bxac, bxad, bxa8, bxa1, bxa6, bxaf, bxa2, bxa3, bxa4, bxa7, bxab, bxaa, bxa9, bxae)
 VRDigits(27)=Array(bxb0, bxb5, bxbc, bxbd, bxb8, bxb1, bxb6, bxbf, bxb2, bxb3, bxb4, bxb7, bxbb, bxba, bxb9, bxbe)
 VRDigits(28)=Array(bxc0, bxc5, bxcc, bxcd, bxc8, bxc1, bxc6, bxcf, bxc2, bxc3, bxc4, bxc7, bxcb, bxca, bxc9, bxce)
 VRDigits(29)=Array(bxd0, bxd5, bxdc, bxdd, bxd8, bxd1, bxd6, bxdf, bxd2, bxd3, bxd4, bxd7, bxdb, bxda, bxd9, bxde)
 VRDigits(30)=Array(bxe0, bxe5, bxec, bxed, bxe8, bxe1, bxe6, bxef, bxe2, bxe3, bxe4, bxe7, bxeb, bxea, bxe9, bxee)
 VRDigits(31)=Array(bxf0, bxf5, bxfc, bxfd, bxf8, bxf1, bxf6, bxff, bxf2, bxf3, bxf4, bxf7, bxfb, bxfa, bxf9, bxfe)
 
Sub FadeDisplay(object, onoff, num)
	If OnOff = 1 Then
			object.color = RGB(255,88,32)
		Object.Opacity = 20
	Else
		Object.Color = RGB(1,1,1)
		Object.Opacity = 1
	End If
End Sub

Sub InitDigits()
	dim tmp, x, obj
	for x = 0 to uBound(VRDigits)
		if IsArray(VRDigits(x) ) then
			For each obj in VRDigits(x)
				obj.height = obj.height + 0
				FadeDisplay obj, 0 , 0
			next
		end If
	Next
End Sub

Sub setup_backglass()
	Dim obj, ix,xx,yy
	Dim xoff:xoff = -50
	Dim yoff:yoff = 128
	Dim zoff:zoff = 699
	Dim xrot:xrot = -90
	Dim zscale:zscale = 0.00000001

	Dim xcen:xcen = 0  '(130 /2) - (92 / 2)
	Dim ycen:ycen = (780 /2 ) + (203 /2)
	'backglass Ledds
	for ix = 0 to Ubound(VRDigits)
		For Each obj In VRDigits(ix)

			xx = obj.x  
				
			obj.x = (xoff - xcen) + xx
			yy = obj.y ' get the yoffset before it is changed
			if ix < 20 then
				obj.y = yoff + 28 'This will set the Topsegment closer to user
			Elseif ix < 40 then
				obj.y = yoff + 33 'This will set the lower segment closer to user
			Else	
				obj.y = yoff 
			end if
			

			If (yy < 0.) then
				yy = yy * -1
			end if

			obj.height = (zoff - ycen) + yy - (yy * (zscale))
			obj.rotx = xrot
		Next
	Next
	'Backglass Flashers
	For Each obj In VRBackglassLights
		yy=obj.height '' because of rotation, ned to adjustthe objects a bit
		obj.x = obj.x
		obj.height = - obj.y
		obj.y = xoff+yy'adjusts the distance from the backglass towards the user
		obj.rotx = xrot
	Next

	'For Each obj In VRBackglassflasher
	'	obj.visible = false
	'Next
	' BackGlass Image
	'X 484,1859
	'y 49,22256
	VRBackglass.y = 109
	VRBackglass.x = 484
end sub

'**********************************************************************************************************