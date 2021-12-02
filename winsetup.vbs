Const DefaultAutoRepeatDelay = 150 'windows default is 250ms
Const DefaultMouseHoverDelay = 150 'windows default is 400ms

dim runasAdmin
If WScript.Arguments.length=0 Then
    runasAdmin  = MsgBox("Change settings globally?"&vbNewLine&vbNewLine&"(click No if just for this user)", 3, "Run as admin?")
    If vbCancel = runasAdmin Then WScript.Quit
    If vbYes    = runasAdmin Then
        Set objShell = CreateObject("Shell.Application")
        objShell.ShellExecute "wscript.exe", Chr(34) & _
        WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    Else
        MainDialogue
    End If
Else
    runasAdmin  = vbYes
    MainDialogue
End If

Function MainDialogue()
    Const HKEY_CLASSES_ROOT  = &H80000000
    Const HKEY_CURRENT_USER  = &H80000001
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const HKEY_USERS         = &H80000003
    Const REG_SZ        = 1
    Const REG_EXPAND_SZ = 2
    Const REG_BINARY    = 3
    Const REG_DWORD     = 4
    Const REG_MULTI_SZ  = 7
    Const strComputer = "."
    Set objShell = CreateObject("WScript.Shell")
    Set objRegistry = GetObject ("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
    Dim SIDs : SIDs = Array("", ".DEFAULT\", "S-1-5-18\", "S-1-5-19\", "S-1-5-20\")
    Dim hiveTarget, KeyPath, ValueName
    Dim MarkCAccelCurveX, MarkCAccelCurveY
    Dim MouseHoverDelay : MouseHoverDelay = InputBox("Enter your preferred hover delay in miliseconds."&vbNewLine&vbNewLine&"Windows default is 400 ms."&vbNewLine&"Recommended is 150 ms."&vbNewLine&vbNewLine&"Click cancel if you don't want to change it.","Set Mouse Hover Delay", DefaultMouseHoverDelay)
    Dim FilterKeysDelay : FilterKeysDelay = InputBox("Enter your preferred repeat delay in miliseconds."&vbNewLine&vbNewLine&"Default Windows minimum is 250 ms."&vbNewLine&"Zowie Celeritas in PS/2 mode is 200 ms."&vbNewLine&vbNewLine&"Recommend 150 ms if you have fast fingers."&vbNewLine&vbNewLine&"Click cancel if you don't want to change it.","Set Keyboard Repeat Delay", DefaultAutoRepeatDelay)
    CalculateMarkCString MarkCAccelCurveX, MarkCAccelCurveY

    For each User in SIDs
        hiveTarget = HKEY_USERS
        If User = "" Then
            hiveTarget = HKEY_CURRENT_USER
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPath", 1 ' show full path in titlebar
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel", "AllItemsIconView", 2 ' control panel show all icons
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSecondsInSystemClock", 1 ' ticking seconds in tray
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "LaunchTo", 1 ' explorer default to ThisPC
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowCortanaButton", 0 ' hide cortana
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 0 ' show file extension
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "UseCompactView", 1 ' explorer compact view
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 0 ' default to darktheme
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme", 0 'default to darktheme
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Windows\CurrentVersion\Search", "SearchboxTaskbarMode", 0 ' hide search bar
            objRegistry.SetDWORDValue  hiveTarget, "Software\Microsoft\Multimedia\Audio", "UserDuckingPreference", 3 ' don't lower system volume when there's communication audio
            objRegistry.SetDWORDValue  hiveTarget, "System\GameConfigStore", "GameDVR_Enabled", 0 ' 
            objRegistry.SetDWORDValue  hiveTarget, "System\GameConfigStore", "GameDVR_FSEBehavior", 2 ' 
            objRegistry.SetDWORDValue  hiveTarget, "System\GameConfigStore", "GameDVR_FSEBehaviorMode", 2 ' 
            objRegistry.SetDWORDValue  hiveTarget, "System\GameConfigStore", "GameDVR_HonorUserFSEBehaviorMode", 1 '
            objRegistry.SetDWORDValue  hiveTarget, "System\GameConfigStore", "GameDVR_DXGIHonorFSEWindowsCompatible", 1 ' 
            ' https://gist.github.com/CHEF-KOCH/ddd1fa24b899ab9fa2d4 for more options
        End If

        ' set tasktray clock and date format (TODO: need to add customization)
        KeyPath = User + "Control Panel\International"
            objRegistry.SetStringValue hiveTarget, KeyPath, "sLongDate", "ddd, MMMM d, yyyy"
            objRegistry.SetStringValue hiveTarget, KeyPath, "sShortDate", "ddd yyyy-MM-dd"
            objRegistry.SetStringValue hiveTarget, KeyPath, "sShortTime", "h:mm tt"
            objRegistry.SetStringValue hiveTarget, KeyPath, "sTimeFormat", "h:mm:ss tt"
            objRegistry.SetStringValue hiveTarget, KeyPath, "s1159", "AM"
            objRegistry.SetStringValue hiveTarget, KeyPath, "s2359", "PM"

        ' disable mouse accel
        KeyPath = User + "Control Panel\Mouse"
            objRegistry.SetStringValue hiveTarget, KeyPath, "MouseSpeed", "0"
            objRegistry.SetStringValue hiveTarget, KeyPath, "MouseSensitivity", "10"
            objRegistry.SetStringValue hiveTarget, KeyPath, "MouseThreshold1", "0"
            objRegistry.SetStringValue hiveTarget, KeyPath, "MouseThreshold2", "0"
            objRegistry.SetStringValue hiveTarget, KeyPath, "MouseHoverTime", MouseHoverDelay
            objRegistry.SetBinaryValue hiveTarget, KeyPath, "SmoothMouseXCurve", MarkCAccelCurveX
            objRegistry.SetBinaryValue hiveTarget, KeyPath, "SmoothMouseYCurve", MarkCAccelCurveY
        KeyPath = User + "Control Panel\Desktop"
            objRegistry.SetStringValue hiveTarget, KeyPath, "MenuShowDelay", MouseHoverDelay

        ' improve keyboard responsiveness
        KeyPath = User + "Control Panel\Keyboard"
            objRegistry.SetStringValue hiveTarget, KeyPath, "KeyboardDelay", "0"
            objRegistry.SetStringValue hiveTarget, KeyPath, "KeyboardSpeed", "31"
        KeyPath = User + "Control Panel\Accessibility\Keyboard Response"
         If Not FilterKeysDelay = "" Then
            objRegistry.SetStringValue hiveTarget, KeyPath, "AutoRepeatDelay", FilterKeysDelay
            objRegistry.SetStringValue hiveTarget, KeyPath, "AutoRepeatRate", "1"
            objRegistry.SetStringValue hiveTarget, KeyPath, "BounceTime", "0"
            objRegistry.SetStringValue hiveTarget, KeyPath, "DelayBeforeAcceptance", "0"
            objRegistry.SetStringValue hiveTarget, KeyPath, "Flags", "1"
         End If
            objRegistry.SetDWORDValue  hiveTarget, KeyPath, "Last BounceKey Setting", 0
            objRegistry.SetDWORDValue  hiveTarget, KeyPath, "Last Valid Delay", DefaultAutoRepeatDelay
            objRegistry.SetDWORDValue  hiveTarget, KeyPath, "Last Valid Repeat", 1
            objRegistry.SetDWORDValue  hiveTarget, KeyPath, "Last Valid Wait", 0

        KeyPath = User + "Control Panel\Accessibility\MouseKeys"
            objRegistry.SetStringValue hiveTarget, KeyPath, "Flags", "126"
            objRegistry.SetStringValue hiveTarget, KeyPath, "MaximumSpeed", "358"
            objRegistry.SetStringValue hiveTarget, KeyPath, "TimeToMaximumSpeed", "1000"

        ' underline ALT hotkeys
        KeyPath = User + "Control Panel\Accessibility\Keyboard Preference"
            objRegistry.SetStringValue hiveTarget, KeyPath, "On", "1"

        ' disable the annoying popup when you press SHIFT five times
        KeyPath = User + "Control Panel\Accessibility\StickyKeys"
            objRegistry.SetStringValue hiveTarget, KeyPath, "Flags", "506"

        ' disable the annoying popup when you hold SHIFT for eight seconds
        KeyPath = User + "Control Panel\Accessibility\ToggleKeys"
            objRegistry.SetStringValue hiveTarget, KeyPath, "Flags", "58"

    Next
    MsgBox "Remember to log off to apply settings.", 0, "Success!"
End Function

Function CalculateMarkCString(ByRef MarkCAccelCurveX, ByRef MarkCAccelCurveY)
    dim MarkCStringX : MarkCStringX = "00,00,00,00,00,00,00,00,C0,CC,0C,00,00,00,00,00,80,99,19,00,00,00,00,00,40,66,26,00,00,00,00,00,00,33,33,00,00,00,00,00"
    dim MarkCStringY : MarkCStringY = "00,00,00,00,00,00,00,00,00,00,38,00,00,00,00,00,00,00,70,00,00,00,00,00,00,00,A8,00,00,00,00,00,00,00,E0,00,00,00,00,00"
    dim MarkCConfirm : MarkCConfirm = MsgBox("Would you like apply MarkC's Mouse Fix?"&vbNewLine&vbNewLine&"(Linearize the transfer function to 1-to-1, for legacy games that force-on Windows Pointer Accel)", 3, "MarkC's Mouse Fix (Neutralize Windows Pointer Accel)")
    If  MarkCConfirm = vbCancel Then WScript.Quit
    If  MarkCConfirm = vbYes    Then 
        If vbYes = MsgBox("Would you like to customize the fix?", 4, "Customize MarkC's Mouse Fix?") Then MarkC_Dialogues MarkCStringX, MarkCStringY
        MarkCAccelCurveX = CurveStringToBinary(MarkCStringX)
        MarkCAccelCurveY = CurveStringToBinary(MarkCStringY)
    End If
End Function

Function CurveStringToBinary(str)
    dim uBinary : uBinary = Split(str, ",")
    For i = 0 To UBound(uBinary) : uBinary(i) = CByte("&H" & uBinary(i)) : Next
    CurveStringToBinary = uBinary
End Function



' ============================================= MarkC code begins =========================================================

const XPVistaFix  = "XP+Vista"
const Windows7Fix = "Windows 7"
const Windows8Fix = "Windows 8.1+8"
const Windows10Fix = "Windows 10"
dim WshShell, WMIService, OS, VideoController, Shell, Folder, RegFile, FileSystem ' as Object
dim OSVersion, OSVersionNumber, FixType, DPI, RefreshRate, MouseSensitivity, DPISensitivity, EPPOffMouseScaling
dim MouseSensitivityFactor, PointerSpeed, Scaling, ScalingDescr
dim RegFilename, FolderName, OSConfirmText, RegComment
dim SmoothX, SmoothY, SmoothYPixels, DPIFactor
dim SpeedSteps, Threshold1, Threshold2, ScalingAfterThreshold1, ScalingAfterThreshold2
dim OldWindowsAccelName, OldWindowsAccelCode, OldWindowsThreshold1, OldWindowsThreshold2
dim SmoothX0, SmoothX1, SmoothX2, SmoothX3, SmoothX4
dim SmoothY0, SmoothY1, SmoothY2, SmoothY3, SmoothY4
function MarkC_Dialogues(ByRef MarkCStringX, ByRef MarkCStringY)

    set WshShell = WScript.CreateObject("WScript.Shell")
    set WMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

    ' Get the OS on this machine, and display DPI
    for each OS in WMIService.ExecQuery("select * from Win32_OperatingSystem where Primary='True'")
        OSVersion = Left(OS.Version, InStr(InStr(OS.Version,".")+1, OS.Version, ".")-1)
        OSVersionNumber = OSVersion
        DPI = 96 ' Default in case RegRead errors
        EPPOffMouseScaling = 1
        on error resume next ' On 7, Desktop\LogPixels not present until DPI is changed
        select case OSVersion
        case "5.1","6.0"
            FixType = XPVistaFix
            OSVersion = XPVistaFix
            DPI = WshShell.RegRead("HKEY_CURRENT_CONFIG\Software\Fonts\LogPixels")
        case "6.1"
            FixType = Windows7Fix
            OSVersion = Windows7Fix
            DPI = WshShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\LogPixels")
        case "6.2"
            FixType = Windows8Fix
            OSVersion = Windows8Fix
            DPI = WshShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\LogPixels")
        case "6.3"
            FixType = Windows8Fix
            OSVersion = Windows8Fix
            DPI = WshShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\AppliedDPI")
        case "10.0"
            FixType = Windows10Fix
            OSVersion = Windows10Fix
            DPI = WshShell.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\AppliedDPI")
            if Fix(OS.BuildNumber) >= 16299 then EPPOffMouseScaling = -1 ' New EPP=OFF mouse scaling for Windows 10 1709+
        case else
            FixType = Windows10Fix
        end select
        on error goto 0 ' Normal errors
    next

    ' Get which OS the fix will be used for
    do
        FixType = InputBox( _
            "This script program modifies the registry accel curve to remove Windows' mouse acceleration," _
            & " or emulate Windows 2000 or Windows 98 or Windows 95 acceleration." _
            & vbNewLine & vbNewLine _
            & "The fix works like the CPL and Cheese and MarkC fixes," _
            & " but is customized for your specific desktop display text size (DPI)," _
            & " your specific mouse pointer speed slider setting, your specific refresh rate" _
            & " and has any pointer speed scaling/sensitivity factor you want." _
            & vbNewLine & vbNewLine _
            & "Enter the operating system that the fix will be used for." _
            & vbNewLine & vbNewLine _
            & "1/XP+Vista   	= Windows XP or Vista" & vbNewLine _
            & "2/Windows 7  	= Windows 7" & vbNewLine _
            & "3/Windows 8.x	= Windows 8.1 or 8" & vbNewLine _
            & "4/Windows 10 	= Windows 10", _
            "Operating System - MarkC Mouse Acceleration Fix", FixType)
        select case LCase(FixType)
        case ""
            WScript.Quit
        case "1", "xp", "vista", "xpvista", "xp+vista", "xp or vista"
            FixType = XPVistaFix
            exit do
        case "2", "7", "win7", "windows7", "windows 7"
            FixType = Windows7Fix
            DPI = int((100*DPI/96)+0.5) ' round() rounds to even: we don't want that
            exit do
        case "3", "8", "win8", "windows8", "windows 8", "8.1", "win8.1", "windows8.1", "windows 8.1", "windows 8.x"
            FixType = Windows8Fix
            DPI = int((100*DPI/96)+0.5) ' round() rounds to even: we don't want that
            exit do
        case "4", "10", "win10", "windows10", "windows 10"
            FixType = Windows10Fix
            DPI = int((100*DPI/96)+0.5) ' round() rounds to even: we don't want that
            exit do
        case else
            WshShell.Popup "'" & FixType & "' is not valid.",, "Error", vbExclamation
        end select
    loop while true

    ' Get the display DPI the fix will be used with
    dim CurrentDPI : CurrentDPI = DPI
    do
        dim NumDPI
        if FixType = XPVistaFix then
            DPI = InputBox( _
                "Enter the desktop Control Panel, Display, font size (DPI) scaling setting that will be used." _
                & vbNewLine & vbNewLine _
                & "Your current font size (DPI) is " & CurrentDPI & ".", _
                "Display DPI - MarkC Mouse Acceleration Fix", DPI)
            if DPI = "" then WScript.Quit
            if IsNumeric(DPI) then
                NumDPI = CInt(DPI)
                if NumDPI > 0 and CStr(NumDPI) = DPI then DPI = NumDPI : exit do
            end if
            WshShell.Popup "'" & DPI & "' is not valid.",, "Error", vbExclamation
        else if FixType = Windows7Fix then
            if InStr(DPI,"%") = 0 then DPI = DPI & "%"
            if InStr(CurrentDPI,"%") = 0 then CurrentDPI = CurrentDPI & "%"
            DPI = InputBox( _
                "Enter the desktop Control Panel, Display, text size (DPI) that will be used." _
                & vbNewLine & vbNewLine _
                & "Your current text size (DPI) is " & CurrentDPI & ".", _
                "Text Size DPI - MarkC Mouse Acceleration Fix", DPI)
            if DPI = "" then WScript.Quit
            if InStr(DPI,"%") = len(DPI) and IsNumeric(left(DPI,len(DPI)-1)) then
                NumDPI = CInt(left(DPI,len(DPI)-1))
                if NumDPI > 0 and CStr(NumDPI) & "%" = DPI then DPI = NumDPI : exit do
            end if
            WshShell.Popup "'" & DPI & "' is not valid.",, "Error", vbExclamation
        else ' Windows10Fix or Windows8Fix
            if InStr(DPI,"%") = 0 then DPI = DPI & "%"
            if InStr(CurrentDPI,"%") = 0 then CurrentDPI = CurrentDPI & "%"
            dim Windows81Text : Windows81Text = ""
            if OSVersionNumber >= "6.3" then _
                Windows81Text = "If not using one scaling level for all displays, then" _
                    & vbNewLine & "- the 1st slider position should be 100%," _
                    & vbNewLine & "- the 2nd slider position should be 125%," _
                    & vbNewLine & "- the 3rd slider position (might not be shown) should be 150%" _
                    & vbNewLine _
                    & vbNewLine & "(Very high DPI monitors might need a custom size to get exact 1-to-1.)" _
                    & vbNewLine & vbNewLine
            DPI = InputBox( _
                "Enter the desktop Settings, Display, Scale and layout, size of items setting that will be used." _
                & vbNewLine & vbNewLine _
                & Windows81Text _
                & "Your current size of items setting is " & CurrentDPI & ".", _
                "Items Size - MarkC Mouse Acceleration Fix", DPI)
            if DPI = "" then WScript.Quit
            if InStr(DPI,"%") = len(DPI) and IsNumeric(left(DPI,len(DPI)-1)) then
                NumDPI = CInt(left(DPI,len(DPI)-1))
                if NumDPI > 0 and CStr(NumDPI) & "%" = DPI then DPI = NumDPI : exit do
            end if
            WshShell.Popup "'" & DPI & "' is not valid.",, "Error", vbExclamation
        end if end if
    loop while true

    if FixType = XPVistaFix then
        ' Get the monitor refresh rate the fix will be used with
        for each VideoController in WMIService.InstancesOf("Win32_VideoController")
            RefreshRate = VideoController.CurrentRefreshRate
            if not IsNull(RefreshRate) then exit for
        next
        if IsNull(RefreshRate) then
            for each VideoController in WMIService.InstancesOf("Win32_DisplayConfiguration")
                RefreshRate = VideoController.DisplayFrequency
                if not IsNull(RefreshRate) then exit for
            next
        end if
        if IsNull(RefreshRate) then RefreshRate = "(unknown)"
        dim CurrentRefreshRate : CurrentRefreshRate = RefreshRate
        do
            RefreshRate = InputBox( _
                "Enter the in-game monitor refresh rate that will be used." _
                & vbNewLine & vbNewLine _
                & "NOTE: Your desktop refresh rate is " & CurrentRefreshRate & "Hz." _
                & vbNewLine & vbNewLine _
                & "Enter the refresh rate USED BY YOUR GAME, when the fix will be active.", _
                "Refresh Rate - MarkC Mouse Acceleration Fix", RefreshRate)
            if RefreshRate = "" then WScript.Quit
            if IsNumeric(RefreshRate) then
                dim NumRefreshRate : NumRefreshRate = CInt(RefreshRate)
                if NumRefreshRate > 0 and CStr(NumRefreshRate) = RefreshRate then RefreshRate = NumRefreshRate : exit do
            end if
            WshShell.Popup "'" & RefreshRate & "' is not valid, enter a number.",, "Error", vbExclamation
        loop while true
    end if

    ' Get the pointer speed slider setting the fix will be used with
    MouseSensitivity = CInt(WshShell.RegRead("HKEY_CURRENT_USER\Control Panel\Mouse\MouseSensitivity"))
    if MouseSensitivity > 1 then PointerSpeed = MouseSensitivity/2 else PointerSpeed = 0
    dim FormattedPointerSpeed : FormattedPointerSpeed = CStr(PointerSpeed) & "th"
    if PointerSpeed = 1 then FormattedPointerSpeed = CStr(PointerSpeed) & "st"
    if PointerSpeed = 2 then FormattedPointerSpeed = CStr(PointerSpeed) & "nd"
    if PointerSpeed = 3 then FormattedPointerSpeed = CStr(PointerSpeed) & "rd"
    do
        MouseSensitivity = InputBox( _
            "Enter the Pointer Speed value that will be used." _
            & vbNewLine & vbNewLine _
            & "Value	Notch	RTS	noEPP	EPP ON" & vbNewLine _
            & "1	0th	5%	x1/32	x0.1" & vbNewLine _
            & "2	1st	10%	x1/16	x0.2" & vbNewLine _
            & "3		15%	x1/8	x0.3" & vbNewLine _
            & "4	2nd	20%	x2/8	x0.4" & vbNewLine _
            & "5		25%	x3/8	x0.5" & vbNewLine _
            & "6	3rd	30%	x4/8	x0.6" & vbNewLine _
            & "7		35%	x5/8	x0.7" & vbNewLine _
            & "8	4th	40%	x6/8	x0.8" & vbNewLine _
            & "9		45%	x7/8	x0.9" & vbNewLine _
            & "10	5th	50%	x1	x1.0" & vbNewLine _
            & "11		55%	x1.25	x1.1" & vbNewLine _
            & "12	6th	60%	x1.5	x1.2" & vbNewLine _
            & "13		65%	x1.75	x1.3" & vbNewLine _
            & "14	7th	70%	x2	x1.4" & vbNewLine _
            & "15		75%	x2.25	x1.5" & vbNewLine _
            & "16	8th	80%	x2.5	x1.6" & vbNewLine _
            & "17		85%	x2.75	x1.7" & vbNewLine _
            & "18	9th	90%	x3	x1.8" & vbNewLine _
            & "19		95%	x3.25	x1.9" & vbNewLine _
            & "20	10th	100%	x3.5	x2.0" _
            & vbNewLine & vbNewLine _
            & "Your current Pointer Speed is " & MouseSensitivity & vbNewLine & "(The " & FormattedPointerSpeed & " notch in Control Panel)", _
            "Pointer Speed Slider - MarkC Mouse Acceleration Fix", CStr(MouseSensitivity))
        if MouseSensitivity = "" then WScript.Quit
        if IsNumeric(MouseSensitivity) then
            dim NumSpeed : NumSpeed = CDbl(MouseSensitivity)
            if NumSpeed >= 1 and NumSpeed <= 20 and int(NumSpeed) = NumSpeed then
                MouseSensitivity = NumSpeed
                exit do
            end if
        end if
        WshShell.Popup "'" & MouseSensitivity & "' is not valid.",, "Error", vbExclamation
    loop while true

    ' Convert pointer speed slider to a numeric sensitivity
    if MouseSensitivity > 1 then PointerSpeed = MouseSensitivity/2 else PointerSpeed = 0
    if MouseSensitivity <= 2 then
        MouseSensitivityFactor = MouseSensitivity / 32
    elseif MouseSensitivity <= 10 then
        MouseSensitivityFactor = (MouseSensitivity-2) / 8
    else
        MouseSensitivityFactor = (MouseSensitivity-6) / 4
    end if

    ' Get the number of pointer acceleration zones
    SpeedSteps = "No acceleration" : Threshold1 = 0 : Threshold2 = 0
    do
        dim SpeedStepsPrompt
        if FixType = Windows10Fix then
            SpeedStepsPrompt = _
                "Enter the number of pointer speed acceleration zones that you want." & vbNewLine & vbNewLine _
                & "0 = No acceleration" & vbNewLine _
                & "1 = Accelerate the pointer speed when the mouse is faster than a threshold" & vbNewLine _
                & "2 = [Not available for Windows 10]" & vbNewLine _
                & vbNewLine _
                & "Low	= Emulate Windows 2000 Low accel" & vbNewLine _
                & "Medium	= [Not available for Windows 10]" & vbNewLine _
                & "High	= [Not available for Windows 10]" & vbNewLine _
                & vbNewLine _
                & "2/7  = Emulate Windows 95+98 2/7 pointer speed" & vbNewLine _
                & "3/7  = Emulate Windows 95+98 3/7 pointer speed" & vbNewLine _
                & "4/7  = Emulate Windows 95+98 4/7 pointer speed" & vbNewLine _
                & "5+/7 = [Not available for Windows 10]"
        else
            SpeedStepsPrompt = _
                "Enter the number of pointer speed acceleration zones that you want." & vbNewLine & vbNewLine _
                & "0 = No acceleration" & vbNewLine _
                & "1 = Accelerate the pointer speed when the mouse is faster than a threshold" & vbNewLine _
                & "2 = Accelerate the pointer speed when the mouse is faster than threshold 1," _
                & " and accelerate again when the mouse is faster than threshold 2" & vbNewLine _
                & vbNewLine _
                & "Low	= Emulate Windows 2000 Low accel" & vbNewLine _
                & "Medium	= Emulate Windows 2000 Medium accel" & vbNewLine _
                & "High	= Emulate Windows 2000 High accel" & vbNewLine _
                & vbNewLine _
                & "2/7  = Emulate Windows 95+98 2/7 pointer speed" & vbNewLine _
                & "n/7  = Emulate Windows 95+98 n/7 pointer speed" & vbNewLine _
                & "7/7  = Emulate Windows 95+98 7/7 pointer speed"
        end if
        SpeedSteps = InputBox(SpeedStepsPrompt, "Pointer Speed Acceleration - MarkC Mouse Acceleration Fix", SpeedSteps)
        if SpeedSteps = "" then WScript.Quit
        select case LCase(SpeedSteps)
        case "0", "no acceleration", "none", "no"
            SpeedSteps = 0
            exit do
        case "low"
            SpeedSteps = 1 : Threshold1 = 7
            OldWindowsAccelName = "Windows 2000 Low" : OldWindowsAccelCode = "W2K_Low"
            exit do
        case "medium"
            SpeedSteps = 2 : Threshold1 = 4 : Threshold2 = 12
            OldWindowsAccelName = "Windows 2000 Medium" : OldWindowsAccelCode = "W2K_Medium"
            exit do
        case "high"
            SpeedSteps = 2 : Threshold1 = 4 : Threshold2 = 6
            OldWindowsAccelName = "Windows 2000 High" : OldWindowsAccelCode = "W2K_High"
            exit do
        case "2/7", "3/7", "4/7", "5/7", "6/7", "7/7"
            OldWindowsAccelName = "Windows 95+98 " & SpeedSteps
            OldWindowsAccelCode = "W95+98_" & Replace(SpeedSteps, "/", "of")
            select case SpeedSteps
            case "2/7"
                SpeedSteps = 1 : Threshold1 = 10
            case "3/7"
                SpeedSteps = 1 : Threshold1 = 7
            case "4/7"
                SpeedSteps = 1 : Threshold1 = 4
            case "5/7"
                SpeedSteps = 2 : Threshold1 = 4 : Threshold2 = 12
            case "6/7"
                SpeedSteps = 2 : Threshold1 = 4 : Threshold2 = 9
            case "7/7"
                SpeedSteps = 2 : Threshold1 = 4 : Threshold2 = 6
            end select
            exit do
        end select
        if IsNumeric(SpeedSteps) then
            dim NumSpeedSteps : NumSpeedSteps = CInt(SpeedSteps)
            if NumSpeedSteps >= 0 and NumSpeedSteps <= 2 and CStr(NumSpeedSteps) = SpeedSteps then _
                SpeedSteps = NumSpeedSteps : exit do
        end if
        WshShell.Popup "'" & SpeedSteps & "' is not valid.",, "Error", vbExclamation
    loop while true
    ' Record standard thresholds for info messages later
    if OldWindowsAccelName <> "" then OldWindowsThreshold1 = Threshold1 : OldWindowsThreshold2 = Threshold2

    ' SpeedSteps = 2 causes BugCheck/BSOD on Windows 10 x64
    if FixType = Windows10Fix and SpeedSteps = 2 then
        WshShell.Popup _
            "2 pointer speed acceleration zones are not available for Windows 10," _
            & " because it causes BugChecks/BSOD.", , _
            "Not available for Windows 10", vbCritical
        WScript.Quit
    end if

    ' Get the scaling (sensitivity) factor
    dim SegmentText, ThresholdText
    if SpeedSteps > 0 then
        SegmentText = " when the pointer is not accelerated,"
        if MouseSensitivityFactor = 1 then ScalingDescr = "1-to-1" else ScalingDescr = CStr(MouseSensitivityFactor)
        ThresholdText = vbNewLine _
            & vbNewLine _
            & "The pointer speed factor used by Windows 2000 at " & CStr(PointerSpeed) & "th notch is " _
            & ScalingDescr & "." & vbNewLine _
            & "The pointer speed factor used by Windows 95+98 is 1-to-1."
    end if
    Scaling = "1-to-1"
    if FixType = XPVistaFix then
        DPISensitivity = (Max(60,RefreshRate)/Max(96,DPI)) * 96 / 60
        EPPOffMouseScaling = 1
    else
        DPISensitivity = round(DPI*96/100)/96
        if EPPOffMouseScaling = -1 then EPPOffMouseScaling = DPISensitivity
    end if
    do
        Scaling = InputBox( _
            "Enter the pointer speed scaling (sensitivity) factor that you want" & SegmentText _
            & " when the pointer speed slider is at the " & CStr(PointerSpeed) & "th notch." _
            & vbNewLine & vbNewLine _
            & "1/1-to-1	= Exactly 1-to-1 (RECOMMENDED)" & vbNewLine _
            & "E	= x " & CStr(DPISensitivity * MouseSensitivity/10) & " (same as EPP=ON, enter 'E')" & vbNewLine _
            & "N	= x " & Cstr(EPPOffMouseScaling * MouseSensitivityFactor) & " (same as EPP=OFF, enter 'N')" & vbNewLine _
            & replace(CStr(1.111),"1","n",1,-1) & "	= a custom speed factor (example: " & CStr(1.25) & ")" _
            & ThresholdText, _
            "Pointer Speed Scaling - MarkC Mouse Acceleration Fix", Scaling)
        if Scaling = "" then WScript.Quit
        select case LCase(Scaling)
        case "1", "1-to-1", "1/1-to-1"
            Scaling = 1
            exit do
        case "e"
            Scaling = DPISensitivity * MouseSensitivity/10
            exit do
        case "n"
            Scaling = EPPOffMouseScaling * MouseSensitivityFactor
            exit do
        end select
        if IsNumeric(Scaling) then if CDbl(Scaling) > 0 and CDbl(Scaling) <= 16 then Scaling = CDbl(Scaling) : exit do
        WshShell.Popup "'" & Scaling & "' is not valid.",, "Error", vbExclamation
    loop while true

    dim NumThreshold, ThresholdNotes
    if SpeedSteps > 0 then
        ' Get the first (or only) threshold and sensitivity when speed is > threshold1
        if MsgBox("Notes:" & vbNewLine _
            & vbNewLine _
            & "- If your current mouse has a different polling rate than the mouse you used with" _
            & " your old version of Windows, then the thresholds may need to be adjusted before" _
            & " mouse response will be similar." & vbNewLine _
            & vbNewLine _
            & "- Acceleration will most closely match your old version of Windows for movements" _
            & " that are mainly horizontal or mainly vertical." & vbNewLine _
            & "If your mouse movements are often diagonal or at an angle, then the thresholds may" _
            & " need to be increased by 10% to 30% before mouse response will be similar." & vbNewLine _
            & vbNewLine _
            & "- See file !Threshold_Acceleration_ReadMe.txt for guidance.", _
            vbOKCancel + vbInformation, "Acceleration Thresholds - MarkC Mouse Acceleration Fix") <> vbOK then WScript.Quit
        do ' Get threshold1
            if SpeedSteps = 1 then
                SegmentText = ""
            else
                SegmentText = "first "
            end if
            ThresholdText = ""
            if OldWindowsAccelName <> "" then _
                ThresholdText = vbNewLine & vbNewLine _
                    & "The " & SegmentText & "threshold for " & OldWindowsAccelName & " acceleration is " _
                    & OldWindowsThreshold1 & "."
            Threshold1 = InputBox( _
                "Enter the " & SegmentText & "acceleration threshold that you want." & vbNewLine _
                & vbNewLine _
                & "When the mouse is faster than this, the pointer speed will be accelerated." & vbNewLine _
                & vbNewLine _
                & "See the !Threshold_Acceleration_ReadMe.txt file for guidance." _
                & ThresholdText, _
                "Pointer Speed Acceleration - MarkC Mouse Acceleration Fix", Threshold1)
            if Threshold1 = "" then WScript.Quit
            if IsNumeric(Threshold1) then
                NumThreshold = CInt(Threshold1)
                if NumThreshold > 0 and CStr(NumThreshold) = Threshold1 then Threshold1 = NumThreshold : exit do
            end if
            WshShell.Popup "'" & Threshold1 & "' is not valid (must be greater than 0).",, "Error", vbExclamation
        loop while true

        ThresholdText = vbNewLine _
            & vbNewLine _
            & "The pointer speed factor used by Windows 2000 at " & CStr(PointerSpeed) & "/11 is " _
            & CStr(2*MouseSensitivityFactor) & "." & vbNewLine _
            & "The pointer speed factor used by Windows 95+98 is 2."
        ScalingAfterThreshold1 = 2 * Scaling
        do ' Get the scaling (sensitivity) factor when faster than threshold1
            ScalingAfterThreshold1 = InputBox( _
                "Enter the pointer speed scaling (sensitivity) factor that you want" _
                & " when the mouse is faster than " & CStr(Threshold1) & "." _
                & ThresholdText, _
                "Pointer Speed Scaling - MarkC Mouse Acceleration Fix", ScalingAfterThreshold1)
            if ScalingAfterThreshold1 = "" then WScript.Quit
            if IsNumeric(ScalingAfterThreshold1) then _
                if CDbl(ScalingAfterThreshold1) > Scaling and CDbl(ScalingAfterThreshold1) <= 16 then _
                    ScalingAfterThreshold1 = CDbl(ScalingAfterThreshold1) : exit do
            WshShell.Popup "'" & ScalingAfterThreshold1 & "' is not valid (must be greater than " & CStr(Scaling) & ")." _
                ,, "Error", vbExclamation
        loop while true
    end if

    if SpeedSteps = 2 then
        ' Get the second threshold and sensitivity when speed is > threshold2
        do ' Get threshold2
            ThresholdText = ""
            if OldWindowsAccelName <> "" then _
                ThresholdText = vbNewLine & vbNewLine _
                    & "The second threshold for " & OldWindowsAccelName & " acceleration is " & OldWindowsThreshold2 & "."
            Threshold2 = InputBox( _
                "Enter the second acceleration threshold that you want." & vbNewLine _
                & vbNewLine _
                & "When the mouse is faster than this, the pointer speed will be further accelerated." & vbNewLine _
                & vbNewLine _
                & "See the !Threshold_Acceleration_ReadMe.txt file for guidance." _
                & ThresholdText, _
                "Pointer Speed Acceleration - MarkC Mouse Acceleration Fix", Threshold2)
            if Threshold2 = "" then WScript.Quit
            if IsNumeric(Threshold2) then
                NumThreshold = CInt(Threshold2)
                if NumThreshold > Threshold1 and CStr(NumThreshold) = Threshold2 then Threshold2 = NumThreshold : exit do
            end if
            WshShell.Popup "'" & Threshold2 & "' is not valid (must be greater than " & CStr(Threshold1) & ")." _
                ,, "Error", vbExclamation
        loop while true

        ThresholdText = vbNewLine _
            & vbNewLine _
            & "The pointer speed factor used by Windows 2000 at " & CStr(PointerSpeed) & "/11 is " _
            & CStr(4*MouseSensitivityFactor) & "." & vbNewLine _
            & "The pointer speed factor used by Windows 95+98 is 4."
        ScalingAfterThreshold2 = 4 * Scaling
        do ' Get the scaling (sensitivity) factor when faster than threshold1
            ScalingAfterThreshold2 = InputBox( _
                "Enter the pointer speed scaling (sensitivity) factor that you want" _
                & " when the mouse is faster than " & CStr(Threshold2) & "." _
                & ThresholdText, _
                "Pointer Speed Scaling - MarkC Mouse Acceleration Fix", ScalingAfterThreshold2)
            if ScalingAfterThreshold2 = "" then WScript.Quit
            if IsNumeric(ScalingAfterThreshold2) then _
                if CDbl(ScalingAfterThreshold2) > ScalingAfterThreshold1 and CDbl(ScalingAfterThreshold2) <= 16 then _
                    ScalingAfterThreshold2 = CDbl(ScalingAfterThreshold2) : exit do
            WshShell.Popup "'" & ScalingAfterThreshold2 & "' is not valid (must be greater than " _
                & CStr(ScalingAfterThreshold1) & ")." _
                ,, "Error", vbExclamation
        loop while true
    end if


    ' Compute the magic SmoothMouseCurve numbers
    if FixType = Windows10Fix or FixType = Windows8Fix then
        DPI = round(DPI*96/100)
        DPIFactor = B(Max(96,DPI)/120)
    else if FixType = Windows7Fix then
        DPI = round(DPI*96/100)
        DPIFactor = B(Max(96,DPI)/150)
    else
        DPIFactor = B(Max(60,RefreshRate)/Max(96,DPI))
    end if end if

    if SpeedSteps = 0 then

        ' No acceleration anywhere on the curve; original acceleration fix curve
        SmoothY = B(16*3.5)
        if MouseSensitivity  = 1 then SmoothY = SmoothY * 2 ' Ensure we
        if MouseSensitivity <= 2 then SmoothY = SmoothY * 2 ' have enough
        if Scaling > 3 then SmoothY = SmoothY * 2           ' bits of
        if Scaling > 6 then SmoothY = SmoothY * 2           ' precision
        if Scaling > 9 then SmoothY = SmoothY * 2           ' using
        if Scaling > 12 then SmoothY = SmoothY * 2          ' somewhat arbitrary
        if DPI > 144 and Scaling > 1 then SmoothY = SmoothY * 2 ' rules

        SmoothYPixels = BMult(BMult(SmoothY, DPIFactor), B(MouseSensitivity/10))
        SmoothX = B(SmoothYPixels/(B(Scaling)*3.5))
        ' Make sure the magic numbers give the exact result
        SmoothY = GetSmoothY(SmoothX, Scaling, 0, 0)
        ' if ActualScaling <> B(Scaling) now, then I don't care: close enough!

        SmoothX0 = 0
        SmoothX1 = SmoothX
        SmoothX2 = 2*SmoothX
        SmoothX3 = 3*SmoothX
        SmoothX4 = 4*SmoothX
        SmoothY0 = 0
        SmoothY1 = SmoothY
        SmoothY2 = 2*SmoothY
        SmoothY3 = 3*SmoothY
        SmoothY4 = 4*SmoothY

    else if SpeedSteps = 1 then

        ' Windows 2000 and earlier style 'step-up' acceleration with 1 threshold
        ' Mouse movement > the threshold uses higher scaling
        SmoothX0 = 0
        SmoothY0 = 0

        SmoothX1 = 0
        SmoothY1 = 0

        ' A segment for speeds lower than Threshold1
        SmoothX2 = round((Threshold1 + 0.75)/3.5 * 8) * &h2000
        SmoothY2 = GetSmoothY(SmoothX2, Scaling, 0, 0)

        SmoothX3 = 0
        SmoothY3 = 0

        ' A segment for speeds higher than Threshold1
        SmoothX4 = B(40)
        SmoothY4 = GetSmoothY(SmoothX4, ScalingAfterThreshold1, 0, 0)

    else if SpeedSteps = 2 then

        if FixType = Windows10Fix then WScript.Quit

        ' Windows 2000 and earlier style 'step-up' acceleration with 2 thresholds
        ' Mouse movement > threshold1 uses higher scaling, > threshold2 uses even higher scaling
        ' Check for a blog about this @

        ' A magic segment -1>0 with SmoothX=Threshold1 (and the same slope as segment 0>1)
        SmoothX0 = round((Threshold1 + 0.75)/3.5 * 8) * &h2000
        SmoothY0 = GetSmoothY(SmoothX0, ScalingAfterThreshold1, 0, 0)

        ' A segment for speeds higher than Threshold1 and lower than Threshold2
        SmoothX1 = round((Threshold2 + 0.75)/3.5 * 8) * &h2000
        SmoothY1 = GetSmoothY(SmoothX1, ScalingAfterThreshold1, SmoothX0, SmoothY0)

        SmoothX2 = 0
        SmoothY2 = 0

        ' A segment for speeds higher than Threshold2 (and lower than a magic high limit)
        if ScalingAfterThreshold2 <= 4 then
            SmoothX3 = B(&h900)
        else
            SmoothX3 = B(int(&h900 * 4 / ScalingAfterThreshold2))
        end if
        SmoothY3 = GetSmoothY(SmoothX3, ScalingAfterThreshold2, 0, 0)

        ' A magic segment with the scaling for speeds lower than Threshold1
        SmoothX4 = B(&h24920000) ' A bit less than 2^31/3.5
        SmoothY4 = -BDiv(BDiv(-B(Scaling), B(MouseSensitivity/10)), DPIFactor)

    else
        Err.Raise 0,, "Invalid value for SpeedSteps."
    end if end if end if

    MarkCStringX = CurveHex(SmoothX0) & "," & CurveHex(SmoothX1) & "," & CurveHex(SmoothX2) & "," & CurveHex(SmoothX3) & "," & CurveHex(SmoothX4)
    MarkCStringY = CurveHex(SmoothY0) & "," & CurveHex(SmoothY1) & "," & CurveHex(SmoothY2) & "," & CurveHex(SmoothY3) & "," & CurveHex(SmoothY4)

end function



' Convert to fixed point (n.16) binary
function B(n)
	B = int(&h10000 * n)
end function

' Fixed point (n.16) binary multiply
function BMult(m1, m2)
	BMult = int(m1 * m2 / &h10000)
end function

' Fixed point (n.16) binary divide
function BDiv(n, d)
	BDiv = int(&h10000 * n / d)
end function

' Calculate the SmoothY value that gives the desired TargetScaling
function GetSmoothY(ByRef SmoothX, TargetScaling, PreviousSmoothX, PreviousSmoothY)

	dim SmoothY, ExtraX, Slope, Intercept
	dim SmoothXMickeys, SmoothYPixels, PreviousSmoothXMickeys, PreviousSmoothYPixels
	PreviousSmoothXMickeys = BMult(PreviousSmoothX, B(3.5))
	PreviousSmoothYPixels = BMult(BMult(PreviousSmoothY, DPIFactor), B(MouseSensitivity/10))

	for ExtraX = 0 to 128 ' (Can go as high as +1100 and still be Mickeys < Threshold+1)
		SmoothY = -BDiv( _
			BDiv( _
				BMult( _
					-B(TargetScaling), _
					BMult(SmoothX + ExtraX, B(3.5)) - PreviousSmoothXMickeys) _
				+ -PreviousSmoothYPixels, _
				B(MouseSensitivity/10)), _
			DPIFactor)
		if ExtraX = 0 then GetSmoothY = SmoothY ' lock in at least the first value

		' Check if SmoothY & SmoothX are exactly the right slope & intercept
		SmoothYPixels = BMult(BMult(SmoothY, DPIFactor), B(MouseSensitivity/10))
		SmoothXMickeys = BMult(SmoothX + ExtraX, B(3.5))
		Slope = BDiv(SmoothYPixels - PreviousSmoothYPixels, SmoothXMickeys - PreviousSmoothXMickeys)
		Intercept = SmoothYPixels - BMult(Slope, SmoothXMickeys)

		if Slope = B(TargetScaling) and Intercept = 0 then
			' Exact match: return
			SmoothX = SmoothX + ExtraX
			GetSmoothY = SmoothY
			exit function
		end if
		' Bump SmoothX a little & try again (eventually a calculation is usually exact for normal input values)
	next

end function

' Convert number to registry REG_BINARY hex: format
function CurveHex(n)
	dim h, ch, i, high, low6
	high = int(n / &h1000000) ' 16^6: 6 hex digits
	low6 = n - &h1000000 * high
	h = right("00000" & hex(high), 6) & right("00000" & hex(low6), 6)
	ch = ""
	for i = 5 to 0 step -1
		ch = ch & mid(h, i*2+1, 2) & ","
	next
	CurveHex = ch & "00,00"
end function

function Max(n1, n2)
	if n1 > n2 then
		Max = n1
	else
		Max = n2
	end if
end function

function DWordFromBytes(i, Bytes)
	i = 8*i
	DWordFromBytes = 256*(256*(256*(256*(256*Bytes(i+5) + Bytes(i+4)) + Bytes(i+3)) + Bytes(i+2)) + Bytes(i+1)) + Bytes(i)
end function
