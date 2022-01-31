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
            objRegistry.SetStringValue hiveTarget, KeyPath, "iMeasure", "0"
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
        MarkCAccelCurveX = CurveStringToBinary(MarkCStringX)
        MarkCAccelCurveY = CurveStringToBinary(MarkCStringY)
    End If
End Function

Function CurveStringToBinary(str)
    dim uBinary : uBinary = Split(str, ",")
    For i = 0 To UBound(uBinary) : uBinary(i) = CByte("&H" & uBinary(i)) : Next
    CurveStringToBinary = uBinary
End Function
