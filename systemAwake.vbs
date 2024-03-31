Set objShell = CreateObject("WScript.Shell")
Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objService = objLocator.ConnectServer(".", "root\cimv2")

Dim startTime
startTime = Now ' Store the start time

Dim screenOffTime
screenOffTime = DateAdd("n", 120, Now) ' Add 30 minutes to the current time (example)

' Prevent the screen from turning off
objShell.Run "powercfg -change -monitor-timeout-ac 0", 0, True

Do
    elapsedTime = DateDiff("n", startTime, Now) ' Calculate elapsed time in minutes
    If elapsedTime >= 3 Then ' Check if 3 minutes have passed
        MoveMouse() ' If 3 minutes have passed, move the mouse
        startTime = Now ' Reset the start time
    End If
    
    ' Check if it's time to turn off the screen
    If Now >= screenOffTime Then
        ' Restore default screen timeout
        objShell.Run "powercfg -change -monitor-timeout-ac 15", 0, True ' Set timeout to 15 minutes
        Exit Do
    End If
    
    WScript.Sleep 1000 ' Adjust the delay as necessary
Loop

Sub MoveMouse()
    ' Subroutine to move the mouse cursor
    screenWidth = objService.ExecQuery("SELECT * FROM Win32_DesktopMonitor").ItemIndex(0).ScreenWidth
    screenHeight = objService.ExecQuery("SELECT * FROM Win32_DesktopMonitor").ItemIndex(0).ScreenHeight
    Randomize
    randomX = Int((screenWidth - 1 + 1) * Rnd + 1)
    randomY = Int((screenHeight - 1 + 1) * Rnd + 1)
    objShell.SendKeys "{ESC}" ' Release any keys that might be pressed
    objShell.SendKeys "% " & randomX & " " & randomY ' % indicates Alt key to ensure proper positioning
End Sub
