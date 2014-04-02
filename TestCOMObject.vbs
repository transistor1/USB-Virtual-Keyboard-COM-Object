'This function will select all in a text application (like Notepad),
'then type "The Quick Brown Fox Jumps Over The Lazy Dogs"
'and press enter.
Sub test()
    Dim p
    Set p = CreateObject("FieldEffect.USBKeyboard")
    p.Connect "COM3"
    For i = 1 To 10
        'Special character \s141 is Ctrl-A, or Select All
        p.SendKey "\s141"
        'Send a string
        p.SendString "The Quick Brown Fox Jumps Over the Lazy Dogs"
        'Send a hex key; in this case a carriage return
        p.SendKey "\x0d"
    Next
    p.Disconnect
End Sub

Sub Wait(Seconds)
	t = Timer
	Do While Timer - t < Seconds
		'Wait a few sec
	Loop
End Sub

'this test opens the Windows Task Manager, then runs notepad
'and types some things into the notepad window
Sub test2()
	Dim p
    Set p = CreateObject("FieldEffect.USBKeyboard")
    p.Connect "COM3"
	'p.SendKey "\sB2"
	p.SendKeyCode &H07&,&H4C&,&H05&
	Wait 3
	'p.SendString "\s999" '"\s999" => [0x07,0x17,0x04],	 #alt-t
	
	p.SendKeyCode &H07, &H17, &H04 'Alt-T to Bring up task manager
	Wait 1
	p.SendKeyCode &H07&, &H09&, &H04& 'Alt-F for File menu
	Wait 1
	p.SendString "r" 'Run
	p.SendKey "\x0d" 'Enter
	Wait 2
	p.SendKey "\s141" 'Select all (Ctrl-A)
	p.SendString "notepad.exe"
	p.SendKey "\x0d" 'Enter
	
	For i = 1 To 20
		p.SendKey "\x09" 'Tab
		p.SendString "The Quick Brown Fox Jumps Over The Lazy Dog " & i & " Times!"
	Next
	p.Disconnect
End Sub



Call test2()
