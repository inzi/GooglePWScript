''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Versions
'
' Version: 1.8.1
' source: inzi.com
'
' Change Log
'
' v1.0 
' Original
'
' v1.5  10/23/2013
' Added iSlowConnectionFactor, an easy way to tweak speed of script
' Updated default path to Chrome
' Added 2nd tab to account for lost password link at line 109
'
'
' v1.7 11/18/2014
' Changed path to Chrome.exe to just be "chrome.exe" Full path seems to be an issue for some users
' Increased delay for google login
' apparently - the google login is fluxing - added commenting so people can modify the dialog navigation. 
' Fixed the final password reset step so it tabs twice
' Added detailed of comments so people know what it's doing
' 
'
' v1.8 11/18/2014
' change the login URL so that the email field is editable on each login
' 
' v1.8.1 11/18/2014
' Bug - left out a quote
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' When launching this, you might pipe the output to a file in case your computer turns off, so you know what the password is
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declare Variables & Objects
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim oShell
Dim oAutoIt

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialise Variables & Objects
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set oShell = WScript.CreateObject("WScript.Shell")
Set oAutoIt = WScript.CreateObject("AutoItX3.Control")

WScript.Echo "This script will reset your google password 100 times so you can use an old password."



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You should only edit value after this
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Your Google username (email address)
sUN = "username@gmail.com" 

' Replace with the Current password to your google account
curPW = "????????" 

' Replace with the final password you want assigned to the account. The one you want to set the account's password *back to*. 
oldPW = "????????" 

' Is it going to fast? You can slow it down by adjusting this value.
' If you set it to 2, it will run twice as slow
' So if it is entering data into the wrong fields, try increasing this.
' It might help.
iSlowConnectionFactor = 1

' If your password has a quote in it ("), then use "" in its place.
' For example, let's say your password was 
' MyPass"word!-55
'
' The proper VBScript way to put that into a variable would look like this
' curPW = "MyPass""word!-55"
'
' See Microsoft's website for more detail

' Where is the Chrome executable? Replace this with its location.
'     Point app to Chrome Manually
'     An easy way to find this is to right click the Chrome shortcut and copy the value in Target.
'     Click Start, type Chrome, right click Google Chrome, click Properties, copy *everything* in Target, and put it here.

' This example path is for 64 bit windows
' You might need the full path: ChromeEXE = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
ChromeEXE = "chrome.exe"


' This example path is for 32 bit windows
' "C:\Program Files\Google\Chrome\Application\chrome.exe" 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You should not have to edit anything after this
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Start of Script
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'Some of this code uses the AutoIT com object. See their documentation for more details.
' https://www.autoitscript.com/site/autoit/
' you have to install that to use this.


oShell.Run ChromeEXE, 1, False

' Wait for the Google Chrome window to become active
oAutoIt.WinWaitActive "New Tab - Google Chrome", ""			'open chrome and wait for the window to appear
oAutoIt.Sleep 3000

WScript.Echo "Entering Loop"
' Enter the loop, change the password 99 more times
tCurPw = curPW

For x = 1 To 99
    WScript.Echo "Step " & x
    WScript.Echo "Current PW: " & tCurPw
    tNewPW = curPW & x
    WScript.Echo "Setting the password to: " & tNewPW			' output the password in case it crashes the user can see what it is.
    
    GLogin sUN, tCurPw							' log into google with current password				
    GEditPW								' load the edit password page, old password should have focus after it loads.
    oAutoIt.Send tCurPw & "{TAB}{TAB}"					' enter the old (last) password and tab TWICE (once out of field, and once past password reset link)
    oAutoIt.Sleep 250 * iSlowConnectionFactor 				' waits x ms times slow connection
    oAutoIt.Send tNewPW & "{TAB}"					' type new password
    oAutoIt.Sleep 250 * iSlowConnectionFactor 				' waits x ms times slow connection
    oAutoIt.Send tNewPW & "{TAB}"					' types the new password again then tab (should highlight "change password button")
    oAutoIt.Sleep 250 * iSlowConnectionFactor 				' waits x ms times slow connection
    oAutoIt.Send "{ENTER}"						' Hit enter to submit the password reset.
    oAutoIt.Sleep 3000 * iSlowConnectionFactor 				' waits x ms times slow connection
    
    tCurPw = tNewPW							' updates the current password.

    GLogout								' logs out of google - it has to do this for password change to stick.

									' At this point, the process should repeat.

Next 

WScript.Echo "Final Change"						' Last time
GLogin sUN, tCurPw							' log into google with current password.
GEditPW									' load the edit password page

oAutoIt.Send tCurPw & "{TAB}{TAB}"					' enter the old (last) password and tab TWICE (once out of field, and once past password reset link)			
oAutoIt.Sleep 250 * iSlowConnectionFactor 				' waits x ms times slow connection
oAutoIt.Send oldPW & "{TAB}"						' enter the password, used password and hit tab
oAutoIt.Sleep 250 * iSlowConnectionFactor				' waits x ms times slow connection
oAutoIt.Send oldPW & "{TAB}"						' enter the password, used password and hit tab
oAutoIt.Sleep 250 * iSlowConnectionFactor 				' waits x ms times slow connection
oAutoIt.Send "{ENTER}"							' Submit the password reset
oAutoIt.Send "https://www.google.com/accounts/Logout{ENTER}"		' Log out to make the password stick
oAutoIt.Sleep 2000 * iSlowConnectionFactor 				' waits x ms times slow connection
WScript.Echo "Password reset"

WScript.Quit

' Script complete.


Function GLogin(un, pw) ' Opens the Google Login page, enters the supplied Username (un) and Password (pw), and presses Enter.
    WScript.Echo "Logging in: " & un & ", " & pw
    oAutoIt.Send "!d"							' This goes to the address bar
    oAutoIt.Sleep 250 * iSlowConnectionFactor 				' waits x ms times slow connection
    'oAutoIt.Send "https://accounts.google.com/Login{ENTER}"		' types this url and hits enter. Upon load, email field should have focus
    oAutoIt.Send "https://accounts.google.com/ServiceLogin?Email=%22%22{ENTER}" ' types this url and hits enter. Upon load, email field should have focus. Email param makes it empty.
    oAutoIt.Sleep 2000 * iSlowConnectionFactor 				' waits x ms times slow connection
    oAutoIt.Send un & "{TAB}"						' types username and hits tab
    oAutoIt.Sleep 250 * iSlowConnectionFactor 				' waits x ms times slow connection
    oAutoIt.Send pw & "{ENTER}"						' types password and hits enter
    oAutoIt.Sleep 5000 * iSlowConnectionFactor 				' waits x ms times slow connection.

End Function

Function GEditPW() ' Opens the Google Change Password web page
    oAutoIt.Send "!d"							' go to address bar
    oAutoIt.Sleep 250 * iSlowConnectionFactor 				' waits x ms times slow connection
    oAutoIt.Send "https://accounts.google.com/b/0/EditPasswd{ENTER}"	' go the edit password page
    oAutoIt.Sleep 4000 * iSlowConnectionFactor 				' waits x ms times slow connection
End Function

Function GLogout() ' Logs out from google. This is necessary for the password change to take effect. Trust me, I tried to do it without logging out. No luck.
    WScript.Echo "Logging out"
    oAutoIt.Send "!d"							' this goes to the address bar
    oAutoIt.Sleep 250 * iSlowConnectionFactor 				' waits x ms times slow connection
    oAutoIt.Send "https://www.google.com/accounts/Logout{ENTER}"	' This logs out of google. 
    oAutoIt.Sleep 5000 * iSlowConnectionFactor 				' waits x ms times slow connection

End Function
