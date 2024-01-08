$message = "Help me";

#time in s to close the message
Invoke-WmiMethod -Class win32_process -Name create -ArgumentList  "c:\windows\system32\msg.exe * /time:5 $message" 