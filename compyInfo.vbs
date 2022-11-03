strComputer = "."
   
Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\directory\LDAP")
Set objComputers = objWMI.ExecQuery("SELECT * FROM ds_computer where ((ds_name LIKE 'XXXXX%') OR (ds_name LIKE 'xxxxx%')) ")
   
if objComputers.Count = 0 then
   Wscript.Echo "Nothin' here, yo."
else
   for each objComputer in objComputers
        If Not IsNull(objComputer.ds_description) Then 
        	For i=LBound(objComputer.ds_description) to UBound(objComputer.ds_description)
                WScript.Echo objComputer.ds_name & "~" & objComputer.ds_description(i)
        next
        Else WScript.Echo objComputer.ds_name & "~"
	end if
      WScript.Echo ""
   next
end if