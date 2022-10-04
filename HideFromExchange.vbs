oN ERROR GOTO 0
' Error Resume Next
DIM strUserDN
DIM iMaxComputersSaved

dim mu_path, index_1  

Set wshArguments = WScript.Arguments
Set objUser = GetObject(wshArguments(0))


				If (objUser.msExchHideFromAddressLists <>"") then

			objUser.PutEx 1, "msExchHideFromAddressLists", ""
				WScript.Echo "Появился в списках GAL"
				Else

			objUser.put "msExchHideFromAddressLists", True
				WScript.Echo "Убран их списков GAL"
				End If
				
				
				
'Set the info (save) into the attribute
objUser.SetInfo


