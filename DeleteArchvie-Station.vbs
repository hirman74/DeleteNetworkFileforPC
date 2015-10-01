
PcName = Array ("mft-occ-1", "mft-occ-2", "mft-occ-3", "mft-occ-5", "mft-occ-6", "mft-occ-7", "mft-occ-8", "mft-occ-9", "mft-kcd-1", "mft-kcd-2", "mft-cdt-1", "mft-dbg-1", "mft-bbs-1", "mft-epn-1", "mft-pmn-1", "mft-nch-1", "mft-sdm-1", "mft-mbt-1", "mft-dkt-1", "mft-pyl-1", "mft-mps-1", "mft-tsg-1", "mft-bly-1", "mft-ser-1", "mft-lrc-1", "mft-bsh-1", "mft-mrm-1", "mft-cdt-1", "mft-bkb-1", "mft-btn-1", "mft-frr-1", "mft-hlv-1", "mft-bnv-1", "mft-onh-1", "mft-krg-1", "mft-hpv-1","mft-ppj-1", "mft-lbd-1", "mft-tlb-1", "mft-hbf-1", "mft-bft-1", "mft-mrb-1")

ON ERROR RESUME NEXT

strDriveLetter = "M:" 
strUser = "hirman"
strPassword = "password"
strProfile = "false"

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim objNetwork
Dim objFSO
Dim objShell
Set objNetwork = WScript.CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("WScript.Shell") 
Dim StringCMD

'vbscript method
For Each Item In PcName

	strRemotePath = "\\"&Item&"\c$" 
	Set AllDrives = objNetwork.EnumNetworkDrives() 
	AlreadyConnected = False 

	For i = 0 To AllDrives.Count - 1 Step 2 
	If AllDrives.Item(i) = strDriveLetter Then AlreadyConnected = True 
	Next 

	If AlreadyConnected = True then 
		objNetwork.RemoveNetworkDrive strDriveLetter
		'Wscript.Sleep 1000
		objNetwork.MapNetworkDrive strDriveLetter, strRemotePath, strProfile, strUser, strPassword

	Else 
		objNetwork.MapNetworkDrive strDriveLetter, strRemotePath, strProfile, strUser, strPassword
	End if 

	StringCMD = "cmd /c del M:\transactive\log\*Archive* "
	objShell.Run StringCMD, 0 , True

	'StringCMD = "cmd /c mkdir "& Item
	'objShell.Run StringCMD, 0 , True


	Set AllDrives = objNetwork.EnumNetworkDrives() 
	AlreadyConnected = False 
	For i = 0 To AllDrives.Count - 1 Step 2 
	If AllDrives.Item(i) = strDriveLetter Then AlreadyConnected = True 
	Next 
	If AlreadyConnected = True then 
	intReturn = objShell.Popup("Releasing \\"& Item &"\c$ from M:Drive.", 2, "Delete Archive Logs", vbOKOnly)
	objNetwork.RemoveNetworkDrive strDriveLetter
	Wscript.Sleep 1000
	End if 

Next

'DOS batch pushD and popD method
For Each Item In PcName
	StringCMD = "cmd /c net use \\" & Item & "\c$ /user:" & strUser & " " & strPassword & " /persistent:yes"
	objShell.Run StringCMD, 0 , True
	StringCMD = "cmd /c pushd \\"& Item & "\c$\transactive\log\"
	objShell.Run StringCMD, 0 , True
	StringCMD = "cmd /c del .\*Archive*"
	objShell.Run StringCMD, 0 , True
	StringCMD = "cmd /c popd"
	objShell.Run StringCMD, 0 , True	
Next

Set objNetwork = Nothing
Set objFSO = Nothing
Set objShell = Nothing
