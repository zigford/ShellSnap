Dim oFSO, oTsEnv, oShell, oComputer, oAdvertisement, oResID, arrResID()
Dim sSMSMP, sSMSSite, sTSMachID, sTSAdvert, sCollection, sSMSAcct, sSMSPwd
Dim i

Set oShell =  WScript.CreateObject("WScript.Shell")
Set oTsEnv = CreateObject("Microsoft.SMS.TsEnvironment")

sSMSMP = oTsEnv("_SMSTSMP")
sSMSSite = oTsEnv("_SMSTSSiteCode")
sTSMachID = oTsEnv("_SMSTSMachineName")
sTSAdvert = oTsEnv("_SMSTSAdvertID")
sSMSAcct = [Acct With Rights To Remove Direct Membership]
sSMSPwd = [Password For Account]

Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSMS = objLocator.ConnectServer(sSMSMP, "root/sms/site_" & sSMSSite, sSMSAcct, sSMSPwd)
Set colAdvertisement = objSMS.ExecQuery("Select * From SMS_Advertisement WHERE AdvertisementID='" & sTSAdvert & "'")

For Each oAdvertisement In colAdvertisement
  sCollection = oAdvertisement.CollectionID
Next

Set colComputer = objSMS.ExecQuery("Select * From SMS_R_System WHERE Name='" & sTSMachID & "'")
i = Null
For Each oComputer In colComputer
  If IsNull(i) Then
    i = 0
  Else
    i = i + 1
  End If
Next

Redim arrResID(i)
For Each oComputer In colComputer
  arrResID(i) = oComputer.ResourceID
  If i > 0 Then
    i = i - 1
  Else
    Exit For
  End If
Next

For Each oResID In arrResID
  RemDirMember oResID, sCollection
Next

Function RemDirMember(ResID, CollID)
  Set instCollection = objSMS.Get("SMS_Collection.CollectionID='" & CollID  & "'" )
  Set instDirectRule = objSMS.Get("SMS_CollectionRuleDirect").SpawnInstance_
  instDirectRule.ResourceID = ResID
  instCollection.DeleteMembershipRule instDirectRule
  instCollection.RequestRefresh True
End Function