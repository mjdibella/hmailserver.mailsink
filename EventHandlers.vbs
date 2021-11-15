'   Sub OnClientConnect(oClient)
'   End Sub

'   Sub OnSMTPData(oClient, oMessage)
'   End Sub

'	Sub OnAcceptMessage(oClient, oMessage)
'	End Sub

'   Sub OnDeliveryStart(oMessage)
'   End Sub

'   Sub OnDeliverMessage(oMessage)
'   End Sub

'   Sub OnBackupFailed(sReason)
'   End Sub

'   Sub OnBackupCompleted()
'   End Sub

'   Sub OnError(iSeverity, iCode, sSource, sDescription)
'   End Sub

'   Sub OnDeliveryFailed(oMessage, sRecipient, sErrorMessage)
'   End Sub

'   Sub OnExternalAccountDownload(oFetchAccount, oMessage, sRemoteUID)
'   End Sub

Sub SendNDR(oMessage)
	set oMailServer = GetApplicationObject	
	sNDRBody = oMailServer.Settings.ServerMessages.ItemByName("SEND_FAILED_NOTIFICATION").Text
	sNDRBody = Replace(sNDRBody, "%MACRO_SENT%", oMessage.Date)
	sNDRBody = Replace(sNDRBody, "%MACRO_SUBJECT%", oMessage.Subject)
	sNDRBody = Replace(sNDRBody, "%MACRO_TO%", oMessage.To)
	sNDRBody = Replace(sNDRBody, "%MACRO_FROM%", oMessage.From)
	sEnvelopeRecipients = "  " & Replace(oMessage.HeaderValue("X-Envelope-Recipients"), ",", "  " & vbCRLF)
	sReason = "One or more of following addresses was unreachable: " & vbCRLF & vbCRLF & sEnvelopeRecipients
	sNDRBody = Replace(sNDRBody, "%MACRO_RECIPIENTS%", sReason)
	set oNDR = CreateObject("hMailServer.Message")
	oNDR.From = "mailer-daemon@" & oMailServer.Settings.HostName
	oNDR.FromAddress = "mailer-daemon@" & oMailServer.Settings.HostName
	oNDR.Subject = "Message undeliverable: " & oMessage.Subject
	sReturnPath = CleanAddress(oMessage.HeaderValue("Return-Path"))
	if Len(sReturnPath) = 0 then
		sReturnPath = CleanAddress(oMessage.From)
	end if
	oNDR.AddRecipient "", sReturnPath
	oNDR.HeaderValue("To") = sReturnPath
	oNDR.Body = sNDRBody
	oNDR.Save
end sub

function GetApplicationObject
	const HMS_USER = "postmaster@domain"
	const HMS_PASS = "password"
	set oApplicationObject = CreateObject("hMailServer.Application")
	oApplicationObject.Authenticate HMS_USER, HMS_PASS
	set GetApplicationObject = oApplicationObject
end function

function CleanAddress(sAddress)
	dim i
  i = InStrRev(sAddress, "<")
  if i > 0 then
	sAddress = Mid(sAddress, i + 1)
	i = InStr(sAddress, ">")
	if i > 0 then
		sAddress = Mid(sAddress, 1, i - 1)
	end if
	sAddress = CleanAddress(sAddress)
  end if
  CleanAddress = lcase(sAddress)
end function
