<%
	const REGKEY = "HKEY_LOCAL_MACHINE\SOFTWARE\HMSMailsink\"
 	set oRegistry = CreateObject("WScript.Shell")
	sHMSUser = oRegistry.RegRead(REGKEY & "HMSUser")
	sHMSPass = oRegistry.RegRead(REGKEY & "HMSPass")
	sPostmaster = oRegistry.RegRead(REGKEY & "Postmaster")
	set oRegistry = Nothing
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		aPostBody = Request.BinaryRead(Request.TotalBytes)
		for i = 1 To LenB(aPostBody)
			sPostBody = sPostBody & Chr(AscB(MidB(aPostBody, i, 1)))
		next
		set oJSONParser = CreateObject("MSScriptControl.ScriptControl")
		oJSONParser.Language = "JScript"
		sMessageType = Request.ServerVariables("HTTP_X_AMZ_SNS_MESSAGE_TYPE")
		if Len(sMessageType) > 0 then
			set oMailServer = CreateObject("hMailServer.Application")
			on error resume next
			set oAuthUser = oMailServer.Authenticate(sHMSUser, sHMSPass)
			on error goto 0 
			if oAuthUser is nothing then
				Response.Status = "503 Service unavailable"
				Response.Write "Service unavailable"
			else
				if sMessageType = "Notification" then
					ProcessReceivedMessage
				elseif sMessageType = "SubscriptionConfirmation" then
					SendSubscriptionConfirmation
				else
					NotifyUnknownMessageType
				end if
			end if
		else
			Response.Status = "400 Bad Request"
			Response.Write "Bad request"
		end if
	else	
		Response.Status = "405 Method Not Allowed"
		Response.Write "Method not allowed"
	end if
	
	sub ProcessReceivedMessage
		on error resume next
		set oPostJson = oJSONParser.Eval("(" & sPostBody & ")")
		sMessageData = oPostJson.Content
		on error goto 0
		set oMessage = CreateObject("hMailServer.Message")
		sFileName = oMessage.FileName
		set objFS = Server.CreateObject("Scripting.FileSystemObject") 
		set fEmlFile = objFS.CreateTextFile(sFileName)
		fEmlFile.Write sMessageData
		fEmlFile.Close
		oMessage.RefreshContent
		sOriginalTo = oMessage.HeaderValue("To")
		sOriginalCC = oMessage.HeaderValue("CC")
		oMessage.ClearRecipients
		set oRecipients = oPostJson.Receipt.Recipients
		bValidRecipient = false
		for each sEnvelopeRecipient in oRecipients
			if IsDeliverable(sEnvelopeRecipient) then
				bValidRecipient = true
				oMessage.AddRecipient "", sEnvelopeRecipient
				if Len(sEnvelopeRecipients) > 0 then
					sEnvelopeRecipients = sEnvelopeRecipients & "," & sEnvelopeRecipient
				else 
					sEnvelopeRecipients = sEnvelopeRecipient
				end if
			else
				sFailedRecipients = sFailedRecipients & "    " & sEnvelopeRecipient & vbCRLF
			end if
		next
		if bValidRecipient then
			oMessage.HeaderValue("To") = sOriginalTo
			oMessage.HeaderValue("CC") = sOriginalCC
			oMessage.HeaderValue("X-Envelope-Recipients") = sEnvelopeRecipients
			oMessage.Save
			Response.Write "Message dispatched to " & sEnvelopeRecipients
		else
			Response.Write "No valid recipients to dispatch."
		end if
		if Len(sFailedRecipients) > 0 then
			SendNDR oMessage, sFailedRecipients
		end if
	end sub

	sub SendSubscriptionConfirmation
		on error resume next
		set oPostJson = oJSONParser.Eval("(" & sPostBody & ")")
		sMessageBody = oPostJson.Message & vbCRLF & vbCRLF & oPostJson.SubscribeURL
		on error goto 0
		set oMessage = CreateObject("hMailServer.Message")
		oMessage.From = "Amazon Web Services <noreply@aws.amazon.com>"
		oMessage.FromAddress = "noreply@aws.amazon.com"
		oMessage.Subject = "SNS Subscription Confirmation"
		oMessage.AddRecipient "", sPostmaster
		oMessage.Body = sMessageBody
		oMessage.Save
		Response.Write "Confirmation sent"
	end sub
	
	sub NotifyUnknownMessageType
		set oMessage = CreateObject("hMailServer.Message")
		oMessage.From = "Amazon Web Services <noreply@aws.amazon.com>"
		oMessage.FromAddress = "noreply@aws.amazon.com"
		oMessage.Subject = "Unknown SNS Message Received"
		oMessage.AddRecipient "", sPostmaster
		oMessage.Body = Request.ServerVariables("ALL_HTTP") & vbCRLF & sPostBody
		oMessage.Save
		Response.Write "Unknown message notification sent"
	end sub
	
	Sub SendNDR(oMessage, sFailedRecipients)
		sNDRBody = oMailServer.Settings.ServerMessages.ItemByName("SEND_FAILED_NOTIFICATION").Text
		sNDRBody = Replace(sNDRBody, "%MACRO_SENT%", oMessage.Date)
		sNDRBody = Replace(sNDRBody, "%MACRO_SUBJECT%", oMessage.Subject)
		sNDRBody = Replace(sNDRBody, "%MACRO_TO%", oMessage.To)
		sNDRBody = Replace(sNDRBody, "%MACRO_FROM%", oMessage.From)
		sNDRBody = Replace(sNDRBody, "%MACRO_RECIPIENTS%", sFailedRecipients)
		set oNDR = CreateObject("hMailServer.Message")
		oNDR.From = sPostmaster
		oNDR.FromAddress = sPostmaster
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

	function IsDeliverable(sAddress)
		IsDeliverable = false
		on error resume next
		set oDomain = oMailServer.Domains.ItemByName(DomainFromAddress(sAddress))
		on error goto 0
		if IsObject(oDomain) then
			on error resume next
			set oAccount = oDomain.Accounts.ItemByAddress(sAddress)
			on error goto 0
			if IsObject(oAccount) then
				if oAccount.Active then
					IsDeliverable = true
				end if
			end if
		end if
		if not IsDeliverable then
			on error resume next
			set oRoute = oMailServer.Settings.Routes.ItemByName(DomainFromAddress(sAddress))
			on error goto 0
			if IsObject(oRoute) then
				if oRoute.AllAddresses then
					IsDeliverable = true
				else
					if oRoute.Addresses.Count > 0 then 
						for i = 0 to oRoute.Addresses.Count - 1
							if oRoute.Addresses.Item(i).Address = sAddress then
								IsDeliverable = true
							end if
						next
					end if
				end if
			end if
		end if
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

	function DomainFromAddress(sAddress)
		dim aTemp
		aTemp = Split(sAddress, "@")
		if UBound(aTemp) > 0 then
			DomainFromAddress = aTemp(1)
		else
			DomainFromAddress = sAddress
		end if
	end function

	function UsernameFromAddress(sAddress)
		dim aTemp
		aTemp = Split(sAddress, "@")
		if UBound(aTemp) > 0 then
			UsernameFromAddress = aTemp(0)
		else
			UsernameFromAddress = sAddress
		end if
	end function
%>
