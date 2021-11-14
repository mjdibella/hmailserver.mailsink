<%
	const POSTMASTER = "hostmaster@dibella.net"
	const INFORMATION = 4
	const HMS_USER = "postmaster@sns.otamdm.net"
	const HMS_PASS = "kc106N12T0kHBr6g"
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		aPostBody = Request.BinaryRead(Request.TotalBytes)
		for i = 1 To LenB(aPostBody)
			sPostBody = sPostBody & Chr(AscB(MidB(aPostBody, i, 1)))
		next
		set oJSONParser = CreateObject("MSScriptControl.ScriptControl")
		oJSONParser.Language = "JScript"
		sMessageType = Request.ServerVariables("HTTP_X_AMZ_SNS_MESSAGE_TYPE")
		if Len(sMessageType) > 0 then
			set oHMS = CreateObject("hMailServer.Application")
			on error resume next
			oHMS.Authenticate HMS_USER, HMS_PASS
			if sMessageType = "Notification" then
				ProcessReceivedMessage
			elseif sMessageType = "SubscriptionConfirmation" then
				SendSubscriptionConfirmation
			else
				NotifyUnknownMessageType
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
		for each sEnvelopeRecipient in oRecipients
			oMessage.AddRecipient "", sEnvelopeRecipient
			if Len(sEnvelopeRecipients) > 0 then
				sEnvelopeRecipients = sEnvelopeRecipients & "," & sEnvelopeRecipient
			else 
				sEnvelopeRecipients = sEnvelopeRecipient
			end if
		next
		oMessage.HeaderValue("To") = sOriginalTo
		oMessage.HeaderValue("CC") = sOriginalCC
		oMessage.HeaderValue("X-Envelope-Recipients") = sEnvelopeRecipients
		oMessage.Save
		Response.Write "Message dispatched to " & sEnvelopeRecipients
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
		oMessage.AddRecipient "", POSTMASTER
		oMessage.Body = sMessageBody
		oMessage.Save
		Response.Write "Confirmation sent"
	end sub
	
	sub NotifyUnknownMessageType
		set oMessage = CreateObject("hMailServer.Message")
		oMessage.From = "Amazon Web Services <noreply@aws.amazon.com>"
		oMessage.FromAddress = "noreply@aws.amazon.com"
		oMessage.Subject = "Unknown SNS Message Received"
		oMessage.AddRecipient "", POSTMASTER
		oMessage.Body = Request.ServerVariables("ALL_HTTP") & vbCRLF & sPostBody
		oMessage.Save
		Response.Write "Unknown message notification sent"
	end sub
%>
