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
	aEnvelopeRecipients = Split(oMessage.HeaderValue("X-Envelope-Recipients"), ",")
	for each sEnvelopeRecipient in aEnvelopeRecipients
		if not IsDeliverable(sEnvelopeRecipient) then
			sEnvelopeRecipients = sEnvelopeRecipients & "    " & sEnvelopeRecipient & vbCRLF
		end if
	next
	sNDRBody = Replace(sNDRBody, "%MACRO_RECIPIENTS%", sEnvelopeRecipients)
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

function IsDeliverable(sAddress)
	IsDeliverable = false
	sOriginalAddress = sAddress
	sUserName = UsernameFromAddress(sAddress)
	sOriginalDomain = DomainFromAddress(sAddress)
	set oMailServer = GetApplicationObject
	on error resume next
	set oRoute = oMailServer.Settings.Routes.ItemByName(DomainFromAddress(sAddress))
	set oDomain = oMailServer.Domains.ItemByName(DomainFromAddress(sAddress))
	on error goto 0
	if not IsObject(oDomain) then
		' find domain by alias
		if oMailServer.Domains.Count > 0 then
			for i = 0 to oMailServer.Domains.Count - 1
				if oMailServer.Domains(i).DomainAliases.Count > 0 then
					for j = 0 to oMailServer.Domains(i).DomainAliases.Count - 1
						if oMailServer.Domains(i).DomainAliases(j).AliasName = DomainFromAddress(sAddress) then
							set oDomain = oMailServer.Domains.ItemByDBID(oMailServer.Domains(i).DomainAliases(j).DomainID)
							exit for
						end if
					next
					if IsObject(oDomain) then
						exit for
					end if
				end if
			next
		end if
		if IsObject(oDomain) then
			sAddress = sUserName & "@" & oDomain.Name
		end if
	end if
	if not IsDeliverable then
		if (not IsObject(oDomain)) and (not IsObject(oRoute)) then
			' non-authorative domain, deliverable via relay
			IsDeliverable = true
		else
			if IsObject(oDomain) then
				on error resume next
				set oAccount = oDomain.Accounts.ItemByAddress(sAddress)
				on error goto 0
				if IsObject(oAccount) then
					if oAccount.Active then
						IsDeliverable = true
					end if
				else
					on error resume next
					set oAlias = oDomain.Aliases.ItemByName(sAddress)
					on error goto 0
					if IsObject(oAlias) then
						IsDeliverable = true
					else
						on error resume next
						set oDistributionList = oDomain.DistributionLists.ItemByAddress(sAddress)
						on error goto 0
						if IsObject(oDistributionList) then
							IsDeliverable = true
						end if
					end if
				end if
			end if
			if not IsDeliverable then
				on error resume next
				set oRoute = oMailServer.Settings.Routes.ItemByName(sDomainName)
				on error goto 0
				if IsObject(oRoute) then
					if oRoute.AllAddresses then
						IsDeliverable = true
					else
						if oRoute.Addresses.Count > 0 then 
							for i = 0 to oRoute.Addresses.Count - 1
								if oRoute.Addresses.Item(i).Address = sOriginalAddress then
									IsDeliverable = true
								end if
							next
						end if
					end if
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


