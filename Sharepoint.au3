#comments-start
SPGetTenantRealmID(ByRef $ojbect, $string)
Function obtains Bearer realm aka Tenat ID
and ResourceID which in turn are used in order
to obtain the security token
Return value: one-dimensional array with two elements
1st element is number of tokens returned
2nd element is Bearer/TenantID
3rd element is ResourceID
#comments-end
Func SPGetTenantRealmID(ByRef $_oHTTP, $_sSiteUrl)
	Local $__sResponseHeader
	Local $__aRxMatch
	Local $__aReturnValues = [0,"",""] ;
	If StringRight($_sSiteUrl,1) == "/" Then
		$_sSiteUrl = $_sSiteUrl & "_vti_bin/client.svc"
	Else
		$_sSiteUrl = $_sSiteUrl & "/_vti_bin/client.svc"
	EndIf
	$_oHTTP.open("GET", $_sSiteUrl, False)
	$_oHTTP.setRequestHeader("Authorization", "Bearer")
	$_oHTTP.send()
	If $_oHTTP.Status == 401 Then ; 401 is expected at this stage
		$__sResponseHeader = $_oHTTP.getResponseHeader("WWW-Authenticate")
		$__aRxMatch = StringRegExp($__sResponseHeader,"realm=""[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}",$STR_REGEXPARRAYMATCH)
		If UBound($__aRxMatch) >= 1 Then
			; Sould really match only once
			$__aReturnValues[1] = StringRight($__aRxMatch[0],StringLen($__aRxMatch[0]) - 7)
			$__aReturnValues[0] = $__aReturnValues[0] + 1
		Else
			$__aReturnValues[0] = $__aReturnValues[0] + 0
		EndIf
		$__aRxMatch = StringRegExp($__sResponseHeader,"client_id=""[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}",$STR_REGEXPARRAYMATCH)
		If UBound($__aRxMatch) >= 1 Then
			; Should really match only once
			$__aReturnValues[2] = StringRight($__aRxMatch[0],StringLen($__aRxMatch[0]) - 11)
			$__aReturnValues[0] = 2
		Else
			$__aReturnValues[0] = $__aReturnValues[0] + 1
		EndIf
		Return $__aReturnValues
	Else
		$__aReturnValues[0] = 0
		Return $__aReturnValues
	EndIf
EndFunc

#comments-start
SPGetSecurityToken(ByRef $object, $string, $string)
Function obtains Security token used in the subsequent
operations
Arguments:
$_oHTTP -> object. winhttp.winhttprequest or similar object
$_sTenantDomainName -> string. e.g. volvogroup.sharepoint.com
$_sTenantID -> string. TenantID aka RealmID. Obtained by SPGetTenantRealmID()
$_sResourceID -> string. Obtained by SPGetTenantRealmID()
$_sClientID -> string. Generated by sharepoint at trusted app registration
$_sClientSecret -> string. Generated by sharepoint at trusted app registration
Return values:
If http request is successfull and token is found string containing token is returned.
If http request fails than http status as number is returned
If http request succeeds but function fails to find token 1 is returned
#comments-end
Func SPGetSecurityToken(ByRef $_oHTTP, $_sTenantDomainName, $_sTenantID, $_sResourceID, $_sClientID, $_sClientSecret)
	If StringLeft($_sTenantDomainName,1) <> "/" Then
		$_sTenantDomainName = "/" & $_sTenantDomainName
	EndIf
	Local $__aStringSplit1
	Local $__aStringSplit2
	Local $__sToken
	Local $__sAuthUrl1 = "https://accounts.accesscontrol.windows.net/"
	Local $__sAuthUrl2 = "/tokens/OAuth/2"
	Local $__sHttpBody = "grant_type=client_credentials&client_id=" & $_sClientID & "@" & $_sTenantID & "&client_secret=" & $_sClientSecret & "&resource=" & $_sResourceID & $_sTenantDomainName & "@" & $_sTenantID
	With $_oHTTP
		.open("POST", $__sAuthUrl1 & $_sTenantID & $__sAuthUrl2, False)
		.setRequestHeader("Host", "accounts.accesscontrol.windows.net")
		.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
		.setRequestHeader("Content-Length", StringLen($__sHttpBody))
		.send($__sHttpBody)
	EndWith

	If $_oHTTP.status == 200 Then
		$__aStringSplit1 = StringSplit($_oHTTP.responseText,",")
		If $__aStringSplit1[0] == 0 Then
			Return 1 ; Couldn't split response text or there is nothing in the responsetext. Token not found
		Else
			For $__i = 1 To $__aStringSplit1[0]
				$__aStringSplit2 = StringSplit($__aStringSplit1[$__i],":")
				For $__j = 1 To $__aStringSplit2[0]
					If $__aStringSplit2[$__j] == """access_token""" Then
						$__sToken = StringLeft($__aStringSplit2[$__j + 1],StringLen($__aStringSplit2[$__j + 1]) - 2)
						$__sToken = StringRight($__sToken,StringLen($__sToken) - 1)
						Return $__sToken
					EndIf
				Next
			Next
			Return 1 ; Couldn't find token.
		EndIf
	Else
		Return $_oHTTP.status ; HTTP error. Token not found
	EndIf
EndFunc

#comments-start
SPGetXDigestValue(ByRef $object, $string)
#comments-end
Func SPGetXDigestValue(ByRef $_oHTTP, $_sSiteUrl, $_sSecurityToken)
	Local $_aRxMatch
	Local $_sDigest
	If StringRight($_sSiteUrl,1) == "/" Then
		$_sSiteUrl = $_sSiteUrl & "_api/contextinfo"
	Else
		$_sSiteUrl = $_sSiteUrl & "/_api/contextinfo"
	EndIf

	With $_oHTTP
		.open("POST", $_sSiteUrl, False)
		.setRequestHeader("accept", "application/json;odata=verbose")
		.setRequestHeader("authorization", "Bearer " & $_sSecurityToken)
		.send()
	EndWith

	If $_oHTTP.status == 200 Then
		$_aRxMatch = StringRegExp($_oHTTP.responseText,"FormDigestValue"":""0x[a-fA-F0-9]+,",$STR_REGEXPARRAYMATCH)
		If UBound($_aRxMatch) > 0 Then
			$_sDigest = StringRight($_aRxMatch[0], StringLen($_aRxMatch[0]) - 18)
			$_sDigest = StringLeft($_sDigest, StringLen($_sDigest) - 1)
			Return $_sDigest
		EndIf
	EndIf

	Return 1 ; Couldn't find x digest value
EndFunc
