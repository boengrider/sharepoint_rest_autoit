#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <Constants.au3>
#include <String.au3>
#include <StringConstants.au3>




Local $oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
;Local $oHTTP = ObjCreate("MSXML2.ServerXMLHTTP.6.0")
MsgBox(0,"Bearer realm",SPGetTenantRealmID($oHTTP,"https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it"))

#comments-start
SPGetTenantRealmID(ByRef $ojbect, $string)
Function obtains Bearer realm aka Tenat ID
which in turn used in order to obtain the
security token
#comments-end
Func SPGetTenantRealmID(ByRef $_oHTTP, $_sSiteUrl)
	Local $__sResponseHeader
	Local $__aRxMatch
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
			Return StringRight($__aRxMatch[0],StringLen($__aRxMatch[0]) - 7)
		Else
			Return 0
		EndIf
	Else
		Return 0
	EndIf
EndFunc

#comments-start
SPGetSecurityToken(ByRef $object, $string, $string

