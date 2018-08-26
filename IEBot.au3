#include <IE.au3>
#include <Excel.au3>

#comments-start
	Excel file has two worksheets "Register" or "Booking" which are tied by ID.

#comments-end

$rIndex = "2" ; Start of ID numbers that verify if users need to be booked
$cIndex = "1" ; ID Column
Global Const $SHEETREG = "Register"
Global Const $SHEETBOO = "Booking"
$userEntered = False

Call("initBot")

Func initBot()
	Local $xlsDir = @MyDocumentsDir & "\AutoIt\Clients.xls"
	Global $oXls = _Excel_Open()
	Global $oWorkbook = _Excel_BookOpen($oXls,$xlsDir)
EndFunc

Func startBrowser()
	Global $oBrowser =_IECreate("http://newtours.demoaut.com/")
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel Open Failed", "Error while creating excel application." & @CRLF & "@error = " & @error)
		;1 - $oExcel is not an object or not an application object
		;2 - $vObject is not an object or an invalid A1 range. @error is set to the COM error code
		;3 - $sFileName is empty
		;4 - Error exporting the object. @extended is set to the COM error code returned by the ExportAsFixedFormat method
EndFunc

While Call("passengersExist",$cIndex,$rIndex)
	If Not $userEntered Then Call("startBrowser"); Calling new browser only when User is yet not registered or
	$userEntered = Call("registerUser",$rIndex,$cIndex);Excel Sheet handling _Excel_ColumnToLetter($columnIndex) & $rowIndex
	If $userEntered Then
		_IENavigate($oBrowser,"http://newtours.demoaut.com/") ;reset browser view.
	Else
	Call("loginUser",$rIndex)
	Call("bookFlight",$rIndex)
	EndIf
	$rIndex += 1
WEnd



Func bookFlight($rI)
	$id = _Excel_RangeRead($oWorkbook,$SHEETREG,_Excel_ColumnToLetter("1") & $rI)
	_Excel_RangeWrite($oWorkbook,$SHEETREG,"X",_Excel_ColumnToLetter("1") & $rI)
	$bRI = Call("getBookingRangeOfID",$id)
	$pCount = Call("bookingDataTransfer",$bRI)
	$buttonContinue = _IEGetObjByName($oBrowser,"findFlights")
	_IEAction($buttonContinue,"click")
	_IELoadWait($oBrowser)
	$buttonContinue = _IEGetObjByName($oBrowser,"reserveFlights")
	_IEAction($buttonContinue,"click")
	_IELoadWait($oBrowser)
	Call("passangerDataTransfer",$bRI,$pCount)
	$buttonBuy = _IEGetObjByName($oBrowser,"buyFlights")
	_IEAction($buttonBuy,"click")
	_Excel_BookSave($oWorkbook)
EndFunc

Func passangerDataTransfer($bRI,$pCount)
	For $i = 0 To $pCount-1
		$fName = _IEGetObjByName($oBrowser,"passFirst" & $i)
		$lName = _IEGetObjByName($oBrowser,"passLast" & $i)
		_IEPropertySet($fName,"innerText",Call("getBookingData","2",$bRI+$i+1))
		_IEPropertySet($lName,"innerText",Call("getBookingData","3",$bRI+$i+1))
	Next
	$creditnumber = _IEGetObjByName($oBrowser,"creditnumber")
	_IEPropertySet($creditnumber,"innerText",Call("getBookingData","2",$bRI+$pCount+1))
EndFunc

Func bookingDataTransfer($bRI)
	$tripType = "tripType"
	$fromPort = _IEGetObjByName($oBrowser,"fromPort")
	$passCount = _IEGetObjByName($oBrowser,"passCount")
	$fromMonth = _IEGetObjByName($oBrowser,"fromMonth")
	$fromDay = _IEGetObjByName($oBrowser,"fromDay")
	$toPort = _IEGetObjByName($oBrowser,"toPort")
	$toMonth = _IEGetObjByName($oBrowser,"toMonth")
	$toDay = _IEGetObjByName($oBrowser,"toDay")
	$servClass = "servClass"
	$airline = _IEGetObjByName($oBrowser,"airline")
	$form = _IEFormGetObjByName($oBrowser,"findflight")
	For $i = 1 To 10
		Switch $i
			Case "1"
				$tripTypeData = StringLower(StringStripWS(Call("getBookingData","1"+$i,$bRI),$STR_STRIPALL))
				_IEFormElementRadioSelect($form,$tripTypeData,$tripType,1,"byValue")
			Case "2"
				$pass = Call("getBookingData","1"+$i,$bRI)
				_IEFormElementOptionSelect($passCount,Call("getBookingData","1"+$i,$bRI),1,"byText")
			Case "3"
				_IEFormElementOptionSelect($fromPort,Call("getBookingData","1"+$i,$bRI),1,"byText")
			Case "4"
				_IEFormElementOptionSelect($fromMonth,Call("getBookingData","1"+$i,$bRI),1,"byText")
			Case "5"
				_IEFormElementOptionSelect($fromDay,Call("getBookingData","1"+$i,$bRI),1,"byText")
			Case "6"
				_IEFormElementOptionSelect($toPort,Call("getBookingData","1"+$i,$bRI),1,"byText")
			Case "7"
				_IEFormElementOptionSelect($toMonth,Call("getBookingData","1"+$i,$bRI),1,"byText")
			Case "8"
				_IEFormElementOptionSelect($toDay,Call("getBookingData","1"+$i,$bRI),1,"byText")
			Case "9"
				$servClassData = StringLower(StringStripWS(Call("getBookingData","1"+$i,$bRI),$STR_STRIPALL))
				If $servClassData == "economyclass" Then
					_IEFormElementRadioSelect($form,"Coach",$servClass,1,"byValue")
				ElseIf $servClassData == "businessclass" Then
					_IEFormElementRadioSelect($form,"Business",$servClass,1,"byValue")
				ElseIf $servClassData == "firstclass" Then
					_IEFormElementRadioSelect($form,"First",$servClass,1,"byValue")
				EndIf
			Case "10"
				_IEFormElementOptionSelect($airline,Call("getBookingData","1"+$i,$bRI),1,"byText")
		EndSwitch
	Next
	Return($pass)
EndFunc

Func getBookingRangeOfID($ID)
	$bCI = "1"
	$bRI = "2"
	$findID = _Excel_RangeRead($oWorkbook,$SHEETBOO,_Excel_ColumnToLetter($bCI)&$bRI)
	While $findID <> $ID
		$bRI += 1
		$findID = _Excel_RangeRead($oWorkbook,$SHEETBOO,_Excel_ColumnToLetter($bCI)&$bRI)
	WEnd
	Return($bRI)
EndFunc

Func passengersExist($cI,$rI)
	$cheking = True
	$pE = False
	$cI = "1"
	$rI = "2"
	While $cheking
		$cellData = _Excel_RangeRead($oWorkbook,$SHEETREG,_Excel_ColumnToLetter($cI) & $rI)
		If (IsNumber($cellData) And ($cellData <> "0")) Then $pE = True
		If $cellData = "" Then $cheking = False
		$rI += 1
	WEnd
	If $pE Then
		Return(True)
	Else
		Return(False)
	EndIf
EndFunc

Func loginUser($rIndex)
	$BUTTONSIG = " sign-in "
	_IELinkClickByText($oBrowser,$BUTTONSIG)
	_IELoadWait($oBrowser)
	$uName = _IEGetObjByName($oBrowser,"userName")
	$pass = _IEGetObjByName($oBrowser,"password")
	_IEPropertySet($uName,"innerText",Call("getUserData","12",$rIndex))
	_IEPropertySet($pass,"innerText",Call("getUserData","13",$rIndex))
	$buttonLogin = _IEGetObjByName($oBrowser,"login")
	_IEAction($buttonLogin,"click")
	_IELoadWait($oBrowser)
EndFunc

Func registerUser($rowIndex,$columnIndex)
	$userEntered = False
	Const $BUTTONREG = "REGISTER"
	$vRange = _Excel_ColumnToLetter($columnIndex) & $rowIndex
	_IELinkClickByText($oBrowser,$BUTTONREG)
	_IELoadWait($oBrowser)
	If IsNumber(_Excel_RangeRead($oWorkbook,$SHEETREG,$vRange)) Then
		Call("UserDataTransfer",$rowIndex,$columnIndex)
	Else
		$userEntered = True
	EndIf
	$buttonRegister = _IEGetObjByName($oBrowser,"register")
	If $userEntered Then
		Return(True)
	Else
		_IEAction($buttonRegister,"click")
		_IELoadWait($oBrowser)
		Return(False)
	EndIf
EndFunc

Func UserDataTransfer($rIndex,$cIndex)
	$fName = _IEGetObjByName($oBrowser,"firstName")
	$lName = _IEGetObjByName($oBrowser,"lastName")
	$phone = _IEGetObjByName($oBrowser,"phone")
	$uName = _IEGetObjByName($oBrowser,"userName")
	$addr1 = _IEGetObjByName($oBrowser,"address1")
	$addr2 = _IEGetObjByName($oBrowser,"address2")
	$city = _IEGetObjByName($oBrowser,"city")
	$state = _IEGetObjByName($oBrowser,"state")
	$pCode = _IEGetObjByName($oBrowser,"postalCode")
	$country = _IEGetObjByName($oBrowser,"country")
	$email = _IEGetObjByName($oBrowser,"email")
	$password = _IEGetObjByName($oBrowser,"password")
	$confirmPassword = _IEGetObjByName($oBrowser,"confirmPassword")
	For $i = 1 To 12
		Switch $i
			Case "1"
				_IEPropertySet($fName,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "2"
				_IEPropertySet($lName,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "3"
				_IEPropertySet($phone,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "4"
				_IEPropertySet($uName,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "5"
				_IEPropertySet($addr1,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "6"
				_IEPropertySet($addr2,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "7"
				_IEPropertySet($city,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "8"
				_IEPropertySet($state,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "9"
				_IEPropertySet($pCode,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "10"
				_IEFormElementOptionSelect($country,Call("getUserData",$cIndex+$i,$rIndex),1,"byText")
			Case "11"
				_IEPropertySet($email,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
			Case "12"
				_IEPropertySet($password,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
				_IEPropertySet($confirmPassword,"innerText",Call("getUserData",$cIndex+$i,$rIndex))
		EndSwitch
	Next
EndFunc

Func getUserData($cI,$rI)
	Return(_Excel_RangeRead($oWorkbook,$SHEETREG,_Excel_ColumnToLetter($cI) & $rI))
EndFunc

Func getBookingData($cI,$rI)
	Return(_Excel_RangeRead($oWorkbook,$SHEETBOO,_Excel_ColumnToLetter($cI) & $rI))
EndFunc