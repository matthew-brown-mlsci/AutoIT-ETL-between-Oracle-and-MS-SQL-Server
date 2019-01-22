#cs ----------------------------------------------------------------------------

Copy the old information from HLAB_Lookup db to the new Oracle database for the new MLab environment

- Used during 2017 MLab (Now Allscripts Lab) upgrade to v 16.03 - This is an AutoIT-esque
  method to connect to 2 different databases (in this case, Oracle 11G & MS SQL Server) and
  perform ETL operations.  We copy data from the MS SQL Server to an Oracle 11G
  database.

  The Oracle side are defined as ODBC system DSN's \w Oracle's InstantClient driver.  The SQL
  Server side uses the SQL Server driver + pass-thru auth.

#ce ----------------------------------------------------------------------------
#include <FTPEx.au3>
#include <Array.au3>
#include <StringConstants.au3>
#include <Constants.au3>
#include <InetConstants.au3>
#include <MsgBoxConstants.au3>
#include <Date.au3>

Global $sqlResults[1][1]
$sqlResults[0][0] = "EmptyArrayReturn0"
Global $sqlResults2[1][1]
$sqlResults2[0][0] = "EmptyArrayReturn0"

Func _ErrFunc($oError)
    ; Do anything here.
    ConsoleWrite(@ScriptName & " (" & $oError.scriptline & ") : ==> COM Error intercepted !" & @CRLF & _
            @TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
            @TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
            @TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
            @TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
            @TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
            @TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
            @TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
            @TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
            @TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc   ;==>_ErrFunc

; This function takes SQL SELECT statements, results from queries go into the global sqlResults[][]
Func sqlSelect($sqlObj, $sql, $error_desc)

	; Blank out the global holder for the returned rows
	ReDim $sqlResults[1][1]
	$sqlResults[0][0] = "EmptyArrayReturn0"

	; Create the RecordSet object
	$rs = ObjCreate( "ADODB.RecordSet" )
	$rs.open($sql, $sqlObj)

	With $rs
		While Not .EOF ; repeat until End-Of-File (EOF) is reached
			; Write the content of all fields to the console separated by | by processing the fields collection
			;ConsoleWrite("Process the fields collection:     ")
			ReDim $sqlResults[UBound($sqlResults, 1) + 1][$rs.Fields.Count]
			$i = 0
			For $oField In .Fields
				;ConsoleWrite($oField.Value & "|")
				$sqlResults[UBound($sqlResults, 1) - 2][$i] = $oField.Value
				$i = $i + 1
			Next
			;ConsoleWrite(@CR)
			; Write a second line by accessing all fields of the collection by item number
			.MoveNext ; Move To the Next record
		WEnd
	EndWith

	; Clean up the last line in the array we added but don't need (unless we didn't add any!)
	If UBound($sqlResults, 1) > 1 Then
		ReDim $sqlResults[UBound($sqlResults, 1) - 1][$rs.Fields.Count]
	EndIf

	$rs.close()

EndFunc

; This function takes insert or update statements (Execute on ADODB object, catch errors, that's it)
Func sqlInsertUpdate($sqlObj, $sql, $error_desc)

	; Blank out the global holder for the returned rows - we won't return any rows with this function, but making sure
	; there is *no* data to return from our wrapper functions helps with debugging mishaps.
	ReDim $sqlResults[1][1]
	$sqlResults[0][0] = "EmptyArrayReturn0"

	; Execute the statement
	$sqlObj.Execute($sql, $sqlObj)

	; Write out errors to log file if applicable:
	If @error Then
		ConsoleWrite(_NowCalc() & "******************************************")
		ConsoleWrite($error_desc)
		Return
	EndIf

EndFunc

; This pulls data from the array $sqlResults, one row at a time (starting with the first row returned), until it's empty.  If empty, return 0.
Func sqlFetch()
	Local $result_array[UBound($sqlResults, 2)]

	; If there is only one row in $sqlResults, we process here to see if it's "empty" or has some sort of data in it
	If UBound($sqlResults, 1) = 1 Then
		If $sqlResults[0][0] = "EmptyArrayReturn0" Then
			;ConsoleWrite("Return $result_array[0] = EmptyArrayReturn0" & @CRLF)
			Return 0
		Else
			For $i = 1 To UBound($sqlResults, 2)
				$result_array[$i - 1] = StringStripWS($sqlResults[0][$i - 1], 2)
			Next
			ReDim $sqlResults[1][1]
			$sqlResults[0][0] = "EmptyArrayReturn0"
			Return $result_array
		EndIf
	EndIf

	; If there is more than one row in $sqlResults we copy the top row into $result_array, delete the top row, and return the 1D array
	For $i = 1 To UBound($sqlResults, 2)
		$result_array[$i - 1] = StringStripWS($sqlResults[0][$i - 1], 2)
	Next
	_ArrayDelete($sqlResults, 0)
	Return $result_array

EndFunc


Func Main()

   Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")

   ; get all info from the old db
   $sqlObj_olddb = ObjCreate("ADODB.Connection")
   $sqlError = ObjCreate("ADODB.Error")
   $sqlObj_olddb.Open(("DRIVER={SQL Server};SERVER=MKLABWIDB01,1433;DATABASE=HLAB_Lookup;UID=;PWD=;"))

   $sql = "SELECT Workstation, ProdCitrix, TestCitrix, ProdThick, TestThick, Location, AddlInfo, Facility FROM [HLAB_Lookup].[dbo].[Clients]"
   sqlSelect($sqlObj_olddb, $sql, $sql)
   Local $old_db_values[1][8]
   $count = 0
   While (1)
	  $temp_array = sqlFetch()
	  If @error = -1 Then ExitLoop

	  If UBound($temp_array) > 1 Then
		 ReDim $old_db_values[UBound($old_db_values) + 1][8]
		 $old_db_values[$count][0] = $temp_array[0]
		 $old_db_values[$count][1] = $temp_array[1]
		 $old_db_values[$count][2] = $temp_array[2]
		 $old_db_values[$count][3] = $temp_array[3]
		 $old_db_values[$count][4] = $temp_array[4]
		 $old_db_values[$count][5] = $temp_array[5]
		 $old_db_values[$count][6] = $temp_array[6]
		 $old_db_values[$count][7] = $temp_array[7]
		 $count = $count + 1
		 ConsoleWrite("Reading row: " & String($count) & @CRLF)
	  Else
		 ExitLoop
	  EndIf
   WEnd

   $sqlObj_olddb.Close
   ;_ArrayDisplay($old_db_values)


   ; Do upgrade database
   $db_mlab_upgrade = ObjCreate("ADODB.Connection")
   $sqlError = ObjCreate("ADODB.Error")
   $db_mlab_upgrade.Open(("DRIVER={SQL Server};SERVER=comwfdbtest1,1433;DATABASE=mlab_upgrade;UID=;PWD=;"))

   ConsoleWrite("UBound($old_db_values,1) " & UBound($old_db_values,1) & @CRLF)
   ConsoleWrite("UBound($old_db_values,2) " & UBound($old_db_values,2) & @CRLF)

   $sql = "SELECT client_workstation FROM [mlab_upgrade].[dbo].[client] WHERE client_workstation = 'eat'"

   sqlSelect($db_mlab_upgrade, $sql, "na")
   $does_client_exist = sqlFetch()
   ConsoleWrite($sql & @CRLF)

   Local $row[1]
   For $row_index = 0 To UBound($old_db_values,1) - 1
	  ConsoleWrite("Processing row " & $row_index & " of " & String(UBound($old_db_values,1) - 1) & @CRLF)
	  Dim $row[UBound($old_db_values,2)]
	  For $column_index = 0 To UBound($old_db_values,2) - 1
		 $row[$column_index] = ($old_db_values[$row_index][$column_index])
	  Next
	  $sql = "SELECT client_workstation FROM [mlab_upgrade].[dbo].[client] WHERE client_workstation = '" & $row[0] & "'"
	  ;ConsoleWrite($sql & @CRLF)
	  sqlSelect($db_mlab_upgrade, $sql, "na")
	  $does_client_exist = sqlFetch()
	  If UBound($does_client_exist) > 0 Then
		 ConsoleWrite("Updating client_workstation: " & $row[0] & @CRLF)
		 $sql = "UPDATE client SET client_node_citrix = '" & $row[2] & "', "
		 $sql = $sql & "client_node_thick = '" & $row[4] & "', "
		 $sql = $sql & "location = '" & $row[5] & "', "
		 $sql = $sql & "facility = '" & $row[7] & "', "
		 $sql = $sql & "notes = '" & $row[6] & "', "
		 $sql = $sql & "edited = '" & _NowCalc() & "', "
		 $sql = $sql & "editedby = 'update script' "
		 $sql = $sql & "WHERE client_workstation = '" & $row[0] & "'"
		 ;ConsoleWrite($sql & @CRLF)
		 sqlInsertUpdate($db_mlab_upgrade, $sql, $sql)
	  Else
		 ConsoleWrite("Inserting client_workstation: " & $row[0] & @CRLF)
		 $sql = "INSERT INTO client (client_workstation, client_node_citrix, client_node_thick, facility, location, notes, created, createdby) VALUES "
		 $sql = $sql & "('" & StringReplace($row[0], "'", "") & "', '" & StringReplace($row[2], "'", "") & "', '" & StringReplace($row[4], "'", "") & "', '" & StringReplace($row[7], "'", "") & "', '" & StringReplace($row[5], "'", "") & "', '" & StringReplace($row[6], "'", "") & "', SYSDATETIME(), "
		 $sql = $sql & "'update script')"
		 sqlInsertUpdate($db_mlab_upgrade, $sql, $sql)
	  EndIf
   Next

   ;$db_mlab_prod.Close
   ;$db_mlab_test.Close
   $db_mlab_upgrade.Close
EndFunc

Main()
