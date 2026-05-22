dbMemo "SQL" ="SELECT tblSession.SessionID, tblSession.StepID, tblSession.SessionDate, tblSessi"
    "on.SessionStartTime, tblSession.SessionEndTime, tblSession.ContextReset, tblSess"
    "ion.RevisionCycles, tblSession.Notes\015\012FROM tblSession\015\012WHERE IIf(Tem"
    "pVarsLong(\"lngStepID\")=0,0,tblSession.StepID) IN (0,TempVarsLong(\"lngStepID\""
    "))\015\012ORDER BY tblSession.SessionDate DESC , tblSession.SessionStartTime DES"
    "C;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
