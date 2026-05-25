dbMemo "SQL" ="SELECT tblSession.SessionID, tblSession.StepID, tblSession.SessionDate, tblSessi"
    "on.SessionStartTime, tblSession.SessionEndTime, tblSession.ContextReset, tblSess"
    "ion.RevisionCycles, tblSession.Notes, tblStep.ProjectID\015\012FROM tblStep INNE"
    "R JOIN tblSession ON tblStep.StepID = tblSession.StepID\015\012WHERE (((IIf(Temp"
    "VarsLong(\"lngProjectID\")=0,0,[tblStep].[ProjectID])) In (0,TempVarsLong(\"lngP"
    "rojectID\"))) AND ((IIf(TempVarsLong(\"lngStepID\")=0,0,[tblSession].[StepID])) "
    "In (0,TempVarsLong(\"lngStepID\"))))\015\012ORDER BY tblSession.SessionDate DESC"
    " , tblSession.SessionStartTime DESC;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "NoFormat" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblSession.RevisionCycles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.StepID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.SessionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.SessionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.SessionStartTime"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.SessionEndTime"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.ContextReset"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStep.ProjectID"
        dbLong "AggregateType" ="-1"
    End
End
