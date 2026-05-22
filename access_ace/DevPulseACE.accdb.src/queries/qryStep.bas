dbMemo "SQL" ="SELECT tblStep.StepID, tblStep.ProjectID, tblStep.StepName, tblStep.StepDescript"
    "ion, tblStep.StepTypeID, tblStep.StatusID, tblStep.PredecessorStepID, tblStep.St"
    "artDate, tblStep.CompletionDate, tblStep.Notes\015\012FROM tblStep\015\012WHERE "
    "IIf(TempVarsLong(\"lngStepID\")=0,0,tblStep.StepID) IN (0,TempVarsLong(\"lngStep"
    "ID\")) AND IIf(TempVarsLong(\"lngProjectID\")=0,0,tblStep.ProjectID) IN (0,TempV"
    "arsLong(\"lngProjectID\"))\015\012ORDER BY tblStep.StepID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
