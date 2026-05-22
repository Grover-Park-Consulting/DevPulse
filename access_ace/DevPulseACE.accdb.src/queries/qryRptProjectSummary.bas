dbMemo "SQL" ="SELECT StatusGroup, StatusGroupLabel, ProjectID, ProjectName, ProjectDescription"
    ", ProjectNotes, ProjectAreaName, ProjectStatus, ProjectStartDate, ProjectComplet"
    "ionDate, StepID, StepName, StepTypeName, StepStatus, StepStartDate, StepCompleti"
    "onDate, SessionCount, RevisionTotal, ResetCount\015\012FROM vw_RptProjectSummary"
    "\015\012ORDER BY StatusGroup, ProjectID, StepID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
