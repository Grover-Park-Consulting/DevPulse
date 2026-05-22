dbMemo "SQL" ="SELECT p.ProjectName, 'Completed Steps' AS StepCategory, Count(IIf(s.StatusID=4,"
    "s.StepID,Null)) AS StepCount\012FROM tblProject AS p INNER JOIN tblStep AS s ON "
    "p.ProjectID = s.ProjectID\012WHERE p.StatusID In (1,2,3)\012GROUP BY p.ProjectNa"
    "me\012UNION ALL SELECT p.ProjectName, 'Open Steps' AS StepCategory, Count(IIf(s."
    "StatusID<>4,s.StepID,Null)) AS StepCount\012FROM tblProject AS p INNER JOIN tblS"
    "tep AS s ON p.ProjectID = s.ProjectID\012WHERE p.StatusID In (1,2,3)\012GROUP BY"
    " p.ProjectName;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
