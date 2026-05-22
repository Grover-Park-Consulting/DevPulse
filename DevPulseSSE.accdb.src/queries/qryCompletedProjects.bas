dbMemo "SQL" ="SELECT tblProject.ProjectID, tblProject.ProjectName, tlkpStatus.StatusName, tblP"
    "roject.StartDate, tblProject.CompletionDate\015\012FROM tblProject LEFT JOIN tlk"
    "pStatus ON tblProject.StatusID = tlkpStatus.StatusID\015\012WHERE tblProject.Sta"
    "tusID IN (4,5)\015\012ORDER BY tblProject.CompletionDate DESC;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
