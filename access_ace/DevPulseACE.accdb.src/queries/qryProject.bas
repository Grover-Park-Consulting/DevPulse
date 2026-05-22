dbMemo "SQL" ="SELECT tblProject.ProjectID, tblProject.ProjectName, tblProject.ProjectDescripti"
    "on, tblProject.ProjectAreaID, tblProject.StatusID, tblProject.StartDate, tblProj"
    "ect.CompletionDate, tblProject.Notes\015\012FROM tblProject\015\012WHERE IIf(Tem"
    "pVarsLong(\"lngProjectID\")=0,0,tblProject.ProjectID) IN (0,TempVarsLong(\"lngPr"
    "ojectID\"))\015\012ORDER BY tblProject.ProjectID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
