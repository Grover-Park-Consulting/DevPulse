dbMemo "SQL" ="SELECT tblSession.StepID, tblSessionTool.ToolTypeID, SUM(tblSessionTool.UseCount"
    ") AS TotalUseCount\015\012FROM tblSessionTool INNER JOIN tblSession ON tblSessio"
    "n.SessionID = tblSessionTool.SessionID\015\012GROUP BY tblSession.StepID, tblSes"
    "sionTool.ToolTypeID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
