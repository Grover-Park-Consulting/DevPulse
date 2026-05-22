SELECT
  tblSession.StepID,
  tblSessionTool.ToolTypeID,
  SUM(tblSessionTool.UseCount) AS TotalUseCount
FROM
  tblSessionTool
  INNER JOIN tblSession ON tblSession.SessionID = tblSessionTool.SessionID
GROUP BY
  tblSession.StepID,
  tblSessionTool.ToolTypeID;
