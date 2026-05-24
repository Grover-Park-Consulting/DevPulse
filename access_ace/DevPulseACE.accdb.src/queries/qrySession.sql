SELECT
  tblSession.SessionID,
  tblSession.StepID,
  tblSession.SessionDate,
  tblSession.SessionStartTime,
  tblSession.SessionEndTime,
  tblSession.ContextReset,
  tblSession.RevisionCycles,
  tblSession.Notes,
  tblStep.ProjectID,
  IIf(
    TempVarsLong("lngProjectID")= 0,
    0,
    [tblStep].[ProjectID]
  ) AS Expr1
FROM
  tblSession
  INNER JOIN tblStep ON tblSession.StepID = tblStep.StepID
WHERE
  (
    (
      (
        IIf(
          TempVarsLong("lngStepID")= 0,
          0,
          [tblSession].[StepID]
        )
      ) In (
        0,
        TempVarsLong("lngStepID")
      )
    )
    AND (
      (
        IIf(
          TempVarsLong("lngProjectID")= 0,
          0,
          [tblStep].[ProjectID]
        )
      ) In (
        0,
        TempVarsLong("lngProjectID")
      )
    )
  )
ORDER BY
  tblSession.SessionDate DESC,
  tblSession.SessionStartTime DESC;
