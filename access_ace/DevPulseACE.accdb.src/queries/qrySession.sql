SELECT
  tblSession.SessionID,
  tblSession.StepID,
  tblSession.SessionDate,
  tblSession.SessionStartTime,
  tblSession.SessionEndTime,
  tblSession.ContextReset,
  tblSession.RevisionCycles,
  tblSession.Notes,
  tblStep.ProjectID
FROM
  tblStep
  INNER JOIN tblSession ON tblStep.StepID = tblSession.StepID
WHERE
  (
    (
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
    AND (
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
  )
ORDER BY
  tblSession.SessionDate DESC,
  tblSession.SessionStartTime DESC;
