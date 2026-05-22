SELECT
  tblSession.SessionID,
  tblSession.StepID,
  tblSession.SessionDate,
  tblSession.SessionStartTime,
  tblSession.SessionEndTime,
  tblSession.ContextReset,
  tblSession.RevisionCycles,
  tblSession.Notes
FROM
  tblSession
WHERE
  IIf(
    TempVarsLong("lngStepID")= 0,
    0,
    tblSession.StepID
  ) IN (
    0,
    TempVarsLong("lngStepID")
  )
ORDER BY
  tblSession.SessionDate DESC,
  tblSession.SessionStartTime DESC;
