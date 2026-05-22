SELECT
  tblStep.StepID,
  tblStep.ProjectID,
  tblStep.StepName,
  tblStep.StepDescription,
  tblStep.StepTypeID,
  tblStep.StatusID,
  tblStep.PredecessorStepID,
  tblStep.StartDate,
  tblStep.CompletionDate,
  tblStep.Notes
FROM
  tblStep
WHERE
  IIf(
    TempVarsLong("lngStepID")= 0,
    0,
    tblStep.StepID
  ) IN (
    0,
    TempVarsLong("lngStepID")
  )
  AND IIf(
    TempVarsLong("lngProjectID")= 0,
    0,
    tblStep.ProjectID
  ) IN (
    0,
    TempVarsLong("lngProjectID")
  )
ORDER BY
  tblStep.StepID;
