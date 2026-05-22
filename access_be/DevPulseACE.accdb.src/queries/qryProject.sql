SELECT
  tblProject.ProjectID,
  tblProject.ProjectName,
  tblProject.ProjectDescription,
  tblProject.ProjectAreaID,
  tblProject.StatusID,
  tblProject.StartDate,
  tblProject.CompletionDate,
  tblProject.Notes
FROM
  tblProject
WHERE
  IIf(
    TempVarsLong("lngProjectID")= 0,
    0,
    tblProject.ProjectID
  ) IN (
    0,
    TempVarsLong("lngProjectID")
  )
ORDER BY
  tblProject.ProjectID;
