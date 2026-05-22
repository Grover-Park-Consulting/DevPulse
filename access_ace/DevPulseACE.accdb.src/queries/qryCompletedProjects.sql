SELECT
  tblProject.ProjectID,
  tblProject.ProjectName,
  tlkpStatus.StatusName,
  tblProject.StartDate,
  tblProject.CompletionDate
FROM
  tblProject
  LEFT JOIN tlkpStatus ON tblProject.StatusID = tlkpStatus.StatusID
WHERE
  tblProject.StatusID IN (4, 5)
ORDER BY
  tblProject.CompletionDate DESC;
