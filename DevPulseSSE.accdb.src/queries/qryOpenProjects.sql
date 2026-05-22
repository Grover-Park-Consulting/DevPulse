SELECT
  tblProject.ProjectID,
  tblProject.ProjectName,
  tlkpStatus.StatusName,
  tblProject.StartDate
FROM
  tblProject
  LEFT JOIN tlkpStatus ON tblProject.StatusID = tlkpStatus.StatusID
WHERE
  (
    (
      (tblProject.StatusID) In (1, 2, 3)
    )
  )
ORDER BY
  tblProject.StartDate;
