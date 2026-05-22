SELECT
  p.ProjectName,
  'Completed Steps' AS StepCategory,
  Count(
    IIf(s.StatusID = 4, s.StepID, Null)
  ) AS StepCount
FROM
  tblProject AS p
  INNER JOIN tblStep AS s ON p.ProjectID = s.ProjectID
WHERE
  p.StatusID In (1, 2, 3)
GROUP BY
  p.ProjectName
UNION ALL
SELECT
  p.ProjectName,
  'Open Steps' AS StepCategory,
  Count(
    IIf(s.StatusID <> 4, s.StepID, Null)
  ) AS StepCount
FROM
  tblProject AS p
  INNER JOIN tblStep AS s ON p.ProjectID = s.ProjectID
WHERE
  p.StatusID In (1, 2, 3)
GROUP BY
  p.ProjectName;
