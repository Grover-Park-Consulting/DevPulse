SELECT
  IIf(
    p.StatusID In (1, 2, 3),
    0,
    1
  ) AS StatusGroup,
  IIf(
    p.StatusID In (1, 2, 3),
    'Open Projects',
    'Closed Projects'
  ) AS StatusGroupLabel,
  p.ProjectID,
  p.ProjectName,
  p.ProjectDescription,
  p.Notes AS ProjectNotes,
  pa.ProjectAreaName,
  ps.StatusName AS ProjectStatus,
  p.StartDate AS ProjectStartDate,
  p.CompletionDate AS ProjectCompletionDate,
  s.StepID,
  s.StepName,
  st.StepTypeName,
  ss.StatusName AS StepStatus,
  s.StartDate AS StepStartDate,
  s.CompletionDate AS StepCompletionDate,
  Count(ses.SessionID) AS SessionCount,
  Nz(
    Sum(ses.RevisionCycles),
    0
  ) AS RevisionTotal,
  Nz(
    Sum(
      IIf(ses.ContextReset, 1, 0)
    ),
    0
  ) AS ResetCount
FROM
  (
    (
      (
        (
          (
            tblProject AS p
            LEFT JOIN tlkpProjectArea AS pa ON p.ProjectAreaID = pa.ProjectAreaID
          )
          LEFT JOIN tlkpStatus AS ps ON p.StatusID = ps.StatusID
        )
        LEFT JOIN tblStep AS s ON p.ProjectID = s.ProjectID
      )
      LEFT JOIN tlkpStepType AS st ON s.StepTypeID = st.StepTypeID
    )
    LEFT JOIN tlkpStatus AS ss ON s.StatusID = ss.StatusID
  )
  LEFT JOIN tblSession AS ses ON s.StepID = ses.StepID
GROUP BY
  IIf(
    p.StatusID In (1, 2, 3),
    0,
    1
  ),
  IIf(
    p.StatusID In (1, 2, 3),
    'Open Projects',
    'Closed Projects'
  ),
  p.ProjectID,
  p.ProjectName,
  p.ProjectDescription,
  p.Notes,
  pa.ProjectAreaName,
  ps.StatusName,
  p.StartDate,
  p.CompletionDate,
  s.StepID,
  s.StepName,
  st.StepTypeName,
  ss.StatusName,
  s.StartDate,
  s.CompletionDate;
