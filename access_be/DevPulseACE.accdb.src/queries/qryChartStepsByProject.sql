SELECT
  ProjectName,
  StepCategory,
  StepCount
FROM
  vw_ChartStepsLong
ORDER BY
  ProjectName,
  StepCategory;
