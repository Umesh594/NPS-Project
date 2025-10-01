SELECT
    COUNT(*) AS Missing_Records
FROM
    NPS_dataset
WHERE
    "Ticket No" IS NULL
    OR "Created Date/Time" IS NULL
    OR "Assigned To" IS NULL;
SELECT
    *
FROM
    NPS_dataset
WHERE
    Status = 'Complete' AND "Sub Status" = 'Pending';
SELECT
    "Issue 2 - NPS",
    AVG(DATEDIFF(day, "Created Date/Time", "Resolved Date/Time")) AS Avg_Resolution_Days
FROM
    NPS_dataset
GROUP BY
    "Issue 2 - NPS"
ORDER BY
    Avg_Resolution_Days DESC;
