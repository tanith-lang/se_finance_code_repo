SELECT *
FROM (SELECT c.id          as credit_id,
             c.version     as version,
             SUM(c.amount) as amount,
             c.date_created,
             c.last_updated,
             c.type,
             c.billing_id,
             c.reason,
             c.status,
             c.currency,
             c.batch_id
      FROM
            credit c
      WHERE MONTH(date_created) = 9 AND YEAR(date_created) = 2018
      GROUP BY
            c.id) AS c
      LEFT JOIN credit_version cv ON
            cv.credit_id = c.credit_id AND cv.version = c.version
;