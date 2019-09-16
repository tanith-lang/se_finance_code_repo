-- SE Credit Report --
SELECT
       c.id                 as credit_id,
       c.version            as credit_version,
       c.date_created,
       c.last_updated,
       c.type               as credit_type,
       c.billing_id,
       c.reason             as credit_reason,
       c.status             as credit_status,
       c.currency           as credit_currency,
       t.name               as territory,
       su.id                as user_id,
       au.id                as affiliate_user_id,
       b.id                 as booking_id,
       b.date_created       as booking_date_created,
       b.type               as booking_type,
       b.status             as booking_status,
       v.id                 as voucher_id,
       v.date_created       as voucher_date_created,
       v.type               as voucher_type,
       v.status             as voucher_status,
       Sum(c.amount)        as amount
FROM
       secretescapes.credit c
       -- Get booking information --
       LEFT JOIN secretescapes.booking_credit bc
              ON bc.credit_id = c.id
       LEFT JOIN secretescapes.booking b
              ON b.id = bc.booking_credits_used_id
       -- Get voucher information --
       LEFT JOIN secretescapes.voucher v
              ON v.credit_id = c.id
       -- Get user information --
       LEFT JOIN secretescapes.affiliate_user au
              ON au.id = b.affiliate_user_id
       LEFT JOIN secretescapes.shiro_user su
              ON su.id = b.user_id
       -- Get territory information --
       LEFT JOIN secretescapes.affiliate a
              ON a.id = su.affiliate_id OR a.id = au.affiliate_id
       LEFT JOIN secretescapes.territory t
              ON t.id = a.territory_id
WHERE
       Year(c.last_updated) = 2019
       AND Month(c.last_updated) = 8
GROUP BY c.id;
