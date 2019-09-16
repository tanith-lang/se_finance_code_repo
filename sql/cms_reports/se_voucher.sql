-- SE CMS Voucher Report
SELECT
       v.id                           AS voucher_id,
       v.version                      AS version,
       v.code                         AS voucher_code,
       v.unique_transaction_reference AS order_code,
       v.credit_id                    AS credit_id,
       v.giftee_id                    AS giftee_id,
       v.gifter_id                    AS gifter_id,
       v.payment_id                   AS payment_id,
       v.type                         AS voucher_type,
       v.status                       AS voucher_status,
       p.type                         AS payment_method,
       p.status                       AS payment_status,
       c.status                       AS credit_status,
       t.currency                     AS currency,
       t.name                         AS territory,
       v.date_created                 AS date_created,
       tc.expires_on                  AS expiry_date,
       v.last_updated                 AS last_updated,
       c.date_created                 AS redemption_date,
       s.date_created                 AS sender_join_date,
       su.date_created                AS recipient_join_date,
       Sum(p.amount)                  AS payment_amount
FROM
       secretescapes.voucher v
       -- Get payment information --
       LEFT JOIN secretescapes.payment p
              ON p.id = v.payment_id
       LEFT JOIN secretescapes.credit c
              ON c.id = v.credit_id
       LEFT JOIN secretescapes.time_limited_credit tc
              ON tc.id = c.id
       -- Get user information --
       LEFT JOIN secretescapes.shiro_user s
              ON s.id = v.gifter_id
       LEFT JOIN secretescapes.shiro_user su
              ON su.id = v.giftee_id
       LEFT JOIN secretescapes.affiliate a
              ON s.affiliate_id = a.id
           -- Get territory information --
       LEFT JOIN secretescapes.territory t
              ON a.territory_id = t.id
       -- Report filters --
WHERE  (v.status != 'ABANDONED' AND v.status != 'NEW')
       AND Year(v.date_created) = 2015
GROUP  BY voucher_id
ORDER  BY voucher_id;
