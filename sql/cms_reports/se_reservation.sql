-- SE Reservations Report --
SELECT
       r.id                             AS reservation_id,
       r.unique_transaction_reference   AS order_code,
       r.user_id                        AS user_id,
       r.affiliate_user_id              AS affiliate_user_id,
       r.payment_id                     AS payment_id,
       r.date_created                   AS date_created,
       r.last_updated                   AS last_updated,
       IF(r.type = 'BOOKING', 'RESERVATION', r.type)
                                        AS booking_type,
       r.status                         AS booking_status,
       p.status                         AS payment_status,
       p.type                           AS payment_method,
       t.name                           AS territory,
       r.currency                       AS currency,
       r.booking_fee                    AS booking_fee,
       p.surcharge                      AS payment_surcharge,
       Sum(c.amount)                    AS credit_amount,
       p.amount                         AS payment_amount
FROM   secretescapes.reservation r
       -- Get payment information --
       LEFT JOIN secretescapes.payment p
              ON p.id = r.payment_id
       LEFT JOIN secretescapes.reservation_credit rc
              ON rc.reservation_credits_used_id = r.id
       LEFT JOIN secretescapes.credit c
              ON c.id = rc.credit_id
       -- Get territory information --
       LEFT JOIN secretescapes.affiliate_user au
              ON au.id = r.affiliate_user_id
       LEFT JOIN secretescapes.shiro_user su
              ON su.id = r.user_id
       LEFT JOIN secretescapes.affiliate a
              ON a.id = su.affiliate_id OR a.id = au.affiliate_id
       LEFT JOIN secretescapes.territory t
              ON t.id = a.territory_id
       -- Report filters --
WHERE  r.status != 'ABANDONED'
       AND Month(r.date_created) = 8
       AND Year(r.date_created) = 2019
GROUP  BY reservation_id
ORDER  BY reservation_id;
