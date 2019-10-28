# SE CMS User Bookings V1
SELECT u.id                AS se_user_id,
       Lower(u.username)   AS user_name,
       u.date_created      AS user_join_date,
       b.date_created      AS booking_date,
       b.status            AS booking_status
FROM   shiro_user u
       -- Get Purchase Information --
       LEFT JOIN booking b
              ON b.user_id = u.id
       LEFT JOIN reservation r
              ON u.id = r.user_id
       LEFT JOIN voucher v
              ON u.id = v.gifter_id
WHERE b.status = 'COMPLETE'
   OR b.status = 'REFUNDED'
;



# SE User Purchase History
SELECT DISTINCT
       u.id                AS se_user_id,
       Lower(u.username)   AS user_name,
       u.date_created      AS user_join_date
FROM   shiro_user u
       -- Get Purchase Information --
       LEFT JOIN booking b
              ON b.user_id = u.id
       LEFT JOIN reservation r
              ON u.id = r.user_id
       LEFT JOIN voucher v
              ON u.id = v.gifter_id
WHERE b.status = 'COMPLETE'
   OR b.status = 'REFUNDED'
   OR r.status = 'COMPLETE'
   OR r.status = 'REFUNDED'
   OR v.status = 'REDEEMED'
   OR v.status = 'READY_TO_REDEEM'
   OR v.status = 'REFUNDED';


# SE Affiliate Users
SELECT
       au.id      AS affiliate_user_id,
       au.date_created AS user_join_date
FROM
       secretescapes.affiliate_user au
WHERE Year(date_created) = 2019;



# SE Shiro Users
SELECT
       su.id           AS user_id,
       su.date_created AS user_join_date
FROM
       secretescapes.shiro_user su
WHERE Year(date_created) = 2019;



# SE CMS Affiliate User Bookings V1
SELECT au.id                AS affiliate_user_id,
       Lower(au.email)      AS user_name,
       au.date_created      AS user_join_date,
       Max(b.date_created)  AS booking_date,
       Max(b.status)        AS booking_status
FROM   affiliate_user au
       -- Get booking information --
       LEFT JOIN booking b
              ON b.affiliate_user_id = au.id
WHERE b.status = 'COMPLETE'
   OR b.status = 'REFUNDED'
GROUP BY au.id;



# SE Credit Report
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
       YEAR(c.last_updated) = 2015
GROUP BY c.id;



# SE CMS Reservations Report V2
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
#        AND Month(r.date_created) = 9
       AND Year(r.date_created) = 2016
GROUP  BY reservation_id
ORDER  BY reservation_id;



# SE CMS Booking Summary V3
  SELECT
       b.id                               AS booking_id,
       b.unique_transaction_reference     AS order_code,
       b.user_id                          AS user_id,
       b.affiliate_user_id                AS affiliate_user_id,
       b.payment_id                       AS payment_id,
       b.date_created                     AS date_created,
       b.completion_date                  AS completion_date,
       b.last_updated                     AS last_updated,
       b.type                             AS booking_type,
       b.status                           AS booking_status,
       b.hold_id                          AS hold_id,
       b.payment_amount_type              AS payment_type,
       p.status                           AS payment_status,
       p.type                             AS payment_method,
       t.name                             AS territory,
       b.currency                         AS currency,
       b.booking_fee                      AS booking_fee,
       b.atol_fee                         AS atol_fee,
       p.surcharge                        AS payment_surcharge,
       Sum(c.amount)                      AS credit_amount,
       p.amount                           AS payment_amount
FROM   secretescapes.booking b
       -- Get payment information --
       LEFT JOIN secretescapes.payment p
              ON p.id = b.payment_id
       LEFT JOIN secretescapes.booking_credit bc
              ON bc.booking_credits_used_id = b.id
       LEFT JOIN secretescapes.credit c
              ON c.id = bc.credit_id
       -- Get territory information --
       LEFT JOIN secretescapes.affiliate_user au
              ON au.id = b.affiliate_user_id
       LEFT JOIN secretescapes.shiro_user su
              ON su.id = b.user_id
       LEFT JOIN secretescapes.affiliate a
              ON a.id = su.affiliate_id OR a.id = au.affiliate_id
       LEFT JOIN secretescapes.territory t
              ON t.id = a.territory_id
       -- Report filters --
WHERE  b.status != 'ABANDONED'
#        AND Month(b.date_created) = 9
       AND Year(b.date_created) = 2014
GROUP  BY booking_id
ORDER  BY booking_id;



# SE CMS Ancillary Product Sales V1
SELECT
       ap.id                      AS ancillary_id,
       ap.version                 AS version,
       b.date_created             AS date_created,
       b.last_updated             AS last_updated,
       ap.amount                  AS amount,
       ap.base_amount             AS base_amount,
       ap.base_currency           AS base_currency,
       ap.booking_id              AS booking_id,
       ap.currency                AS currency,
       ap.external_identifier     AS external_identifier,
       ap.product_type            AS product_type,
       ap.activated               AS activated,
       ap.base_commission         AS base_commission,
       ap.base_vat_on_commission  AS base_vat_on_commission,
       ap.commission              AS commission,
       ap.vat_on_commission       AS vat_on_commission,
       ap.additional_info         AS additional_info,
       ap.base_vat_on_amount      AS base_vat_on_amount,
       ap.vat_on_amount           AS vat_on_amount
FROM   secretescapes.ancillary_product ap
       -- Get booking dates --
       LEFT JOIN secretescapes.booking b
              ON b.id = ap.booking_id
       -- Report filters --
WHERE  b.status != 'ABANDONED'
       AND Month(b.date_created) = 9
       AND Year(b.date_created) = 2019
ORDER  BY ap.booking_id;



-- SE CMS Voucher Report --
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
WHERE  (v.status != 'ABANDONED' OR v.status != 'NEW')
       AND Year(v.date_created) = 2019
GROUP  BY v.id
ORDER  BY v.id;



# SE CMS Abandoned Booking Summary V1
SELECT
       b.id                               AS booking_id,
       b.unique_transaction_reference     AS order_code,
       b.user_id                          AS user_id,
       b.affiliate_user_id                AS affiliate_user_id,
       b.payment_id                       AS payment_id,
       b.date_created                     AS date_created,
       b.completion_date                  AS completion_date,
       b.last_updated                     AS last_updated,
       b.type                             AS booking_type,
       b.status                           AS booking_status,
       b.hold_id                          AS hold_id,
       b.payment_amount_type              AS payment_type,
       p.status                           AS payment_status,
       p.type                             AS payment_method,
       t.name                             AS territory,
       b.currency                         AS currency,
       b.booking_fee                      AS booking_fee,
       b.atol_fee                         AS atol_fee,
       Sum(p.surcharge)                   AS payment_surcharge,
       Sum(c.amount)                      AS credit_amount,
       Sum(p.amount)                      AS payment_amount
FROM   secretescapes.booking b
       -- Get payment information --
       LEFT JOIN secretescapes.payment p
              ON p.id = b.payment_id
       LEFT JOIN secretescapes.booking_credit bc
              ON bc.booking_credits_used_id = b.id
       LEFT JOIN secretescapes.credit c
              ON c.id = bc.credit_id
       -- Get user information --
       LEFT JOIN secretescapes.affiliate_user au
              ON au.id = b.affiliate_user_id
       LEFT JOIN secretescapes.shiro_user su
              ON su.id = b.user_id
       -- Get territory information
       LEFT JOIN secretescapes.affiliate a
              ON a.id = su.affiliate_id OR a.id = au.affiliate_id
       LEFT JOIN secretescapes.territory t
              ON t.id = a.territory_id
       -- Report filters --
WHERE  b.status = 'ABANDONED'
       AND (p.type IS NOT NULL
       AND p.type != 'ZERO_DEPOSIT')
       AND Month(b.date_created) = 8
       AND Year(b.date_created) = 2019
GROUP  BY booking_id
ORDER BY booking_id;



-- Experiment to explore retrieving Allocation and Sale IDs for Booking Summary --
-- This seems to work, but unsure of the cardinality of the relationship --
SELECT
       b.id              AS booking_id,
       b.date_created,
       b.last_updated,
       b.status,
       b.type,
       ba.allocation_id  AS allocation_id,
       a.offer_id        AS offer_id,
       CONCAT_WS('-', a.offer_id, ba.allocation_id, b.id)
FROM secretescapes.booking b
       LEFT JOIN secretescapes.booking_allocations ba
              ON ba.booking_allocations_id = b.id
       LEFT JOIN secretescapes.allocation a
              ON a.id = ba.allocation_id
WHERE Year(b.date_created) = 2019
      AND Month(b.date_created) = 8
LIMIT 100;
