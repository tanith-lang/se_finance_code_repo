-- Affiliate User Bookings V1
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