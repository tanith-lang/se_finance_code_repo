-- SE CMS User Booking Report --
SELECT u.id                AS se_user_id,
       Lower(u.username)   AS user_name,
       u.date_created      AS user_join_date,
       b.date_created      AS booking_date,
       b.status            AS booking_status
FROM   shiro_user u
       -- Get booking information --
       LEFT JOIN booking b
              ON b.user_id = u.id
WHERE b.status = 'COMPLETE'
   OR b.status = 'REFUNDED';
