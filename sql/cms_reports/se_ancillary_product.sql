-- SE CMS Ancillary Product Sales --
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
       AND Month(b.date_created) = 8
       AND Year(b.date_created) = 2019
ORDER  BY ap.booking_id;
