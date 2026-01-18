-- 03_aggregation.sql
-- Aggregation example: total amount and counts per customer
SELECT
  c.customer_id,
  c.name AS customer_name,
  c.country,
  COUNT(o.order_id) AS orders_count,
  ROUND(SUM(o.amount), 2) AS total_amount,
  SUM(CASE WHEN o.status = 'paid' THEN 1 ELSE 0 END) AS paid_orders
FROM customers c
LEFT JOIN orders o ON o.customer_id = c.customer_id
GROUP BY c.customer_id, c.name, c.country
ORDER BY total_amount DESC;
