-- 02_join.sql
-- Join customers with orders
SELECT
  o.order_id,
  o.order_date,
  o.amount,
  o.status,
  c.customer_id,
  c.name AS customer_name,
  c.country
FROM orders o
JOIN customers c ON c.customer_id = o.customer_id
ORDER BY o.order_date DESC, o.order_id DESC;
