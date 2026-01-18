-- 01_simple_select.sql
-- Basic select from a single table (customers)
SELECT
  customer_id,
  name,
  country,
  created_at
FROM customers
ORDER BY created_at DESC;
