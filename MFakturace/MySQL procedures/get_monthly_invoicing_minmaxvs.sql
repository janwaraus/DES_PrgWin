-- získání maximálního a minimálního VS zákazníků, kterým se má fakturovat, pro účely nastavení maximálního rozmezí od-do

DELIMITER //
CREATE PROCEDURE get_monthly_invoicing_minmaxvs(
    IN in_month_start_date date,
    IN in_month_end_date date
)
BEGIN
SELECT
	MIN(variable_symbol) as min_vs,
	MAX(variable_symbol) as max_vs
FROM (
	SELECT id, contract_id, period 
	FROM billing_batches b1 
	WHERE from_date = 
	(
		SELECT MAX(from_date) 
		FROM billing_batches b2 
		WHERE b2.from_date <= in_month_end_date 
		AND b1.contract_id = b2.contract_id
	)
) AS bb
JOIN billing_items bi ON bi.billing_batch_id = bb.id
JOIN contracts co ON co.id = bb.contract_id
JOIN customers cu ON cu.id = co.customer_id 
LEFT JOIN tariffs ON co.tariff_id = tariffs.id
WHERE 
	bb.period = 1
	AND IF(co.credit, 1, 0) = 0
	AND IF(tariffs.credit, 1, 0) = 0	
	AND ((co.invoice = 1) or (co.state = 'canceled' AND co.canceled_at >= in_month_start_date)) 
	AND (co.invoice_from IS NULL OR co.invoice_from <= in_month_end_date) 
	AND co.activated_at <= in_month_end_date; 	
END //
DELIMITER ;


CALL get_monthly_invoicing_minmaxvs('2013-01-01','2013-01-01')

CALL get_monthly_invoicing_minmaxvs('2023-05-01','2023-05-31')

drop procedure get_monthly_invoicing_minmaxvs