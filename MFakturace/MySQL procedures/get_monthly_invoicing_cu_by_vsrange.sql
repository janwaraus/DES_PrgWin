-- získání všech zákazníků, kterým se má fakturovat v daném měsíci
-- (na základě určení měsíce a rozpětí VS)

DELIMITER //
CREATE PROCEDURE get_monthly_invoicing_cu_by_vsrange_all(
    IN in_month_start_date date,
    IN in_month_end_date date,
    IN in_variable_symbol_from VARCHAR(255),
    IN in_variable_symbol_to VARCHAR(255)
)
BEGIN
SELECT DISTINCT 
	cu.variable_symbol AS cu_variable_symbol,
	cu.postal_mail AS cu_postal_mail,
	cu.abra_code AS cu_abra_code,
	cu.disable_mailings AS cu_disable_mailings
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
LEFT JOIN (
	SELECT co2.customer_id, count(*) AS voip_count 
	FROM contracts co2
	WHERE co2.tariff_id IN (1, 3)
	AND ((co2.Invoice = 1) OR (co2.State = 'canceled' AND co2.Canceled_at >= in_month_start_date)) 	
	GROUP BY co2.customer_id 
) AS voip_cnt ON voip_cnt.customer_id = co.customer_id   
WHERE 
	bb.Period = 1
	AND ((co.invoice = 1) or (co.state = 'canceled' AND co.canceled_at >= in_month_start_date)) 
	AND (co.invoice_from IS NULL OR co.invoice_from <= in_month_end_date) 
	AND co.activated_at <= in_month_end_date 	
	AND cu.variable_symbol >= in_variable_symbol_from
	AND cu.variable_symbol <= in_variable_symbol_to	
ORDER BY cu.variable_symbol;
END //
DELIMITER ;

CALL get_monthly_invoicing_cu_by_vsrange_all('2023-05-01','2023-05-31','1','7961230002')

DROP PROCEDURE get_monthly_invoicing_cu_by_vsrange_all

-- -- -- -- --

DELIMITER //
CREATE PROCEDURE get_monthly_invoicing_cu_by_vsrange_voip(
    IN in_month_start_date date,
    IN in_month_end_date date,
    IN in_variable_symbol_from VARCHAR(255),
    IN in_variable_symbol_to VARCHAR(255),
    IN in_get_voip INT
)
BEGIN
SELECT DISTINCT 
	cu.variable_symbol AS cu_variable_symbol,
	cu.postal_mail AS cu_postal_mail,
	cu.abra_code AS cu_abra_code,
	cu.disable_mailings AS cu_disable_mailings
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
LEFT JOIN (
	SELECT co2.customer_id, count(*) AS voip_count 
	FROM contracts co2
	WHERE co2.tariff_id IN (1, 3)
	AND ((co2.Invoice = 1) OR (co2.State = 'canceled' AND co2.Canceled_at >= in_month_start_date)) 	
	GROUP BY co2.customer_id 
) AS voip_cnt ON voip_cnt.customer_id = co.customer_id   
WHERE 
	bb.Period = 1
	AND ((co.invoice = 1) or (co.state = 'canceled' AND co.canceled_at >= in_month_start_date)) 
	AND (co.invoice_from IS NULL OR co.invoice_from <= in_month_end_date) 
	AND co.activated_at <= in_month_end_date 	
	AND cu.variable_symbol >= in_variable_symbol_from
	AND cu.variable_symbol <= in_variable_symbol_to	
	AND (IF (voip_cnt.voip_count,1,0)) = 1
ORDER BY cu.variable_symbol;
END //
DELIMITER ;

CALL get_monthly_invoicing_cu_by_vsrange_voip('2023-05-01','2023-05-31','1','7961230002')

DROP PROCEDURE get_monthly_invoicing_cu_by_vsrange_nonvoip

-- -- -- -- --

DELIMITER //
CREATE PROCEDURE get_monthly_invoicing_cu_by_vsrange_nonvoip(
    IN in_month_start_date date,
    IN in_month_end_date date,
    IN in_variable_symbol_from VARCHAR(255),
    IN in_variable_symbol_to VARCHAR(255)
)
BEGIN
SELECT DISTINCT 
	cu.variable_symbol AS cu_variable_symbol,
	cu.postal_mail AS cu_postal_mail,
	cu.abra_code AS cu_abra_code,
	cu.disable_mailings AS cu_disable_mailings
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
LEFT JOIN (
	SELECT co2.customer_id, count(*) AS voip_count 
	FROM contracts co2
	WHERE co2.tariff_id IN (1, 3)
	AND ((co2.Invoice = 1) OR (co2.State = 'canceled' AND co2.Canceled_at >= in_month_start_date)) 	
	GROUP BY co2.customer_id 
) AS voip_cnt ON voip_cnt.customer_id = co.customer_id   
WHERE 
	bb.Period = 1
	AND ((co.invoice = 1) or (co.state = 'canceled' AND co.canceled_at >= in_month_start_date)) 
	AND (co.invoice_from IS NULL OR co.invoice_from <= in_month_end_date) 
	AND co.activated_at <= in_month_end_date 	
	AND cu.variable_symbol >= in_variable_symbol_from
	AND cu.variable_symbol <= in_variable_symbol_to	
	AND (IF (voip_cnt.voip_count,1,0)) = 0
ORDER BY cu.variable_symbol;
END //
DELIMITER ;

CALL get_monthly_invoicing_cu_by_vsrange_nonvoip('2023-05-01','2023-05-31','1','7961230002')

DROP PROCEDURE get_monthly_invoicing_cu_by_vsrange_nonvoip

-- -- -- -- --