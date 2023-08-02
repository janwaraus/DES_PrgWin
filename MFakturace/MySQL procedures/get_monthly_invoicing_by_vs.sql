-- získání contractů a billing itemů zákazníka, tedy dat úptřebných k vytvoření měsíční faktury


DELIMITER //
CREATE PROCEDURE get_monthly_invoicing_by_vs(
    IN in_month_start_date date,
    IN in_month_end_date date,
    IN in_variable_symbol VARCHAR(255)
)
BEGIN
SELECT
	co.customer_id,
	bb.contract_id,
	bi.billing_batch_id,
	bi.id AS billing_item_id,
	IFNULL (voip_cnt.voip_count, 0) AS cu_voip_count,
	IF (voip_cnt.voip_count, 1, 0) AS cu_has_active_voip,
	co.customer_id AS co_customer_id ,
	co.`number` AS co_number,
	co.`type` AS co_type,
	co.tariff_id AS co_tariff_id,
	co.invoice AS co_invoice, 
	co.activated_at AS co_activated_at,
	co.canceled_at AS co_canceled_at,
	co.invoice_from AS co_invoice_from,
	co.ctu_category AS co_ctu_category, 
	bb.period AS bb_period,
	bi.description AS bi_description,
	bi.price AS bi_price,
	bi.tariff AS bi_is_tariff,
	bi.vat_id AS bi_vat_id, 
	cu.variable_symbol AS cu_variable_symbol,
	cu.postal_mail AS cu_postal_mail,
	cu.abra_code AS cu_abra_code,
	cu.disable_mailings AS cu_disable_mailings,
	tariffs.name AS tariff_name,
	cb1.name AS cu_invoice_sending_method_name,
	cb2.name AS bi_vat_name
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
LEFT JOIN codebooks cb1 ON cu.invoice_sending_method_id = cb1.id 
LEFT JOIN codebooks cb2 ON bi.vat_id = cb2.id 
LEFT JOIN (
	SELECT co2.customer_id, COUNT(*) AS voip_count 
	FROM contracts co2
	WHERE co2.tariff_id IN (1, 3)
	AND ((co2.invoice = 1) OR (co2.state = 'canceled' AND co2.canceled_at >= in_month_start_date)) 	
	GROUP BY co2.customer_id 
) AS voip_cnt ON voip_cnt.customer_id = co.customer_id   
WHERE 
	bb.period = 1
	AND IF(co.credit, 1, 0) = 0
	AND IF(tariffs.credit, 1, 0) = 0	
	AND ((co.invoice = 1) or (co.state = 'canceled' AND co.canceled_at >= in_month_start_date)) 
	AND (co.invoice_from IS NULL OR co.invoice_from <= in_month_end_date) 
	AND co.activated_at <= in_month_end_date
	AND cu.variable_symbol = in_variable_symbol
ORDER BY co.`number`, bi.tariff DESC;
END //
DELIMITER ;


CALL get_monthly_invoicing_by_vs('2023-05-01','2023-05-31','20020004')

drop procedure get_monthly_invoicing_by_vs