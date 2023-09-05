-- kompletní select s plnou granularitou až na billing_itemy pro všechny zákazníky
-- obdoba ZLView jak měl táta (složený ještě z dvou dalších views)
-- ct.tag_id = 4 znaměná "platba předem jinak"

DELIMITER //
CREATE PROCEDURE get_deposit_invoicing_by_vs(
    IN in_active_date date,
    IN in_bb_date date,
    IN in_variable_symbol VARCHAR(255),
    IN in_return_6m_period BOOLEAN,
    IN in_return_12m_period BOOLEAN
)
BEGIN
SELECT
	co.customer_id,
	bb.contract_id,
	bi.billing_batch_id,
	bi.id AS billing_item_id,
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
	cb1.name AS cu_invoice_sending_method_name,
	cb2.name AS bi_vat_name
FROM (
	SELECT id, contract_id, period 
	FROM billing_batches b1 
	WHERE from_date = 
	(
		SELECT MAX(from_date) 
		FROM billing_batches b2 
		WHERE b2.from_date <= in_bb_date 
		AND b1.contract_id = b2.contract_id
	)
) AS bb
JOIN billing_items bi ON bi.billing_batch_id = bb.id
JOIN contracts co ON co.id = bb.contract_id
JOIN customers cu ON cu.id = co.customer_id 
LEFT JOIN codebooks cb1 ON cu.invoice_sending_method_id = cb1.id 
LEFT JOIN codebooks cb2 ON bi.vat_id = cb2.id 
WHERE 
	bb.period > 1
	AND co.invoice = 1
	AND (co.invoice_from IS NULL OR co.invoice_from <= in_active_date) 
	AND co.activated_at <= in_active_date 	
    AND NOT EXISTS (SELECT ct.contract_id FROM contracts_tags ct
		WHERE ct.contract_id = co.id AND ct.tag_id = 4)
	AND cu.variable_symbol = in_variable_symbol
	AND (
		(bb.period = 3) 
		OR (in_return_6m_period AND bb.period = 6) 
		OR (in_return_12m_period AND bb.period = 12)
	) 	
ORDER BY bb.period, co.`number`, bi.tariff DESC;
END //
DELIMITER ;


CALL get_deposit_invoicing_by_vs('2023-05-01','2023-05-31','20020027')

drop procedure get_deposit_invoicing_by_vs