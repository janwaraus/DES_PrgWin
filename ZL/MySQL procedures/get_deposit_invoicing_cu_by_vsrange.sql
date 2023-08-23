-- kompletní select s plnou granularitou až na billing_itemy pro všechny zákazníky
-- obdoba ZLView jak měl táta (složený ještě z dvou dalších views)
-- ct.tag_id = 4 znaměná "platba předem jinak"

DELIMITER //
CREATE PROCEDURE get_deposit_invoicing_cu_by_vsrange(
    IN in_active_date date,
    IN in_bb_date date,
    IN in_variable_symbol_from VARCHAR(255),
    IN in_variable_symbol_to VARCHAR(255),
    IN in_return_6m_period BOOLEAN,
    IN in_return_12m_period BOOLEAN
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
		WHERE b2.from_date <= in_bb_date 
		AND b1.contract_id = b2.contract_id
	)
) AS bb
JOIN billing_items bi ON bi.billing_batch_id = bb.id
JOIN contracts co ON co.id = bb.contract_id
JOIN customers cu ON cu.id = co.customer_id 
WHERE 
	bb.period > 1
	AND co.invoice = 1
	AND (co.invoice_from IS NULL OR co.invoice_from <= in_active_date) 
	AND co.activated_at <= in_active_date 	
    AND NOT EXISTS (SELECT ct.contract_id FROM contracts_tags ct
		WHERE ct.contract_id = co.id AND ct.tag_id = 4)
	AND cu.variable_symbol >= in_variable_symbol_from
	AND cu.variable_symbol <= in_variable_symbol_to		
	AND (
		(bb.period = 3) 
		OR (in_return_6m_period AND bb.period = 6) 
		OR (in_return_12m_period AND bb.period = 12)
	) 	
ORDER BY cu.variable_symbol;
END //
DELIMITER ;


CALL get_deposit_invoicing_cu_by_vsrange('2023-05-01','2023-05-31','1','7961230002',false,true)

drop procedure get_deposit_invoicing_cu_by_vsrange