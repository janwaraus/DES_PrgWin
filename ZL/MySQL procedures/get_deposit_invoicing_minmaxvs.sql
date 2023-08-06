-- kompletní select s plnou granularitou až na billing_itemy pro všechny zákazníky
-- obdoba ZLView jak měl táta (složený ještě z dvou dalších views)
-- ct.tag_id = 4 znaměná "platba předem jinak"

DELIMITER //
CREATE PROCEDURE get_deposit_invoicing_minmaxvs(
    IN in_active_date date,
    IN in_bb_date date
)
BEGIN
SELECT
	MIN(cu.variable_symbol) as min_vs,
	MAX(cu.variable_symbol) as max_vs
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
		WHERE ct.contract_id = co.id AND ct.tag_id = 4);
END //
DELIMITER ;


CALL get_deposit_invoicing_minmaxvs('2023-07-15','2023-06-30')

drop procedure get_deposit_invoicing_minmaxvs