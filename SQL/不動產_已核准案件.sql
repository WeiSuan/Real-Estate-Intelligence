use shin_monthly;

select distinct
	pe.pre_examine_no, 
	--main_no,
	cast(pe.receive_date as date) as receive_date, 
	cast(pe.dispatch_date as date) as dispatch_date,
	cast(pe.approval_date as date) as approval_date,

	case when pe.status = 'CREDIT_070_1' then '(1) 待核准'
		 when pe.status = 'CREDIT_070_2' then '(2) 處理中'
		 when pe.status = 'CREDIT_070_3' then '(4) 核准，業管確認中'
		 when pe.status = 'CREDIT_070_4' then '(4) 強碰'
		 when pe.status = 'CREDIT_070_5' then '(5) 緩議'
		 when pe.status = 'CREDIT_070_6' then '(6) 轉入群組進件'
		 when pe.status = 'CREDIT_070_7' then '(7) 補件'
		 when pe.status = 'CREDIT_070_8' then '(8) 核准，條件變更中'
		 when pe.status = 'CREDIT_070_9' then '(9) 業退'
		 when pe.status = 'CREDIT_070_10' then '(10) 核准'
		 when pe.status = 'CREDIT_070_11' then '(11) 申覆'
		 else NULL end as status_desc, 

	approved_amount,

	btype.bus_name,
	btype2.bus_name as bus_name2,
	code_id_name as pre_code_id_name, 
	case when code_id_name = '土地擔保週轉金' then '(1)土融'
		 when code_id_name = '建築原物料週轉金' then '(2)建融'
		 when code_id_name = '土地擔保/建築原物料週轉金' then '(3)土建融'
		 when code_id_name = '餘屋融資' then '(4)餘屋'
		 else '(5)其他' end as pre_code_type,
		 
	ct.id_no as customer_id_no, 
	ct.name as customer_name, 
	
	qt.qtgroup_id, 
	qt.qtgroup_name,

	case when ln_ever.examine_group_id is not null then 1 else 0 end as loan_activate,
	total_loan_capital, 
	total_loan_remain_capital

from pre_examine pe

left join examine_group eg
	on eg.pre_examine_id = pe.pre_examine_id

left join business_type_parameter btype
	on btype.bus_type = pe.pre_bus_type

left join (select distinct code_id, code_id_name from code_detail where is_activate = '1') cd_gage
	on cd_gage.code_id = eg.industry

left join business_type_parameter2 btype2
	on btype2.bus_type = pe.pre_bus_type2

left join business_type_parameter4 btype4
	on btype4.bus_type = pe.pre_bus_type4

left join (select * from examine_applicant where owner_type = 'PreExamine') ea
	on ea.owner_id = pe.pre_examine_id

left join customer ct
	on ct.customer_id = ea.customer_id

left join quota_group qt
	on qt.quota_group_id = ct.quota_group_id

left join (
	select distinct ex.examine_group_id, 1 as ever_loan
	from loan ln 
	inner join examine ex on ex.examine_id = ln.examine_id) ln_ever
	on ln_ever.examine_group_id = eg.examine_group_id

left join (
	select distinct
		pre_examine_no, 
		sum(loan_capital) as total_loan_capital, 
		sum(loan_remCapital) as total_loan_remain_capital

	from (
		select distinct
			pe_lr.pre_examine_id,
			coalesce(pe_lr.main_no, pe_lr.pre_examine_no) as pre_examine_no, 
			ln_lr.loan_no,
			lr.loan_capital, 
			lr.loan_remCapital

		from loan ln_lr

		inner join examine ex_lr
			on ex_lr.examine_id = ln_lr.examine_id and ex_lr.bus_type2 = 'Ｈ'

		left join (
			select loan_id, 
				SUM(case when k.code='001' then lr.total_amount_receivable else 0 end) as loan_capital,
				SUM(case when k.code='001' then lr.total_amount_receivable-lr.total_amount_received else 0 end) as loan_remCapital
			
			from loan_receivable lr
			inner join customer_receive_kind k on lr.customer_receive_kind_id=k.customer_receive_kind_id
			group by lr.loan_id) lr
			on lr.loan_id = ln_lr.loan_id

		left join examine_group eg_lr
			on eg_lr.examine_group_id = ex_lr.examine_group_id

		left join pre_examine pe_lr
			on pe_lr.pre_examine_id = eg_lr.pre_examine_id
		
		where pe_lr.pre_examine_id is not null) major_lr
		group by pre_examine_no
	) lr_result
	on lr_result.pre_examine_no = pe.pre_examine_no

where pe.pre_bus_type2 = 'H' and
	main_no is null and
	pe.status not in (
	'CREDIT_070_1',	-- 待核准
	'CREDIT_070_2', -- 處理中
	'CREDIT_070_3',	-- 核准，業管確認中
	'CREDIT_070_5',	-- 緩議
	'CREDIT_070_9'	-- 業退
	) 
	and 
	(
		cast(pe.approval_date as date) >= '2025-01-01'
		or (
		cast(pe.approval_date as date) < '2025-01-01' and
		total_loan_remain_capital > 0
		)
		or
		(
		(cast(pe.approval_date as date) < '2025-01-01' or pe.approval_date is null) and
		cd_gage.code_id_name = '建築原物料週轉金') 
	);
