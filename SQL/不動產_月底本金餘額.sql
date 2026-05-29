use shin_monthly;

select distinct
	
	pe.pre_examine_no, 
	main_no, 
	coalesce(main_no, pe.pre_examine_no) as pre_examine_no_major, 

	examine_group_no,
	ct.id_no as customer_id_no,
	ct.name as customer_name, 

	loan_no, 
	cast(sales_date as date) as sales_date, 
	case when loan_type = 'NORMAL_TYPE' then '正式合約' else '暫付合約' end as loan_type,
	loan_capital, 
	loan_remain_capital, 
	loan_term as 期數, 
	case when loan_stat = 'CREDIT_048_N' then '正常'
		 when loan_stat = 'CREDIT_048_E' then 'NIEE'
		 when loan_stat = 'CREDIT_048_D' then '呆帳' end as loan_stat, 

	btype.bus_name as loan_bus_name,
	cd_gage.code_id_name, 
	case when cd_gage.code_id_name = '土地擔保週轉金' then '(1)土融'
		 when cd_gage.code_id_name = '建築原物料週轉金' then '(2)建融'
		 when cd_gage.code_id_name = '土地擔保/建築原物料週轉金' then '(3)土建融'
		 when cd_gage.code_id_name = '餘屋融資' then '(4)餘屋'
		 else '(5)其他' end as loan_code_type, 
	case when ch.examine_group_id is not null then 1 else 0 end as loan_extend, 
	ga.inspect_amt as immovable_worth_system

from loan ln

inner join examine ex 
	on ex.examine_id = ln.examine_id and ex.bus_type2 = 'H'

left join examine_group eg
	on eg.examine_group_id = ex.examine_group_id

left join pre_examine pe
	on pe.pre_examine_id = eg.pre_examine_id

-- [業務別代碼(操作方式)]
left join business_type_parameter btype
	on btype.bus_type = ex.bus_type

-- [業別第一碼(擔保品)]
left join (select distinct code_id, code_id_name from code_detail where is_activate = '1') cd_gage
		on cd_gage.code_id = eg.industry

-- [業別第二碼(產品別)]
left join business_type_parameter2 btype2
	on btype2.bus_type = ex.bus_type2

-- [業別第三碼(性質)]
left join business_type_parameter4 btype4
	on btype4.bus_type = ex.bus_type4

-- 期數
left join (select distinct loan_id, max(inst_no) as loan_term from loan_period group by loan_id) lp
	on lp.loan_id = ln.loan_id

-- 授信本金/本金餘額
left join (
	select loan_id, 
		SUM(case when k.code='001' then lr.total_amount_receivable else 0 end) as loan_capital,
		SUM(case when k.code='001' then lr.total_amount_receivable-lr.total_amount_received else 0 end) as loan_remain_capital
			
	from loan_receivable lr
	inner join customer_receive_kind k on lr.customer_receive_kind_id=k.customer_receive_kind_id
	group by lr.loan_id) lr
	on lr.loan_id = ln.loan_id

-- 申購人
left join customer ct
	on ct.customer_id = ex.applicant_id

-- 產業別
left join industry_parameter ip
	on ip.industry_parameter_id = ct.industry_parameter_id

left join code_detail cdd
	on cdd.native_code_id = ip.industry_category and cdd.code_item = 'INDCAT'

left join change_dlr_br ch
	on ch.examine_group_id = eg.examine_group_id

left join gage ga
	on ga.gage_id = ex.gage_id

left join car ca
	on ga.owner_id = ca.car_id and ga.owner_type in ('Machine', 'Medical' ,'Car')

where loan_stat <> 'CREDIT_048_D' and
	loan_type = 'NORMAL_TYPE' and
	settle_code = 'LOAN_004_5' and
	ln.is_activate = 1;