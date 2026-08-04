with gage_data as (
	select distinct
		
		-- [Gage 主檔]
		gage_id, 
		is_major,		-- 主擔
		quantity,		-- 數量
		inspect_amt,	-- 鑑價金額 
		store_place,	-- 標的物存放地
		ct.name as supplier_name,	-- 供應商姓名
		--ct.id_no as supplier_id_no,	-- 供應商統編

		-- [Car/CarKind 主檔] 
		brand,			-- 廠牌
		car_no,
		--engine_no,
		case when target_name is not null then replace(target_name, ',', '|') else '無資料' end as target_name,				-- 標的物名稱
		case when car_style_name is not null then replace(car_style_name, ',', '|') else '無資料' end as car_style_name,	-- 標的物型號
		case when manufacturer is not null then replace(manufacturer, ',', '|') else '無資料' end as manufacturer,			-- 標的物車體廠
		ca.make_year  -- 標的物年分
 
	from gage ga

	left join car ca
		on ga.owner_id = ca.car_id and ga.owner_type in ('Machine', 'Medical' ,'Car')

	left join car_kind ck
		on ck.car_kind_id = ca.car_kind_id

	left join car_brand cb
		on cb.car_brand_id = ck.car_brand_id
	
	left join customer ct
		on ct.customer_id = ga.supplier_id	
)

select distinct
	eg.examine_group_id,
	examine_group_no, 
	loan_no, 

	-- 客戶
	ct.id_no as customer_id_no, 
	ct.name as customer_name,

	-- 擔保品
	gd.*,


	-- 不動產設定資訊
	im.shin_setting_seq, 
	im.shin_setting_amt, 
	im.setting_date, 
	im.shin_setting_person,
	im.land_office, 
	im.worth, 
	im.comment,

	-- 不動產設定資訊(土地)
	ld.land_no, 
	ld.land_section, 
	ld.land_sub_section,

	-- 不動產設定資訊(建物)
	bd.building_no, 
	bd.building_address

from (select * from examine_group where pre_bus_type2 <> 'R') eg

inner join examine ex
	on ex.examine_group_id = eg.examine_group_id

left join customer ct
	on ct.customer_id = ex.applicant_id

inner join loan ln 
	on ln.examine_id = ex.examine_id

left join (
	select loan_id, 
		SUM(case when k.code='001' then lr.total_amount_receivable else 0 end) as loan_capital,
		SUM(case when k.code='001' then lr.total_amount_receivable-lr.total_amount_received else 0 end) as loan_remCapital
			
	from loan_receivable lr
	inner join customer_receive_kind k on lr.customer_receive_kind_id=k.customer_receive_kind_id
	group by lr.loan_id) lr
	on lr.loan_id = ln.loan_id

-- 擔保品
left join gage_data gd
	on gd.gage_id = ex.gage_id

-- 設定資訊
left join immovable_gage im_ga
	on im_ga.examine_group_id = ex.examine_group_id

left join immovable im	
	on im.immovable_id = im_ga.immovable_id

left join land ld
	on ld.immovable_id = im.immovable_id

left join building bd
	on bd.immovable_id = im.immovable_id

where 
	--(loan_remCapital > 0. and
	loan_stat <> 'CREDIT_048_D' and
	loan_type = 'NORMAL_TYPE' and
	settle_code = 'LOAN_004_5' and
	ln.is_activate = 1
order by eg.examine_group_id desc;
