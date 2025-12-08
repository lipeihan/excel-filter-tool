import pandas as pd
from datetime import datetime, timedelta

def filter_bonus_data():
    input_file = '输入数据.xlsx'
    output_template = '输出数据.xlsx'
    result_file = '筛选结果.xlsx'
    
    # 1. Load Data
    print("Loading data...")
    try:
        # Try to load '工时数据' first, then '累计工时'
        xl = pd.ExcelFile(input_file)
        sheet_names = xl.sheet_names
        
        main_sheet_name = None
        if '工时数据' in sheet_names:
            main_sheet_name = '工时数据'
        elif '累计工时' in sheet_names:
            main_sheet_name = '累计工时'
        else:
            print("Error: Could not find '工时数据' or '累计工时' sheet.")
            return

        print(f"Using main data sheet: {main_sheet_name}")
        df_hours = pd.read_excel(input_file, sheet_name=main_sheet_name)
        df_filter = pd.read_excel(input_file, sheet_name='筛选条件')
        df_certs = pd.read_excel(input_file, sheet_name='过岗数据')
        df_basic = pd.read_excel(input_file, sheet_name='基本数据')
        df_managers = pd.read_excel(input_file, sheet_name='门店负责人')
        df_status = pd.read_excel(input_file, sheet_name='门店状态表')
        
        # Load template columns
        df_template = pd.read_excel(output_template, nrows=0)
        output_cols = df_template.columns.tolist()
        
    except Exception as e:
        print(f"Error loading files: {e}")
        return

    # Determine Bonus Month
    # Look for '奖金月份' column in df_filter
    bonus_month_str = "2025-11" # Default
    if '奖金月份' in df_filter.columns:
        first_val = df_filter['奖金月份'].dropna().iloc[0] if not df_filter['奖金月份'].dropna().empty else None
        if first_val:
            if isinstance(first_val, datetime):
                bonus_month_str = first_val.strftime('%Y-%m')
            else:
                bonus_month_str = str(first_val).strip()
    
    try:
        # Parse YYYY-MM and set to 1st of that month
        BONUS_MONTH_START = datetime.strptime(bonus_month_str, '%Y-%m')
        print(f"Calculating bonus for month starting: {BONUS_MONTH_START.date()}")
    except ValueError:
        print(f"Invalid Bonus Month format: {bonus_month_str}. Using default 2025-11-01.")
        BONUS_MONTH_START = datetime(2025, 11, 1)

    # 2. Apply Filter (筛选条件)
    # If filter sheet is not empty, keep only matching rows in df_hours
    if not df_filter.empty and not df_filter.dropna(how='all').empty:
        print("Applying filters from '筛选条件'...")
        
        # --- PREPARE DATA FOR FILTERING ---
        # Create a combined dataframe that has columns from all relevant sheets
        # mapped with prefix "SheetName-"
        
        print("Preparing combined data for filtering...")
        df_combined = df_hours.copy()
        
        # 1. Merge Basic Data (基本数据)
        # Prefix columns
        df_basic_prefixed = df_basic.add_prefix('基本数据-')
        # Merge on Employee ID
        # df_hours['工号'] <-> df_basic['工号']
        if '工号' in df_combined.columns and '基本数据-工号' in df_basic_prefixed.columns:
            df_combined = df_combined.merge(df_basic_prefixed, left_on='工号', right_on='基本数据-工号', how='left')
        
        # 2. Merge Store Status (门店状态表)
        # df_hours['门店编码'] <-> df_status['ERP门店编码']
        df_status_prefixed = df_status.add_prefix('门店状态表-')
        if '门店编码' in df_combined.columns and '门店状态表-ERP门店编码' in df_status_prefixed.columns:
            df_combined = df_combined.merge(df_status_prefixed, left_on='门店编码', right_on='门店状态表-ERP门店编码', how='left')
            
        # 3. Merge Store Managers (门店负责人)
        # df_hours['门店编码'] <-> df_managers['部门编号']
        df_managers_prefixed = df_managers.add_prefix('门店负责人-')
        if '门店编码' in df_combined.columns and '门店负责人-部门编号' in df_managers_prefixed.columns:
             df_combined = df_combined.merge(df_managers_prefixed, left_on='门店编码', right_on='门店负责人-部门编号', how='left')

        # 4. Rename original df_hours columns to '{main_sheet_name}-'
        # We do this last so we don't break the join keys above
        df_combined = df_combined.rename(columns={col: f'{main_sheet_name}-{col}' for col in df_hours.columns})
        
        # --- APPLY FILTERS ---
        
        # Get valid columns from filter sheet that also exist in combined data
        valid_filter_cols = [col for col in df_filter.columns if col in df_combined.columns]
        
        if not valid_filter_cols:
            print(f"No matching columns found. Available columns in data: {list(df_combined.columns)[:10]}...")
            print("Ignoring filter.")
        else:
            print(f"Using filter columns: {valid_filter_cols}")
            
            final_mask = pd.Series(False, index=df_combined.index)
            valid_filter_rows = df_filter.dropna(how='all')
            
            for _, filter_row in valid_filter_rows.iterrows():
                # Create a mask for this specific filter rule
                rule_mask = pd.Series(True, index=df_combined.index)
                
                for col in valid_filter_cols:
                    val = filter_row[col]
                    # Skip if the filter value itself is NaN/Empty for this row
                    if pd.isna(val):
                        continue
                        
                    rule_mask &= (df_combined[col] == val)
                
                final_mask |= rule_mask
            
            # Filter original df_hours using the mask from df_combined
            # (They share the same index because we used left join)
            df_hours = df_hours[final_mask].copy()
            print(f"Rows after filtering: {len(df_hours)}")
    else:
        print("No filters found in '筛选条件'. Using all data.")

    if df_hours.empty:
        print("No data left after filtering.")
        return

    # 3. Prepare Helper Data
    
    # Certifications: Group by Employee ID and collect list of valid certs with dates
    # Valid Certs: Status == '有效'
    # valid_certs structure: {emp_id: {cert_name: effective_date}}
    valid_certs = {}
    if '生效日期' in df_certs.columns:
        for _, row in df_certs[df_certs['状态'] == '有效'].iterrows():
            eid = row['工号']
            cname = row['证书名称']
            cdate = row['生效日期']
            if pd.isna(cdate):
                continue
            if eid not in valid_certs:
                valid_certs[eid] = {}
            valid_certs[eid][cname] = cdate
    else:
        print("Warning: '生效日期' column not found in '过岗数据'. Using certificate existence only (ignoring date).")
        # Fallback: Use a very old date so date check always passes
        for _, row in df_certs[df_certs['状态'] == '有效'].iterrows():
            eid = row['工号']
            cname = row['证书名称']
            if eid not in valid_certs:
                valid_certs[eid] = {}
            valid_certs[eid][cname] = datetime(2000, 1, 1)

    target_certs = {'【奈雪】大堂服务岗证书', '【奈雪】后厨岗证书', '【奈雪】水吧岗证书'}

    # Entry Dates: Map ID to Entry Date
    entry_dates = df_basic.set_index('工号')['入职日期'].to_dict()
    
    # Store Managers: Map (StoreID, EmpID) to Boolean (True if manager of that store)
    # Check '门店负责人' sheet. '部门编号' is store code, '店长' is EmpID.
    # We can create a set of tuples (StoreCode, EmpID)
    manager_set = set(zip(df_managers['部门编号'], df_managers['店长']))

    # 4. Logic Processing
    eligible_rows = []

    for idx, row in df_hours.iterrows():
        emp_id = row.get('工号')
        name = row.get('姓名')
        job_title = str(row.get('职位名称', '')).strip()
        store_code = row.get('门店编码')
        
        monthly_hours = row.get('考勤工时', 0)
        total_hours = row.get('总工时', 0)
        
        is_eligible = False
        reason = ""

        # Helper to check certs
        # Returns True if any target cert was obtained BEFORE Bonus Month
        def has_valid_cert(eid):
            if eid not in valid_certs:
                return False
            user_certs = valid_certs[eid]
            for t_cert in target_certs:
                if t_cert in user_certs:
                    # Check date: Cert Date < Bonus Month Start
                    # "符合条件的次月参与" -> Cert Date must be in previous month or earlier
                    # effectively: Cert Date < Start of Bonus Month
                    if user_certs[t_cert] < BONUS_MONTH_START:
                        return True
            return False

        # --- Rule 1: Tea Master (茶饮师, 茶饮师（S）) ---
        if job_title in ['茶饮师', '茶饮师（S）']:
            if has_valid_cert(emp_id):
                is_eligible = True
                reason = "Tea Master with Valid Cert (Pre-Bonus Month)"
            else:
                reason = "Tea Master: Missing Cert or Cert too new"

        # --- Rule 2: Part-time (兼职) ---
        # Assumption: Job title contains '兼职'
        elif '兼职' in job_title:
            # Condition 1: Total Hours >= 40
            # Condition 2: Has Valid Cert (Pre-Bonus Month)
            # "以上条件都满足，次月参与" -> Assuming Total Hours is also a "status" check
            # Condition 3: Monthly Hours >= 50 (Current Month)
            
            cond_hours_cumulative = (total_hours >= 40)
            cond_cert = has_valid_cert(emp_id)
            cond_monthly_hours = (monthly_hours >= 50)
            
            if cond_hours_cumulative and cond_cert:
                if cond_monthly_hours:
                    is_eligible = True
                    reason = "Part-time Eligible"
                else:
                    reason = "Part-time: Monthly hours < 50"
            else:
                reason = f"Part-time: Total Hours<40 ({total_hours}) or Missing/New Cert"

        # --- Rule 3: Assistant Manager/Store Manager (副经理, 副店长) ---
        elif job_title in ['副经理', '副店长']:
            entry_date = entry_dates.get(emp_id)
            if pd.notna(entry_date):
                # "入职满30天的次月参加分配"
                # New Formula: (Entry + 29 days) < Start of Bonus Month
                cutoff_date = entry_date + timedelta(days=29)
                if cutoff_date < BONUS_MONTH_START:
                    is_eligible = True
                    reason = "Assistant Manager Eligible"
                else:
                    reason = f"Assistant Manager: Not seasoned enough ({cutoff_date.date()} >= {BONUS_MONTH_START.date()})"
            else:
                reason = "Assistant Manager: Missing Entry Date"

        # --- Rule 4: Store Manager (店长, 店长（S）) ---
        elif job_title in ['店长', '店长（S）']:
            # "是否曾经带过店" -> Check if they are in manager_set matching THIS store
            if (store_code, emp_id) in manager_set:
                is_eligible = True
                reason = "Store Manager Eligible"
            else:
                reason = "Store Manager: Not managing this store"
        
        else:
            reason = f"Role '{job_title}' not covered by rules"

        if is_eligible:
            # Add to result
            eligible_rows.append(row)
            # print(f"kept: {emp_id} {name} ({job_title}) - {reason}")
        else:
            pass
            # print(f"dropped: {emp_id} {name} ({job_title}) - {reason}")

    # 5. Construct Output
    print(f"Eligible employees found: {len(eligible_rows)}")
    
    if not eligible_rows:
        print("No eligible employees found.")
        return

    df_result_source = pd.DataFrame(eligible_rows)
    
    # Map to Output Columns
    # We need to pull data from various sources to fill the output columns
    # Output Columns: ['工号', '姓名', '身份证信息', '门店编码', '部门', '第三方', '工作地区', '职位', '入职日期', '转正日期', '离职日期', '组织类型', '所属区域', '负责人', '开业时间', '闭店时间', '工时', '年假小时数', '总工时', '是否门店负责人']
    
    final_data = []
    
    # Prepare lookups
    basic_lookup = df_basic.set_index('工号').to_dict('index')
    status_lookup = df_status.set_index('ERP门店编码').to_dict('index') # Assuming 'ERP门店编码' matches '门店编码'
    manager_lookup = df_managers.set_index('部门编号').to_dict('index')
    
    for _, row in df_result_source.iterrows():
        emp_id = row.get('工号')
        store_code = row.get('门店编码')
        
        basic_info = basic_lookup.get(emp_id, {})
        store_info = status_lookup.get(store_code, {})
        manager_info = manager_lookup.get(store_code, {})
        
        new_row = {}
        new_row['工号'] = emp_id
        new_row['姓名'] = row.get('姓名')
        new_row['身份证信息'] = basic_info.get('身份证号码')
        new_row['门店编码'] = store_code
        new_row['部门'] = manager_info.get('部门名称')
        new_row['第三方'] = "否" # Placeholder or logic needed? Template said "取基本数据:第三方" but Basic Data doesn't have it. Default to empty or check columns again.
        new_row['工作地区'] = basic_info.get('工作地区')
        new_row['职位'] = basic_info.get('职位') # Or use row['职位名称']? Template says "取基本数据:职位"
        new_row['入职日期'] = basic_info.get('入职日期')
        new_row['转正日期'] = basic_info.get('转正日期')
        new_row['离职日期'] = basic_info.get('离职日期')
        new_row['组织类型'] = store_info.get('品牌') # Template: "取门店状态表的：品牌"
        new_row['所属区域'] = row.get('区域')
        new_row['负责人'] = row.get('区经理')
        new_row['开业时间'] = store_info.get('开始营业') # Check column name in status table. It was '开始营业'
        new_row['闭店时间'] = store_info.get('闭店时间')
        new_row['工时'] = row.get('总工时') # Template says "取工时数据：总工时" for '工时' column? Or '考勤工时'?
                                         # Wait, Template: '工时' -> "取工时数据：总工时". '总工时' -> "判断...".
                                         # Actually in first inspect: '工时' -> "取工时数据：总工时", '总工时' -> "工时+年假小时".
                                         # Let's use '总工时' from source for '工时'.
        new_row['年假小时数'] = 0 # Default
        new_row['总工时'] = row.get('总工时') # Placeholder, user might want calculation
        
        # '是否门店负责人'
        # "判断：门店负责人在的店长，用门店编码和工号判断"
        is_manager = "是" if (store_code, emp_id) in manager_set else "否"
        new_row['是否门店负责人'] = is_manager
        
        final_data.append(new_row)
        
    df_final = pd.DataFrame(final_data, columns=output_cols)
    
    # Write to Excel
    # Format date columns to remove time part
    date_columns = ['入职日期', '转正日期', '离职日期', '开业时间', '闭店时间']
    for col in date_columns:
        if col in df_final.columns:
            # Convert to datetime first to ensure correct type
            df_final[col] = pd.to_datetime(df_final[col], errors='coerce')
            # Format as YYYY-MM-DD string to remove 00:00:00 in Excel
            # Using .dt.date would result in python object, strings are safer for Excel display without time
            df_final[col] = df_final[col].dt.strftime('%Y-%m-%d').fillna('')

    df_final.to_excel(result_file, index=False)
    print(f"Successfully generated {result_file}")

if __name__ == "__main__":
    try:
        filter_bonus_data()
    except Exception as e:
        import traceback
        traceback.print_exc()
    finally:
        input("\nPress Enter to exit...")
