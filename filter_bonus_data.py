import pandas as pd
from datetime import datetime, timedelta

def robust_parse_date(x):
    """Helper to parse dates with support for Chinese format YYYY年MM月DD日"""
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, datetime):
        return x
    
    s = str(x).strip()
    # Try Chinese format first if it looks like it
    if '年' in s and '月' in s and '日' in s:
        try:
            return datetime.strptime(s, '%Y年%m月%d日')
        except ValueError:
            pass
    
    # Fallback to pandas to_datetime for other formats
    try:
        return pd.to_datetime(s)
    except:
        return pd.NaT

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
        df_hours = pd.read_excel(xl, sheet_name=main_sheet_name)
        df_filter = pd.read_excel(xl, sheet_name='筛选条件')
        df_certs = pd.read_excel(xl, sheet_name='过岗数据')
        df_basic = pd.read_excel(xl, sheet_name='基本数据')
        df_managers = pd.read_excel(xl, sheet_name='门店负责人')
        df_status = pd.read_excel(xl, sheet_name='门店状态表')
        try:
            df_roster = pd.read_excel(xl, sheet_name='花名册')
        except:
            print("Warning: Could not find '花名册' sheet. Fallback logic will be skipped.")
            df_roster = pd.DataFrame()
        
        # --- Pre-process Date Columns (Handle Chinese Dates) ---
        print("Preprocessing date columns...")
        
        # 1. Store Status Dates
        for col in ['开始营业', '闭店时间']:
            if col in df_status.columns:
                df_status[col] = df_status[col].apply(robust_parse_date)
                
        # 2. Basic Info Dates
        for col in ['入职日期', '转正日期', '离职日期']:
            if col in df_basic.columns:
                df_basic[col] = df_basic[col].apply(robust_parse_date)

        # 3. Roster Dates
        if not df_roster.empty:
            for col in ['入职日期', '转正日期', '离职日期']:
                if col in df_roster.columns:
                    df_roster[col] = df_roster[col].apply(robust_parse_date)

        # 4. Cert Dates
        if '生效日期' in df_certs.columns:
            df_certs['生效日期'] = df_certs['生效日期'].apply(robust_parse_date)
            
        # Load template columns
        df_template = pd.read_excel(output_template, nrows=0)
        output_cols = df_template.columns.tolist()
        
    except Exception as e:
        print(f"Error loading files: {e}")
        return

    # Determine Bonus Month
    # Look for '奖金月份' column in df_filter
    BONUS_MONTH_START = datetime(2025, 11, 1) # Default
    
    if '奖金月份' in df_filter.columns:
        first_val = df_filter['奖金月份'].dropna().iloc[0] if not df_filter['奖金月份'].dropna().empty else None
        
        if first_val:
            # Try robust parse first (handles YYYY-MM-DD or YYYY年MM月DD日)
            parsed_date = robust_parse_date(first_val)
            
            if pd.notna(parsed_date):
                 BONUS_MONTH_START = parsed_date.replace(day=1)
            else:
                 # Try parsing 'YYYY-MM' or 'YYYY年MM月' (without day)
                 s = str(first_val).strip()
                 try:
                     if '年' in s and '月' in s:
                         # Handle '2025年11月'
                         dt = datetime.strptime(s, '%Y年%m月')
                         BONUS_MONTH_START = dt
                     else:
                         # Handle '2025-11'
                         dt = datetime.strptime(s, '%Y-%m')
                         BONUS_MONTH_START = dt
                 except:
                     print(f"Warning: Could not parse bonus month '{s}'. Using default 2025-11-01.")

    print(f"Calculating bonus for month starting: {BONUS_MONTH_START.date()}")

    # --- Pre-calculate Aggregated Hours per Employee (Before Filtering) ---
    print("Calculating aggregated hours per employee...")
    # Group by '工号' and sum '总工时' and '考勤工时'
    # Use fillna(0) to ensure numeric summation works
    emp_agg_total = df_hours.groupby('工号')['总工时'].sum().to_dict() if '总工时' in df_hours.columns else {}
    emp_agg_monthly = df_hours.groupby('工号')['考勤工时'].sum().to_dict() if '考勤工时' in df_hours.columns else {}
    print(f"Aggregated hours calculated for {len(emp_agg_total)} employees.")

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
    # valid_certs structure: {emp_id: {cert_name: earliest_effective_date}}
    # We store the EARLIEST date for each cert type to maximize eligibility chances.
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
            
            # Keep the earliest date if multiple records exist for same cert name
            if cname in valid_certs[eid]:
                if cdate < valid_certs[eid][cname]:
                    valid_certs[eid][cname] = cdate
            else:
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

    # Define Target Cert Sets
    tea_master_certs_required = {'【奈雪】大堂服务岗证书', '【奈雪】后厨岗证书', '【奈雪】水吧岗证书'}
    
    # --- Prepare Lookups (Handle Duplicates) ---
    # Drop duplicates to ensure unique index for to_dict
    if '工号' in df_basic.columns:
        dup_basic = df_basic[df_basic.duplicated(subset=['工号'], keep=False)]
        if not dup_basic.empty:
            print(f"Warning: Found duplicate '工号' in '基本数据'. Count: {len(dup_basic)}. Keeping first occurrence.")
        df_basic = df_basic.drop_duplicates(subset=['工号'])
    basic_lookup = df_basic.set_index('工号').to_dict('index')
    
    # Verify '门店状态表' key column
    status_key = 'ERP门店编码'
    if status_key not in df_status.columns:
        # Fallback: check if '门店编码' exists
        if '门店编码' in df_status.columns:
            status_key = '门店编码'
        else:
            print(f"Warning: Could not find '{status_key}' in '门店状态表'. Available: {list(df_status.columns)}")
            status_key = None
            
    if status_key:
        dup_status = df_status[df_status.duplicated(subset=[status_key], keep=False)]
        if not dup_status.empty:
             print(f"Warning: Found duplicate '{status_key}' in '门店状态表'. Count: {len(dup_status)}. Keeping first occurrence.")
        df_status = df_status.drop_duplicates(subset=[status_key])
        status_lookup = df_status.set_index(status_key).to_dict('index')
    else:
        status_lookup = {}
        
    if '部门编号' in df_managers.columns:
        dup_managers = df_managers[df_managers.duplicated(subset=['部门编号'], keep=False)]
        if not dup_managers.empty:
            print(f"Warning: Found duplicate '部门编号' in '门店负责人'. Count: {len(dup_managers)}. Keeping first occurrence.")
        df_managers = df_managers.drop_duplicates(subset=['部门编号'])
    manager_lookup = df_managers.set_index('部门编号').to_dict('index')

    # Roster Lookup
    roster_lookup = {}
    if not df_roster.empty and '工号' in df_roster.columns:
        dup_roster = df_roster[df_roster.duplicated(subset=['工号'], keep=False)]
        if not dup_roster.empty:
            print(f"Warning: Found duplicate '工号' in '花名册'. Count: {len(dup_roster)}. Keeping first occurrence.")
        df_roster = df_roster.drop_duplicates(subset=['工号'])
        roster_lookup = df_roster.set_index('工号').to_dict('index')

    # Entry Dates: Map ID to Entry Date
    # Use the cleaned df_basic or basic_lookup
    # entry_dates = df_basic.set_index('工号')['入职日期'].to_dict() # This is now safe as df_basic is deduplicated
    # But better to use basic_lookup directly in logic
    entry_dates = {k: v.get('入职日期') for k, v in basic_lookup.items()}
    
    # Store Managers: Map (StoreID, EmpID) to Boolean (True if manager of that store)
    # Check '门店负责人' sheet. '部门编号' is store code, '店长' is EmpID.
    # We can create a set of tuples (StoreCode, EmpID)
    # Ensure types are consistent (strings)
    manager_set = set()
    for _, row in df_managers.iterrows():
        sc = row.get('部门编号')
        eid = row.get('店长')
        if pd.notna(sc) and pd.notna(eid):
            manager_set.add((str(sc).strip(), str(eid).strip()))

    # 4. Logic Processing
    eligible_rows = []
    excluded_rows = []

    for idx, row in df_hours.iterrows():
        emp_id = row.get('工号')
        name = row.get('姓名')
        
        # Job Title Logic: Prefer '基本数据' > '花名册' > '工时数据'
        basic_info = basic_lookup.get(emp_id, {})
        roster_info = roster_lookup.get(emp_id, {})

        job_title = str(basic_info.get('职位', '')).strip()
        if not job_title or job_title.lower() == 'nan':
            job_title = str(roster_info.get('职位', '')).strip()
            if not job_title or job_title.lower() == 'nan':
                job_title = str(row.get('职位名称', '')).strip()
        
        store_code = row.get('门店编码')
        
        # Ensure store_code is string for consistent lookup
        store_code_str = str(store_code).strip() if pd.notna(store_code) else ''
        
        # Use aggregated hours for logic checks
        monthly_hours = emp_agg_monthly.get(emp_id, 0)
        total_hours = emp_agg_total.get(emp_id, 0)
        
        is_eligible = False
        reason = ""

        # --- Rule 1: Tea Master & Trainers ---
        # "茶饮师 / 茶饮师（S）/pro训练员/茶饮训练员"
        if job_title in ['茶饮师', '茶饮师（S）', 'pro训练员', '茶饮训练员']:
            # Condition: Must have ALL 3 certificates
            # Judgment Date: The LATEST of the 3 certificates.
            if emp_id in valid_certs:
                user_c = valid_certs[emp_id]
                has_all = True
                dates = []
                for req_c in tea_master_certs_required:
                    if req_c not in user_c:
                        has_all = False
                        break
                    # Ensure date is valid datetime
                    c_date = user_c[req_c]
                    if isinstance(c_date, datetime):
                        dates.append(c_date)
                    else:
                        # Try to parse if string
                        parsed = robust_parse_date(c_date)
                        if pd.notna(parsed):
                            dates.append(parsed)
                        else:
                             # Invalid date counts as missing for logic safety
                             has_all = False
                             break
                
                if has_all:
                    # Logic: "证书取3个证书中最晚拿到的时间作为判断"
                    latest_cert_date = max(dates)
                    if latest_cert_date < BONUS_MONTH_START:
                        is_eligible = True
                        reason = f"Tea Master: All 3 certs acquired by {latest_cert_date.date()}"
                    else:
                        reason = f"Tea Master: Certs too new ({latest_cert_date.date()} >= {BONUS_MONTH_START.date()})"
                else:
                    reason = "Tea Master: Missing one or more required certs"
            else:
                reason = "Tea Master: No certs found"

        # --- Rule 2: Part-time & Interns ---
        # "兼职 (职位名称包含“兼职”)/就业见习生"
        elif '兼职' in job_title or job_title == '就业见习生':
            # Condition 1: Total Hours >= 40
            # Condition 2: Has ANY of the 3 certificates
            # Condition 3: Monthly Hours >= 50
            
            # Logic: "证书其中一个取最早拿到的时间作为判断"
            
            cond_hours_cumulative = (total_hours >= 40)
            cond_monthly_hours = (monthly_hours >= 50)
            
            has_any_cert = False
            earliest_cert_date = None
            
            if emp_id in valid_certs:
                user_c = valid_certs[emp_id]
                found_dates = []
                for req_c in tea_master_certs_required:
                    if req_c in user_c:
                        c_date = user_c[req_c]
                        # Ensure date
                        if isinstance(c_date, datetime):
                            found_dates.append(c_date)
                        else:
                             parsed = robust_parse_date(c_date)
                             if pd.notna(parsed):
                                 found_dates.append(parsed)
                
                if found_dates:
                    has_any_cert = True
                    earliest_cert_date = min(found_dates)
            
            cond_cert_time = False
            if has_any_cert and earliest_cert_date:
                if earliest_cert_date < BONUS_MONTH_START:
                    cond_cert_time = True
            
            if cond_hours_cumulative and cond_cert_time and cond_monthly_hours:
                is_eligible = True
                reason = "Part-time/Intern Eligible"
            else:
                reason_parts = []
                if not cond_hours_cumulative: reason_parts.append(f"TotalHours({total_hours})<40")
                if not cond_monthly_hours: reason_parts.append(f"MonthlyHours({monthly_hours})<50")
                if not has_any_cert: reason_parts.append("No Cert")
                elif not cond_cert_time: reason_parts.append(f"Cert too new ({earliest_cert_date.date()})")
                reason = "Part-time: " + ", ".join(reason_parts)

        # --- Rule 3: Assistant Manager/Store Manager (副经理, 副店长) ---
        elif job_title in ['副经理', '副店长']:
            entry_date = entry_dates.get(emp_id)
            if pd.notna(entry_date):
                # "入职满30天的次月参加分配"
                # Formula: (Entry + 29 days) < Start of Bonus Month
                cutoff_date = entry_date + timedelta(days=29)
                if cutoff_date < BONUS_MONTH_START:
                    is_eligible = True
                    reason = "Assistant Manager Eligible"
                else:
                    reason = f"Assistant Manager: Not seasoned enough ({cutoff_date.date()} >= {BONUS_MONTH_START.date()})"
            else:
                reason = "Assistant Manager: Missing Entry Date"

        # --- Rule 4: Store Manager ---
        # "店长 / 店长（S）/见习店长/资深店长"
        elif job_title in ['店长', '店长（S）', '见习店长', '资深店长']:
            # Issue 2 Fix: Store Managers should always be eligible regardless of manager_set check
            is_eligible = True
            reason = "Store Manager: Auto-eligible"
            # if (store_code, emp_id) in manager_set:
            #     is_eligible = True
            #     reason = "Store Manager Eligible"
            # else:
            #     reason = "Store Manager: Not managing this store"
        
        else:
            reason = f"Role '{job_title}' not covered by rules"

        if is_eligible:
            # Add to result
            eligible_rows.append(row)
            # print(f"kept: {emp_id} {name} ({job_title}) - {reason}")
        else:
            excluded_rows.append({
                '工号': emp_id,
                '姓名': name,
                '职位': job_title,
                '门店编码': store_code,
                '总工时': total_hours,
                '月工时': monthly_hours,
                '排除原因': reason
            })
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
    
    for _, row in df_result_source.iterrows():
        emp_id = row.get('工号')
        store_code = row.get('门店编码')
        
        basic_info = basic_lookup.get(emp_id, {})
        store_info = status_lookup.get(store_code, {})
        manager_info = manager_lookup.get(store_code, {})
        roster_info = roster_lookup.get(emp_id, {})
        
        new_row = {}
        new_row['工号'] = emp_id
        new_row['姓名'] = row.get('姓名')
        
        # Helper for fallback: if value is empty/null/nan, try roster
        def get_with_fallback(primary_dict, primary_key, roster_dict, roster_key):
            val = primary_dict.get(primary_key)
            # Check if val is effectively empty
            is_empty = False
            if pd.isna(val):
                is_empty = True
            elif isinstance(val, str) and not val.strip():
                is_empty = True
            elif str(val).lower() == 'nan':
                 is_empty = True
                 
            if is_empty:
                return roster_dict.get(roster_key)
            return val

        new_row['身份证信息'] = get_with_fallback(basic_info, '身份证号码', roster_info, '身份证')
        new_row['门店编码'] = store_code
        new_row['部门'] = manager_info.get('部门名称')
        
        # '第三方' fallback
        # Basic Data '第三方公司' might be the key?
        # My previous inspection showed '第三方公司' in Basic Data.
        new_row['第三方'] = get_with_fallback(basic_info, '第三方公司', roster_info, '第三方公司')
        
        # '工作地区' fallback
        new_row['工作地区'] = get_with_fallback(basic_info, '工作地区', roster_info, '工作城市')
        
        # '职位' fallback
        # First try Basic Data '职位', then Roster '职位', then Source '职位名称'
        val_job = get_with_fallback(basic_info, '职位', roster_info, '职位')
        if pd.isna(val_job) or (isinstance(val_job, str) and not val_job.strip()) or str(val_job).lower() == 'nan':
             val_job = row.get('职位名称')
        new_row['职位'] = val_job
        
        # Dates fallback
        new_row['入职日期'] = get_with_fallback(basic_info, '入职日期', roster_info, '入职日期')
        new_row['转正日期'] = get_with_fallback(basic_info, '转正日期', roster_info, '转正日期')
        new_row['离职日期'] = get_with_fallback(basic_info, '离职日期', roster_info, '离职日期')
        
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
        is_manager = "是" if (store_code_str, str(emp_id).strip()) in manager_set else "否"
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

    # Write Exclusion Report
    if excluded_rows:
        exclusion_file = '筛选排除原因.xlsx'
        df_excluded = pd.DataFrame(excluded_rows)
        df_excluded.to_excel(exclusion_file, index=False)
        print(f"Successfully generated {exclusion_file} with {len(excluded_rows)} excluded records.")
    else:
        print("No excluded employees found.")

if __name__ == "__main__":
    try:
        filter_bonus_data()
    except Exception as e:
        import traceback
        traceback.print_exc()
    finally:
        input("\nPress Enter to exit...")
