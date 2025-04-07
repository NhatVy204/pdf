import pandas as pd
import numpy as np

def convert_units(df, factor, start_col):
    start_idx = df.columns.get_loc(start_col) + 1
    numeric_cols = df.columns[start_idx:]
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce') / factor
    return df

def clean_columns(df, year):
    df.columns = df.columns.str.replace(f"Năm: {year}", "", regex=True)
    df.columns = df.columns.str.replace(r"Đơn vị: (Tỷ|Triệu) VND", "", regex=True)
    df.columns = df.columns.str.replace(r"\bHợp nhất\b|\bQuý: Hàng năm\b", "", regex=True)
    df.columns = df.columns.str.strip()
    df.drop(columns=[col for col in df.columns if "TM" in col], inplace=True)
    return df

def standardize_columns(df):
    df = df.copy()
    df.columns = df.columns.str.strip().str.replace("\n", " ").str.upper()
    return df

def load_all_data(file_paths, start_column, factor=1e9):
    dfs = []
    for i, file_path in enumerate(file_paths):
        df = pd.read_excel(file_path, engine="openpyxl")
        df = clean_columns(df, 2020 + i)
        if i < 3:  # Chỉ chuyển đơn vị từ 2020 đến 2022
            df = convert_units(df, factor, start_column)
        dfs.append(df)
    return dfs

def merge_df(dfs, stock_code):
    years = range(2020, 2025)
    dfs = [standardize_columns(df) for df in dfs]
    data = []

    for df, year in zip(dfs, years):
        if 'MÃ' not in df.columns:
            print(f"LỖI: Cột 'MÃ' không tồn tại trong file năm {year}")
            continue
        stock_data = df[df['MÃ'] == stock_code]
        if not stock_data.empty:
            data.append(stock_data)
    
    return pd.concat(data, ignore_index=True) if data else pd.DataFrame()

def transpose_data(df):
    df = df.T
    df.columns = [f"{year}" for year in range(2020, 2025)]
    df.reset_index(inplace=True)
    df.rename(columns={"index": "Chỉ tiêu"}, inplace=True)
    return df.fillna(0)

def calculate_financial_ratios(transposed_df, labels):
    def get_values(transposed_df, label):
        row = transposed_df[transposed_df["Chỉ tiêu"] == label]
        return row.iloc[:, 1:].values.flatten() if not row.empty else np.zeros(len(transposed_df.columns[1:]))

    def sum_labels(label_list):
        return sum((get_values(transposed_df, label) for label in label_list), np.zeros(len(transposed_df.columns[1:])))

    total_current_assets = sum_labels(labels["total_current_assets"])
    ppe = sum_labels(labels["ppe"])
    total_assets = sum_labels(labels["total_assets"])
    total_current_liabilities = get_values(transposed_df, labels["total_current_liabilities"][0])
    total_long_term_debt = get_values(transposed_df, labels["total_long_term_debt"][0])
    total_liabilities = get_values(transposed_df, labels["total_liabilities"][0])

    net_income = get_values(transposed_df, labels["net_income"][0])
    interest_expense = get_values(transposed_df, labels["interest_expense"][0])
    taxes = get_values(transposed_df, labels["taxes"][0])
    depreciation_amortization = get_values(transposed_df, labels["depreciation_amortization"][0])
    ebitda = net_income + interest_expense + taxes + depreciation_amortization

    operating_profit = get_values(transposed_df, labels["operating_profit"][0])
    other_profit = get_values(transposed_df, labels["other_profit"][0])
    jv_profit = get_values(transposed_df, labels["jv_profit"][0])
    net_income_before_taxes = operating_profit + other_profit + jv_profit

    other_income = get_values(transposed_df, labels["other_income"][0])
    net_income_before_extraordinary_items = net_income + other_income

    revenue = get_values(transposed_df, labels["revenue"][0])
    gross_profit = get_values(transposed_df, labels["gross_profit"][0])
    financial_expense = get_values(transposed_df, labels["financial_expense"][0])
    selling_expense = get_values(transposed_df, labels["selling_expense"][0])
    admin_expense = get_values(transposed_df, labels["admin_expense"][0])
    total_operating_expense = revenue - gross_profit + financial_expense + selling_expense + admin_expense

    total_equity = get_values(transposed_df, labels["total_equity"][0])
    total_debt = get_values(transposed_df, labels["total_debt"][0])

    roe = np.divide(net_income, total_equity, out=np.zeros_like(net_income), where=total_equity != 0)
    roa = np.divide(net_income, total_assets, out=np.zeros_like(net_income), where=total_assets != 0)
    income_after_tax_margin = np.divide(net_income, revenue, out=np.zeros_like(net_income), where=revenue != 0)
    revenue_to_total_assets = np.divide(revenue, total_assets, out=np.zeros_like(revenue), where=total_assets != 0)
    long_term_debt_to_equity = np.divide(total_long_term_debt, total_equity, out=np.zeros_like(total_long_term_debt), where=total_equity != 0)
    total_debt_to_equity = np.divide(total_debt, total_equity, out=np.zeros_like(total_debt), where=total_equity != 0)
    ros = np.divide(net_income, revenue, out=np.zeros_like(net_income), where=revenue != 0)

    years = transposed_df.columns[1:]
    financial_ratios = pd.DataFrame({
        "Năm": years,
        "Total Current Assets": [f"{value:,.2f}" for value in total_current_assets],
        "Property/Plant/Equipment": [f"{value:,.2f}" for value in ppe],
        "Total Assets": [f"{value:,.2f}" for value in total_assets],
        "Total Current Liabilities": [f"{value:,.2f}" for value in total_current_liabilities],
        "Total Long-Term Debt": [f"{value:,.2f}" for value in total_long_term_debt],
        "Total Liabilities": [f"{value:,.2f}" for value in total_liabilities],
        "EBITDA": [f"{value:,.2f}" for value in ebitda],
        "Net Income Before Taxes": [f"{value:,.2f}" for value in net_income_before_taxes],
        "Net Income Before Extraordinary Items": [f"{value:,.2f}" for value in net_income_before_extraordinary_items],
        "Revenue": [f"{value:,.2f}" for value in revenue],
        "Total Operating Expense": [f"{value:,.2f}" for value in total_operating_expense],
        "Net Income After Taxes": [f"{value:,.2f}" for value in net_income],
        "ROE": [f"{value * 100:,.2f}" for value in roe],
        "ROA": [f"{value * 100:,.2f}" for value in roa],
        "ROS": [f"{value * 100:,.2f}" for value in ros],
        "Income After Tax Margin": [f"{value:,.2f}" for value in income_after_tax_margin],
        "Revenue/Total Assets": [f"{value * 100:,.2f}" for value in revenue_to_total_assets],
        "Long Term Debt/Equity": [f"{value * 100:,.2f}" for value in long_term_debt_to_equity],
        "Total Debt/Equity": [f"{value * 100:,.2f}" for value in total_debt_to_equity],
    })

    return financial_ratios

def display_financial_data_table(data, table_name):
    # Tạo DataFrame từ dữ liệu
    financial_data_df = pd.DataFrame(data)

    # Chuyển vị (transpose) bảng để các năm là cột và các chỉ tiêu là hàng
    financial_data_df = financial_data_df.T

    # Đặt tên cột là năm
    financial_data_df.columns = [f"{year}" for year in range(2020, 2025)]
    
    return financial_data_df.to_string()
    
    
# ===== Thực thi =====
    
# Các chỉ tiêu cần thiết

labels = {
    "total_current_assets": [
        "CĐKT. TIỀN VÀ TƯƠNG ĐƯƠNG TIỀN",
        "CĐKT. ĐẦU TƯ TÀI CHÍNH NGẮN HẠN",
        "CĐKT. CÁC KHOẢN PHẢI THU NGẮN HẠN",
        "CĐKT. HÀNG TỒN KHO, RÒNG",
        "CĐKT. TÀI SẢN NGẮN HẠN KHÁC"
    ],
    "ppe": [
        "CĐKT. GTCL TSCĐ HỮU HÌNH",
        "CĐKT. GTCL TÀI SẢN THUÊ TÀI CHÍNH",
        "CĐKT. GTCL TÀI SẢN CỐ ĐỊNH VÔ HÌNH",
        "CĐKT. XÂY DỰNG CƠ BẢN DỞ DANG (TRƯỚC 2015)"
    ],
    "total_assets": [
        "CĐKT. TÀI SẢN NGẮN HẠN",
        "CĐKT. TÀI SẢN DÀI HẠN"
    ],
    "total_current_liabilities": ["CĐKT. NỢ NGẮN HẠN"],
    "total_long_term_debt": ["CĐKT. NỢ DÀI HẠN"],
    "total_liabilities": ["CĐKT. NỢ PHẢI TRẢ"],
    "net_income": ["KQKD. LỢI NHUẬN SAU THUẾ THU NHẬP DOANH NGHIỆP"],
    "interest_expense": ["KQKD. CHI PHÍ LÃI VAY"],
    "taxes": ["KQKD. CHI PHÍ THUẾ TNDN HIỆN HÀNH"],
    "depreciation_amortization": ["KQKD. KHẤU HAO TÀI SẢN CỐ ĐỊNH"],
    "revenue": ["KQKD. DOANH THU THUẦN"],
    "gross_profit": ["KQKD. LỢI NHUẬN GỘP VỀ BÁN HÀNG VÀ CUNG CẤP DỊCH VỤ"],
    "financial_expense": ["KQKD. CHI PHÍ TÀI CHÍNH"],
    "selling_expense": ["KQKD. CHI PHÍ BÁN HÀNG"],
    "admin_expense": ["KQKD. CHI PHÍ QUẢN LÝ DOANH NGHIỆP"],
    "total_equity": ["CĐKT. VỐN CHỦ SỞ HỮU"],
    "total_debt": ["CĐKT. NỢ PHẢI TRẢ"],
    "operating_profit": ["KQKD. LỢI NHUẬN THUẦN TỪ HOẠT ĐỘNG KINH DOANH"],
    "other_profit": ["KQKD. LỢI NHUẬN KHÁC"],
    "jv_profit": ["KQKD. LÃI/ LỖ TỪ CÔNG TY LIÊN DOANH (TRƯỚC 2015)"],
    "other_income": ["KQKD. LỢI NHUẬN KHÁC"]
}

file_paths = [
    r"data\2020-Vietnam.xlsx",
    r"data\2021-Vietnam.xlsx",
    r"data\2022-Vietnam.xlsx",
    r"data\2023-Vietnam.xlsx",
    r"data\2024-Vietnam.xlsx",
]



def calc_financial_ratios():
    start_column = "Trạng thái kiểm toán"
    start_column_clean = start_column.replace("Hợp nhất", "").replace("Hàng năm", "").strip()
    stock_code = "MWG"
    dfs = load_all_data(file_paths, start_column_clean)
    merged_df = merge_df(dfs, stock_code)
    merged_df = merged_df.loc[:, ~merged_df.columns.str.contains("CURRENT RATIO", case=False)]
    transposed_df = transpose_data(merged_df)
    #transposed_df

    financial_ratios = calculate_financial_ratios(transposed_df, labels)
    return financial_ratios

financial_ratios = calc_financial_ratios()

# Alternatively, output to CSV
financial_ratios.to_csv('financial_ratios.csv', index=False)
# Dữ liệu bảng tài chính
balance_sheet_data = {
    "Total Current Assets": financial_ratios["Total Current Assets"],
    "Property/Plant/Equipment": financial_ratios["Property/Plant/Equipment"],
    "Total Assets": financial_ratios["Total Assets"],
    "Total Current Liabilities": financial_ratios["Total Current Liabilities"],
    "Total Long-Term Debt": financial_ratios["Total Long-Term Debt"],
    "Total Liabilities": financial_ratios["Total Liabilities"]
}

income_statement_data = {
    "Revenue": financial_ratios["Revenue"],
    "Total Operating Expense": financial_ratios["Total Operating Expense"],
    "Net Income Before Taxes": financial_ratios["Net Income Before Taxes"],
    "Net Income After Taxes": financial_ratios["Net Income After Taxes"],
    "Net Income Before Extraordinary Items": financial_ratios["Net Income Before Extraordinary Items"]
}

profitability_analysis_data = {
    "ROE, %": financial_ratios["ROE"],
    "ROA, %": financial_ratios["ROA"],
    "Income After Tax Margin, %": financial_ratios["Income After Tax Margin"],
    "Revenue/Total Assets, %": financial_ratios["Revenue/Total Assets"],
    "Long Term Debt/Equity, %": financial_ratios["Long Term Debt/Equity"],
    "Total Debt/Equity, %": financial_ratios["Total Debt/Equity"],
    "ROS, %": financial_ratios["ROS"]
}

#Hiển thị bảng cho mỗi loại dữ liệu
#display_financial_data_table(balance_sheet_data , "BẢNG CÂN ĐỐI KẾ TOÁN")
#display_financial_data_table(income_statement_data, "BÁO CÁO KẾT QUẢ KINH DOANH")
#display_financial_data_table(profitability_analysis_data, "PHÂN TÍCH HIỆU SUẤT SINH LỜI")