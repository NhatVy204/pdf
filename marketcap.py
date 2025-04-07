import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
# Đọc file Excel
df_1 = pd.read_excel(r"data\2024-Vietnam.xlsx")
xls = pd.ExcelFile(r"data\Vietnam_Marketcap.xlsx")
df_1 = df_1[(df_1["Ngành ICB - cấp 1"] == "Dịch vụ Tiêu dùng") & (df_1["Ngành ICB - cấp 2"] == "Bán lẻ")]
df_2 = pd.read_excel(xls, sheet_name='Sheet2')

# Xác định khoảng thời gian cần lọc
start_date = "2024-01-01"
end_date = "2025-01-01"

df_2["Ticker"] = df_2["Code"].str.extract(r"VT:([A-Z]+)\(")
unique_tickers = df_1["Mã"]
# Lọc các cổ phiếu thuộc ngành bán lẻ
df_retail = df_2[df_2["Ticker"].isin(unique_tickers)]

# Lọc dữ liệu theo khoảng thời gian
df_retail = df_retail[["Name", "Code"] + [col for col in df_retail.columns if start_date <= str(col) <= end_date]]

import pandas as pd

def get_market_value(date_target="2024-12-31", row_label="MOBILE WORLD INVESTMENT - MARKET VALUE", date_row_index=0):
    # Đường dẫn đến file Excel
    file_path = "data\\Vietnam_Marketcap.xlsx"
    
    # Đọc file Excel
    df = pd.read_excel(file_path, sheet_name="Sheet2", header=None)

    # Tìm cột ngày mục tiêu
    try:
        date_col_index = next(i for i, d in enumerate(df.iloc[date_row_index].astype(str)) if date_target in d)
    except StopIteration:
        raise ValueError(f"Không tìm thấy cột ngày {date_target} trong file Excel.")

    # Tìm dòng chứa nhãn cần tìm
    try:
        row_index = df[df.iloc[:, 0].astype(str).str.strip().str.contains(row_label, case=False, na=False)].index[0]
    except IndexError:
        raise ValueError(f"Không tìm thấy dòng '{row_label}' trong file Excel. Kiểm tra lại tên!")

    # Lấy giá trị tại ô tương ứng
    market_value = df.iloc[row_index, date_col_index]
    
    # Trả về giá trị đã chuyển sang ngàn VND
    return market_value / 1000

#market_value = get_market_value(date_target="2024-12-31", row_label="MOBILE WORLD INVESTMENT - MARKET VALUE")
#print(f"Vốn hóa thị trường tại ngày 31-12-2024 (ngàn VND): {market_value:,.3f}B")


#Hàm vẽ bubble chart
def plot_marketcap(df_retail, date_column_prefix="2024-12-31"):
    # Cấu hình chung cho font chữ
    plt.rcParams.update({
        "font.family": "serif",
        "font.size": 10,
        "axes.labelweight": "bold",
        "axes.titlesize": 15,
        "axes.titleweight": "bold"
    })

    # Xử lí NaN
    df_retail_cleaned = df_retail.fillna(0)
    df_plot = df_retail_cleaned.copy()

    # Chuyển tên cột sang chuỗi để tìm ngày 31/12/2024
    date_col = None
    for col in df_retail.columns:
        if str(col).startswith(date_column_prefix):
            date_col = col
            break

    if date_col is None:
        raise ValueError(f"Không tìm thấy cột ngày bắt đầu với {date_column_prefix} trong dữ liệu.")

    # Thêm cột 'Ticker' nếu chưa có
    df_plot["Ticker"] = df_plot["Code"].str.extract(r"VT:([A-Z]+)\(")

    # Lấy vốn hóa tại ngày 31/12/2024
    df_plot["marketcap_3112"] = df_plot[date_col]
    
    # Chia thành MWG và các mã khác
    df_mwg = df_plot[df_plot["Ticker"] == "MWG"]
    df_others = df_plot[df_plot["Ticker"] != "MWG"]

    # Vẽ biểu đồ
    plt.figure(figsize=(10, 6))

    # Các cổ phiếu khác
    plt.scatter(df_others["Ticker"], df_others["marketcap_3112"],
                s=df_others["marketcap_3112"] / 1e3,  # scale size
                c='skyblue', alpha=0.6, label='Khác')

    # MWG nổi bật
    plt.scatter(df_mwg["Ticker"], df_mwg["marketcap_3112"],
                s=df_mwg["marketcap_3112"] / 1e3,
                c='orange', alpha=0.9, label='MWG')

    # Tiêu đề và nhãn
    plt.title("Vốn hóa các cổ phiếu ngành Bán lẻ - 31/12/2024")
    plt.ylabel("Vốn hóa (VNĐ)")
    plt.xticks([])  # Không hiển thị nhãn trên trục X

    # Sắp xếp lại bố cục để không bị cắt xén
    plt.tight_layout()

    # Hiển thị đồ thị
    plt.show()

# Sử dụng hàm
#plot_marketcap(df_retail)
