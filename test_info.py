import requests
import ssl
import urllib3  
from urllib3.util.ssl_ import create_urllib3_context
from urllib3.poolmanager import PoolManager
from requests.adapters import HTTPAdapter
from bs4 import BeautifulSoup
from vnstock import *
import webbrowser
import pandas as pd

class CustomHttpAdapter (requests.adapters.HTTPAdapter):
    def __init__(self, ssl_context=None, **kwargs):
        self.ssl_context = ssl_context
        super().__init__(**kwargs)

    def init_poolmanager(self, connections, maxsize, block=False):
        self.poolmanager = urllib3.poolmanager.PoolManager(
            num_pools=connections, maxsize=maxsize,
            block=block, ssl_context=self.ssl_context)


def get_legacy_session():
    ctx = ssl.create_default_context(ssl.Purpose.SERVER_AUTH)
    ctx.options |= 0x4  # OP_LEGACY_SERVER_CONNECT
    session = requests.session()
    session.mount('https://', CustomHttpAdapter(ctx))
    return session

def get_mwg_intro(url):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/91.0.4472.124 Safari/537.36'}
    try:
        response = get_legacy_session().get(url, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            intro = soup.find('div', class_='intro_content').find('p')
            return intro.get_text(strip=True) if intro else "Không tìm thấy phần giới thiệu."
            
    except requests.exceptions.SSLError as e:
        return f"Lỗi SSL: {e}"
    except Exception as e:
        return f"Lỗi khác: {e}"


def get_mwg_info(label=None):
    vietstock_url = "https://finance.vietstock.vn/MWG-ctcp-dau-tu-the-gioi-di-dong.htm"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        response = requests.get(vietstock_url, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            profile_div = soup.find("div", id="profile-1").find("p")
            
            # Xử lý các thẻ HTML thừa như <b>, <a>, <span>
            for tag in profile_div.find_all(['b', 'a', 'span']):
                tag.unwrap()

            # Lấy các đoạn thông tin sau khi đã loại bỏ thẻ không cần thiết
            fields = str(profile_div).replace("</p>", "").split("<p>")[1:]

            # Dictionary để lưu thông tin
            info = {}
            labels = ["Địa chỉ", "Điện thoại", "Website"]

            # Duyệt qua các đoạn thông tin tìm kiếm theo từng label
            for text in fields:
                for key in labels:
                    if key in text:
                        value = text.split(":")[-1].strip()  # Lấy phần giá trị sau dấu ":"
                        if key == "Website":
                            info[key] = f"http:{value}"  # Thêm "http://" cho website
                        else:
                            info[key] = value

            # Trả về thông tin theo label cụ thể hoặc toàn bộ thông tin
            if label:
                return info.get(label, f"Không tìm thấy thông tin cho {label}")
            
            return info  # Trả về toàn bộ thông tin nếu không chỉ định label

    except Exception as e:
        print(f"Error: {e}")
        return None

url = "https://mwg.vn"  # Thay bằng URL thực tế

def get_company_info(ticker, column=None):
    data = Company(symbol=ticker)
    info = data.overview()
    if info.empty:
        return f"Không tìm thấy dữ liệu cho {ticker}"
    if column:
        if column in info.columns:
            return info[column].iloc[0]  # Lấy giá trị của cột
        else:
            return f"Cột '{column}' không tồn tại trong dữ liệu."
    
    return info  # Trả về toàn bộ DataFrame nếu không truyền column

ticker = 'MWG'  # Mã cổ phiếu

import pandas as pd
from vnstock import Vnstock

# Hàm lấy dữ liệu chứng khoán từ Vnstock
def get_stock_data(symbol, start_date="2024-01-01", end_date="2024-12-31"):
  
    stock = Vnstock().stock(symbol=symbol, source="VCI")
    df = stock.quote.history(start=start_date, end=end_date, interval="1D")
    
    if df is not None and not df.empty:
        df["time"] = pd.to_datetime(df["time"])  # Chuyển đổi sang kiểu datetime
        return df, stock
    return None, None

# Hàm tính toán phần trăm thay đổi giá cổ phiếu
def calculate_percentage_changes(df):
    last_close = df["close"].iloc[-1]  # Giá đóng cửa mới nhất

    def get_change(days):
        
        target_date = df["time"].max() - pd.Timedelta(days=days)

        # Tìm ngày gần nhất có dữ liệu
        past_df = df[df["time"] <= target_date].sort_values(by="time", ascending=False)

        if not past_df.empty:
            past_close = past_df["close"].iloc[0]  # Lấy giá đóng cửa gần nhất sau ngày cần tìm
            return round(((last_close - past_close) / past_close) * 100, 2)

        return None  # Nếu không có dữ liệu nào cả

    changes = {
        "1 day": get_change(1),
        "5 day": get_change(4),
        "3 months": get_change(91),
        "6 months": get_change(211),
        "Month to Date": get_change(29),
        "Year to Date": get_change(364)
    }

    return changes

def get_financial_sumary():
    report = """
    Theo báo cáo tài chính quý 4/2024, CTCP Đầu tư Thế giới Di động (mã MWG) ghi nhận tổng doanh thu 34.574 tỷ đồng, tăng 10% so với cùng kỳ năm trước. 
    Biên lợi nhuận gộp đạt 19%, giảm so với mức 19,7% của quý 4/2023. Doanh thu hoạt động tài chính trong kỳ đạt 636 tỷ đồng, tăng 5% so với cùng kỳ.
    """
    return report

def calculate_beta(stock_symbol='MWG', market_symbol='VNINDEX', start_date='2024-01-01', end_date='2024-12-31'):
    def get_stock_data(symbol):
        stock = Vnstock().stock(symbol, source="VCI")
        df = stock.quote.history(start=start_date, end=end_date, interval="1D")
        if df is not None and not df.empty:
            df["time"] = pd.to_datetime(df["time"])
            df.set_index("time", inplace=True)
            return df
        return None

    # Lấy dữ liệu
    stock_df = get_stock_data(stock_symbol)
    market_df = get_stock_data(market_symbol)

    if stock_df is None or market_df is None:
        return None  # hoặc trả về chuỗi thông báo nếu thích

    # Tính lợi suất hàng ngày
    stock_returns = stock_df['close'].pct_change().dropna()
    market_returns = market_df['close'].pct_change().dropna()

    # Kết hợp dữ liệu
    returns_df = pd.DataFrame({
        'stock_return': stock_returns,
        'market_return': market_returns
    }).dropna()

    # Tính Beta
    covariance = returns_df['stock_return'].cov(returns_df['market_return'])
    market_variance = returns_df['market_return'].var()
    beta = covariance / market_variance

    return beta

def get_close_price_on_date(symbol='MWG', date='2024-12-31'):
    stock = Vnstock().stock(symbol=symbol, source='VCI')
    data = stock.quote.history(start=date, end=date, interval='1D')
    
    if not data.empty:
        return data.iloc[0]['close']

def get_market_value(date_target="2024-12-31", row_label="MOBILE WORLD INVESTMENT - MARKET VALUE", date_row_index=0):
    # Đường dẫn đến file Excel
    file_path = "E:\\Goi1\\CK\\data\\Vietnam_Marketcap.xlsx"
    
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

#close_price = get_close_price_on_date()
#print(close_price)

#beta_mwg = calculate_beta('MWG')
#print(f"Beta của MWG: {beta_mwg:.2f}")
        
# Chạy chương trình
#df, stock = get_stock_data("MWG", start_date="2024-01-01", end_date="2024-12-31")

#if df is not None:
    #percentage_changes = calculate_percentage_changes(df)
    #print("Phần trăm thay đổi giá cổ phiếu MWG:")
    #for key, value in percentage_changes.items():
        #print(f"{key}: {value}%")
#else:
    #print("Không thể lấy dữ liệu cổ phiếu.")

# Gọi hàm 
#print(get_financial_sumary())
#print(get_mwg_intro(url))
#print(get_mwg_info("Website"))
#print(get_mwg_info("Điện thoại"))
#print("Mã chứng khoán:", get_company_info(ticker, "symbol"))
#print("Toàn bộ thông tin:\n", get_company_info(ticker))
#beta_mwg = calculate_beta("MWG", "VNINDEX", start_date="2024-01-01", end_date="2024-12-31")
#print(f"{beta_mwg:.2f}")