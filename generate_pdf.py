# Standard library imports
import textwrap
from datetime import datetime
import base64
from openai import OpenAI
import os
import dotenv

dotenv.load_dotenv()


# Third party imports
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.dates import DateFormatter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, black
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from vnstock import *

# Local imports
from financial_ratio import calc_financial_ratios
from test_info import get_company_info, get_mwg_info, get_mwg_intro, get_close_price_on_date, get_financial_sumary, calculate_percentage_changes
from marketcap import get_market_value

# Constants
CHART_PATH_6M = "chart_image/6month.png"
CHART_PATH_5Y = "chart_image/5year.png"
CHART_PATH_PIE = "chart_image/piechar.png"
CHART_PATH_MARKET = "chart_image/maketcap.png"
FONT_PATH_REGULAR = 'Roboto-Regular.ttf'
FONT_PATH_BOLD = 'Roboto-Bold.ttf'
PAGE_MARGIN = 10 * mm
SYMBOL = "MWG"
DATE_TARGET = "2024-12-31"
WIDTH, HEIGHT = A4

def setup_fonts():
    """Register custom fonts"""
    pdfmetrics.registerFont(TTFont('Roboto', FONT_PATH_REGULAR))
    pdfmetrics.registerFont(TTFont('Roboto-Bold', FONT_PATH_BOLD))

def analyze_chart(image_path):
    """
    Analyze a chart image using OpenRouter's AI model from a financial expert perspective.
    
    Args:
        image_path (str): Path to the chart image file
        
    Returns:
        str: ~200 word financial expert analysis of the chart's content and patterns
    """
    try:
        # Read and encode the image
        with open(image_path, "rb") as image_file:
            base64_image = base64.b64encode(image_file.read()).decode('utf-8')
        
        # Initialize OpenAI client with OpenRouter configuration
        client = OpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=os.getenv("OPENROUTER_API_KEY"),
        )
        
        # Create completion request with specific financial expert prompt
        completion = client.chat.completions.create(
            model="meta-llama/llama-4-maverick:free",
            messages=[
                {
                    "role": "system",
                    "content": "You are a seasoned financial analyst with expertise in interpreting financial charts and metrics. Provide professional, insightful analysis focusing on key trends, potential implications, and actionable insights."
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "As a financial expert, please analyze this chart in approximately 120 words, return in Vietnamese paragraph, no format. Include:\n"
                                  "1. Key trends and patterns\n"
                                  "2. Important financial metrics and their implications\n"
                                  "3. Notable market insights\n"
                                  "4. Potential impact on investment decisions\n"
                                  "Keep the analysis concise, professional, and focused on the most significant aspects."
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ]
        )
        
        # Return the analysis
        return completion.choices[0].message.content
        
    except Exception as e:
        return f"Error analyzing chart: {str(e)}"


def plot_stock_price_chart(symbol=SYMBOL, period="6m", start_date="2020-01-01", end_date="2024-12-31", save_path=None):
    """Generate and save stock price chart"""
    stock = Vnstock().stock(symbol, source="VCI")
    df = stock.quote.history(start=start_date, end=end_date, interval="1D")
    
    if df is None or df.empty:
        print(f"Không thể lấy dữ liệu cho mã {symbol}")
        return

    df["time"] = pd.to_datetime(df["time"])
    plt.figure(figsize=(5, 3))

    if period == "6m":
        df_filtered = df[(df["time"] >= "2024-07-01") & (df["time"] <= "2024-12-31")]
        plt.plot(df_filtered["time"], df_filtered["close"], label="Giá đóng cửa", color="blue")
        plt.xticks(pd.date_range(start="2024-07-01", end="2024-12-31", freq="MS"),
                  labels=pd.date_range(start="2024-07-01", end="2024-12-31", freq="MS").strftime('%m/%Y'))
        plt.title(f"{symbol} - 6 tháng", loc="left", fontsize=10)
    else:  # 5y
        plt.plot(df["time"], df["close"], label="Giá đóng cửa", color="blue")
        plt.gca().xaxis.set_major_formatter(DateFormatter("%Y"))
        plt.xticks(pd.date_range(start="2020-01-01", end="2024-12-31", freq="YS"),
                  labels=pd.date_range("2020-01-01", "2024-12-31", freq="YS").strftime("%Y"))
        plt.title(f"{symbol} - 5 năm", loc="left", fontsize=10)

    plt.gca().yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{x:.3f}"))
    plt.grid(True)
    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.close()

def draw_wrapped_text(c, text, x, y, width=100, font="Roboto", font_size=11, leading=14):
    """Draw text with automatic line wrapping"""
    c.setFont(font, font_size)
    text_object = c.beginText(x, y)
    text_object.setFont(font, font_size)
    text_object.setLeading(leading)
    
    wrapped_lines = textwrap.wrap(text, width=width)
    for line in wrapped_lines:
        text_object.textLine(line)
        
    c.drawText(text_object)
    return len(wrapped_lines) * leading

def draw_section_title(c, x, y, title, width):
    """Draw section title with styling"""
    c.setFont("Roboto-Bold", 14)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(x, y, title.upper())
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(x, y - 5, x + width, y - 5)

def draw_table_header(c, x, y, col_labels, title_col_width, data_col_width, row_height, table_width):
    """Draw table header row"""
    c.setFont("Roboto-Bold", 11)
    c.setFillColor(HexColor("#FFD700"))
    c.rect(x, y - row_height, table_width, row_height, stroke=0, fill=1)
    c.setFillColor(black)
    c.drawString(x + 4, y - row_height + 5, "Chỉ tiêu")
    
    for i, year in enumerate(col_labels):
        c.drawRightString(
            x + title_col_width + (i * data_col_width) + data_col_width - 4,
            y - row_height + 5,
            str(year)
        )
    return y - row_height

def draw_table_from_dict(c, data_dict, x, y, row_height=20, section_title=None):
    """Draw table from dictionary data"""
    table_width = WIDTH - 2 * PAGE_MARGIN
    col_labels = list(data_dict.values())[0].index.tolist()
    
    title_col_width = 0.25 * (WIDTH / mm) * mm
    data_col_width = (table_width - title_col_width) / len(col_labels)

    if y < 60:
        c.showPage()
        y = HEIGHT - 40

    if section_title:
        draw_section_title(c, x, y, section_title, table_width)
        y -= 20

    y = draw_table_header(c, x, y, col_labels, title_col_width, data_col_width, row_height, table_width)

    for idx, (label, values) in enumerate(data_dict.items()):
        if y < 60:
            c.showPage()
            y = HEIGHT - 40

        if idx % 2 == 1:
            c.setFillColor(HexColor("#F2F2F2"))
            c.rect(x, y - row_height, table_width, row_height, stroke=0, fill=1)

        c.setStrokeColor(HexColor("#DDDDDD"))
        c.rect(x, y - row_height, table_width, row_height, stroke=1, fill=0)

        c.setFillColor(black)
        c.setFont("Roboto", 10)
        c.drawString(x + 4, y - row_height + 5, label)
        
        for i, val in enumerate(values):
            val_str = f"{val:,.0f}" if isinstance(val, (int, float)) else str(val)
            c.drawRightString(
                x + title_col_width + (i * data_col_width) + data_col_width - 4,
                y - row_height + 5,
                val_str
            )
        y -= row_height

    return y - 20

def prepare_financial_data(financial_ratios):
    """Prepare financial data tables"""
    def create_table(data, years=range(2020, 2025)):
        df = pd.DataFrame(data).T
        df.columns = [str(year) for year in years]
        return {row_label: pd.Series(row.values, index=df.columns) for row_label, row in df.iterrows()}

    balance_sheet = create_table({
        "Tổng tài sản hiện có": financial_ratios["Total Current Assets"],
        "Bất động sản/Nhà xưởng/Thiết bị": financial_ratios["Property/Plant/Equipment"],
        "Tổng tài sản": financial_ratios["Total Assets"],
        "Nợ ngắn hạns": financial_ratios["Total Current Liabilities"],
        "Nợ dài hạn": financial_ratios["Total Long-Term Debt"],
        "Tổng nợ phải trả": financial_ratios["Total Liabilities"]
    })

    income_statement = create_table({
        "Doanh thu": financial_ratios["Revenue"],
        "Chi phí hoạt động": financial_ratios["Total Operating Expense"],
        "Thu nhập ròng trước thuế": financial_ratios["Net Income Before Taxes"],
        "Thu nhập ròng sau thuế": financial_ratios["Net Income After Taxes"],
        "Thu nhập ròng trước bất thường": financial_ratios["Net Income Before Extraordinary Items"]
    })

    profitability = create_table({
        "ROE, %": financial_ratios["ROE"],
        "ROA, %": financial_ratios["ROA"],
        "ROS, %": financial_ratios["ROS"],
        "Biên lợi nhuận hoạt động, %": financial_ratios["Income After Tax Margin"],
        "Doanh thu/Tổng tài sản, %": financial_ratios["Revenue/Total Assets"],
        "Nợ dài hạn/Vốn chủ sở hữu, %": financial_ratios["Long Term Debt/Equity"],
        "TTổng nợ /Vốn chủ sở hữu, %": financial_ratios["Total Debt/Equity"],
    })

    return balance_sheet, income_statement, profitability

def get_stock_details(symbol=SYMBOL):
    """Get stock details including percentage changes"""
    from test_info import get_stock_data, calculate_percentage_changes
    
    df, stock = get_stock_data(symbol, start_date="2024-01-01", end_date=DATE_TARGET)
    
    if df is None or df.empty:
        return None
        
    current_price = df.iloc[-1]['close']
    percentage_changes = calculate_percentage_changes(df)
    
    # Get volume averages
    five_day_volume = df.tail(5)['volume'].mean()
    
    # Get beta value from test_info
    from test_info import calculate_beta
    beta_value = calculate_beta(symbol)
    
    # Get shares outstanding (placeholder - implement actual data fetching)
    shares_outstanding = "6B"  # Example value
    
    return {
        'Giá đóng cửa': current_price,
        'Thay đổi 1 ngày': percentage_changes['1 day'],
        'Thay đổi 5 ngày': percentage_changes['5 day'],
        'Thay đổi 3 tháng': percentage_changes['3 months'],
        'Thay đổi 6 tháng ': percentage_changes['6 months'], 
        'Thay đổi trong tháng': percentage_changes['Month to Date'],
        'Thay đổi trong năm': percentage_changes['Year to Date'],
        'five_day_volume': five_day_volume,
        'Beta': beta_value,
        'Cổ phiếu lưu hành': shares_outstanding
    }

def draw_financial_summary(c, y_position):
    """Draw financial summary section"""
    c.showPage()  # Tạo trang mới
    y_position = HEIGHT - PAGE_MARGIN - 20
    
    # Tiêu đề phần
    c.setFont("Roboto-Bold", 14)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(PAGE_MARGIN, y_position, "TÌNH HÌNH TÀI CHÍNH")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, PAGE_MARGIN + 190 * mm, y_position - 4)

    # Nội dung summary
    y_position -= 20
    c.setFillColor(black)
    c.setFont("Roboto", 11)

    summary_text = get_financial_sumary()
    for line in summary_text.strip().split('\n'):
        draw_wrapped_text(c, line, PAGE_MARGIN, y_position, width=95, font="Roboto", font_size=11)
        y_position -= 30  # Increased margin

    return y_position - 40

def draw_share_details(c, details, y_position):
    """Draw share details and percentage change tables"""
    # Draw titles for both tables
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    
    # Share Detail title
    c.drawString(PAGE_MARGIN, y_position, "THÔNG TIN CHI TIẾT")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, PAGE_MARGIN + 90 * mm, y_position - 4)
    
    # Percentage Change title
    c.drawString(WIDTH/2, y_position, "PHẦN TRĂM THAY ĐỔI")
    c.line(WIDTH/2, y_position - 4, WIDTH/2 + 95 * mm, y_position - 4)
    
    y_position -= 15
    
    # Share Detail data
    left_data = [
    ("Giá đóng cửa", f"{details['Giá đóng cửa']:,.3f}"),
    ("Beta", f"{details['Beta']:,.3f}"), 
    ("Đơn vị tiền", "VND"),
    ("Cổ phiếu lưu hành", str(details['Cổ phiếu lưu hành']))
]
    
    # Draw Share Detail table
    table_width = WIDTH/2 - PAGE_MARGIN - 10
    row_height = 20

    # Draw rows
    y = y_position
    for i, (label, value) in enumerate(left_data):
        # Add alternating background
        if i % 2 == 1:
            c.setFillColor(HexColor("#F2F2F2"))
            c.rect(PAGE_MARGIN, y - row_height, table_width, row_height, stroke=0, fill=1)

        # Draw only horizontal borders
        c.setStrokeColor(HexColor("#DDDDDD"))
        c.line(PAGE_MARGIN, y - row_height, PAGE_MARGIN + table_width, y - row_height)

        c.setFillColor(black)
        c.setFont("Roboto", 10)
        # Draw label and value left-aligned
        c.drawString(PAGE_MARGIN + 4, y - row_height + 5, label)
        c.drawString(PAGE_MARGIN + table_width/2 + 4, y - row_height + 5, str(value))
        
        y -= row_height
    
    # Draw Percentage Change table
    table_data = {
    ' 1 ngày': details['Thay đổi 1 ngày'],
    ' 5 ngày': details['Thay đổi 5 ngày'],
    ' 3 tháng': details['Thay đổi 3 tháng'],
    ' 6 tháng ': details['Thay đổi 6 tháng '], 
    ' Đầu tháng - Hiện tại': details['Thay đổi trong tháng'],
    ' Đầu năm - Hiện tại': details['Thay đổi trong năm'],
}


    table_width = WIDTH/2 - PAGE_MARGIN - 10
    row_height = 20

    # Draw rows
    y = y_position
    for i, (label, value) in enumerate(table_data.items()):
        # Add alternating background
        if i % 2 == 1:
            c.setFillColor(HexColor("#F2F2F2"))
            c.rect(WIDTH/2, y - row_height, table_width, row_height, stroke=0, fill=1)

        # Draw only horizontal borders
        c.setStrokeColor(HexColor("#DDDDDD"))
        c.line(WIDTH/2, y - row_height, WIDTH/2 + table_width, y - row_height)

        c.setFillColor(black)
        c.setFont("Roboto", 10)
        # Draw label and value left-aligned
        c.drawString(WIDTH/2 + 4, y - row_height + 5, label)
        c.drawString(WIDTH/2 + table_width/2 + 4, y - row_height + 5, f"{value:.2f}%")
        
        y -= row_height
    
    y = draw_financial_summary(c, y - 10)
    return y

def draw_header(c, price):
    """Draw the header section of the PDF"""
    c.setFont("Roboto-Bold", 20)
    c.setFillColor(HexColor("#E6B800"))
    c.drawRightString(WIDTH - PAGE_MARGIN, HEIGHT - 40, "THẾ GIỚI DI ĐỘNG-MWG")
    
    c.setFont("Roboto-Bold", 18)
    c.setFillColor(black)
    c.drawRightString(WIDTH - PAGE_MARGIN, HEIGHT - 60, DATE_TARGET)
    
    c.setFont("Roboto", 14)
    c.drawRightString(WIDTH - PAGE_MARGIN, HEIGHT - 75, str(price))

def draw_company_info(c, y_position, market_value):
    """Draw company information sections"""
    # Left column
    left_x = PAGE_MARGIN
    # Draw section title
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(left_x, y_position, "THÔNG TIN CHUNG")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(left_x, y_position - 5, left_x + 250, y_position - 5)
    y_position -= 20

    company_info = {
        "Sàn giao dịch": get_company_info(SYMBOL, "exchange"),
        "Ngành": get_company_info(SYMBOL, "industry"),
        "Nhân viên": get_company_info(SYMBOL, "no_employees"),
        "Vốn hóa (VND)": f"{market_value:,.0f}B" if market_value else "N/A"
    }

    c.setFillColor(black)
    for key, value in company_info.items():
        c.setFont("Roboto-Bold", 11)
        c.drawString(left_x, y_position, f"{key}:")
        key_width = c.stringWidth(f"{key}:", "Roboto-Bold", 11)
        c.setFont("Roboto", 11)
        c.drawString(left_x + key_width + 2, y_position, str(value))
        y_position -= 15

    # Right column
    right_x = WIDTH / 2
    y_position_right = y_position + 80
    # Draw section title
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(right_x, y_position_right, "THÔNG TIN CÔNG TY")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(right_x, y_position_right - 5, right_x + 270, y_position_right - 5)
    y_position_right -= 20

    company_details = {
        "Địa chỉ": get_mwg_info("Địa chỉ"),
        "Điện thoại": get_mwg_info("Điện thoại"),
        "Website": get_mwg_info("Website")
    }

    for key, value in company_details.items():
        if key == "Địa chỉ" and value:
            c.setFont("Roboto-Bold", 11)
            c.setFillColor(black)
            c.drawString(right_x, y_position_right, f"{key}:")
            key_width = c.stringWidth(f"{key}:", "Roboto-Bold", 11)
            
            split_point = value.find("T.")
            if split_point != -1:
                part1 = value[:split_point]
                part2 = value[split_point:]
                c.setFont("Roboto", 11)
                c.setFillColor(black)  # Set text color to black for content
                c.drawString(right_x + key_width + 2, y_position_right, part1)
                y_position_right -= 14
                
                for line in textwrap.wrap(part2, width=40):
                    c.drawString(right_x, y_position_right, line)
                    y_position_right -= 14
            else:
                c.setFont("Roboto", 10)
                c.setFillColor(black)  # Set text color to black for content
                for i, line in enumerate(textwrap.wrap(value, width=40)):
                    if i == 0:
                        c.drawString(right_x + key_width + 2, y_position_right, line)
                    else:
                        y_position_right -= 14
                        c.drawString(right_x, y_position_right, line)
                y_position_right -= 15
        else:
            c.setFont("Roboto-Bold", 11)
            c.drawString(right_x, y_position_right, f"{key}:")
            key_width = c.stringWidth(f"{key.upper()}:", "Roboto-Bold", 11)
            c.setFont("Roboto", 11)
            c.setFillColor(black)  # Set text color to black for content
            c.drawString(right_x + key_width + 2, y_position_right, str(value or ""))
            y_position_right -= 15

    return min(y_position, y_position_right)

def draw_business_summary(c, y_position):
    """Draw business summary section"""
    # Draw section title
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(PAGE_MARGIN, y_position, "TÓM TẮT KINH DOANH")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 5, PAGE_MARGIN + 540, y_position - 5)
    y_position -= 30  # Increased margin above TÓM TẮT KINH DOANH
    
    c.setFillColor(black)
    c.setFont("Roboto", 11)  # Set text color to black for content
    intro = get_mwg_intro("https://mwg.vn") or "Không có thông tin tóm tắt."
    return y_position - draw_wrapped_text(c, intro, x=PAGE_MARGIN, y=y_position, width=95, font="Roboto", font_size=11) - 20

def draw_charts(c, y_position):
    """Draw stock price charts"""
    # Draw chart titles
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    
    c.drawString(PAGE_MARGIN, y_position, " 6 THÁNG")
    c.drawString(105 * mm, y_position, " 5 NĂM")
    
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, 100 * mm, y_position - 4)
    c.line(105 * mm, y_position - 4, 200 * mm, y_position - 4)

    # Draw charts
    y_position -= 10
    chart_width = 90 * mm
    chart_height = 60 * mm
    
    c.drawImage(ImageReader(CHART_PATH_6M), PAGE_MARGIN, y_position - chart_height, 
                width=chart_width, height=chart_height)
    c.drawImage(ImageReader(CHART_PATH_5Y), 105 * mm, y_position - chart_height,
                width=chart_width, height=chart_height)
                
    return y_position - (chart_height + 10)

def main():
    # Initialize
    setup_fonts()
    financial_ratios = calc_financial_ratios()    
    # Generate charts
    plot_stock_price_chart(period="6m", save_path=CHART_PATH_6M)
    plot_stock_price_chart(period="5y", save_path=CHART_PATH_5Y)

    # Prepare financial data
    balance_sheet, income_statement, profitability = prepare_financial_data(financial_ratios)
    
    # Create PDF
    c = canvas.Canvas("K224141709_Hồ Nguyễn Nhật Vy_MWG_1.pdf", pagesize=A4)
    # Draw content
    price = get_close_price_on_date(SYMBOL, DATE_TARGET) or "N/A"
    price = "{:,.3f}".format(price).replace(",", ".")
    draw_header(c, price)
    
    y_position = HEIGHT - 100
    
    # Get stock details
    stock_details = get_stock_details()
    
    market_value = get_market_value(date_target=DATE_TARGET, row_label="MOBILE WORLD INVESTMENT - MARKET VALUE")
    y_position = draw_company_info(c, y_position, market_value)
    y_position = draw_business_summary(c, y_position - 20)  # Added margin above title
    y_position = draw_charts(c, y_position - 20)
    y_position = draw_share_details(c, stock_details, y_position - 20)

    # Draw financial tables
    y_position = draw_table_from_dict(c, balance_sheet, PAGE_MARGIN, y_position, 
                                    section_title="Bảng Cân Đối Kế Toán")
                                    
    # Draw balance sheet chart
    chart_width = WIDTH - 2 * PAGE_MARGIN
    chart_height = 60 * mm
    c.drawImage(ImageReader("chart_image/taisan & no.png"), 
                PAGE_MARGIN, y_position - chart_height,
                width=chart_width, height=chart_height)
    y_position -= (chart_height + 40)
    
    # Add AI analysis section
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(PAGE_MARGIN, y_position, "AI PHÂN TÍCH")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, 200 * mm, y_position - 4)
    y_position -= 20

    analysis = analyze_chart("chart_image/taisan & no.png")
    c.setFillColor(black)  # Set text color to black
    c.setFont("Roboto", 11)
    y_position -= draw_wrapped_text(c, analysis, PAGE_MARGIN, y_position, width=95,font="Roboto", font_size=11)
    y_position -= 20  # Add some spacing

    y_position = draw_table_from_dict(c, income_statement, PAGE_MARGIN, y_position-100,
                                    section_title="Báo Cáo Kết Quả Kinh Doanh")
    
    y_position = draw_table_from_dict(c, profitability, PAGE_MARGIN, y_position-30,
                        section_title="Hiệu Suất Sinh Lời")
    
    # Draw ROA/ROE/ROS chart
    chart_width = WIDTH - 2 * PAGE_MARGIN
    chart_height = 60 * mm
    c.drawImage(ImageReader("chart_image/roa_roe_ros.png"), 
                PAGE_MARGIN, y_position - chart_height,
                width=chart_width, height=chart_height)
    y_position -= (chart_height + 30)
    
    # Add AI analysis for ROA/ROE/ROS
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(PAGE_MARGIN, y_position, "AI PHÂN TÍCH")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, 200 * mm, y_position - 4)
    y_position -= 20

    analysis = analyze_chart("chart_image/roa_roe_ros.png")
    c.setFillColor(black)  # Set text color to black
    c.setFont("Roboto", 11)
    y_position -= draw_wrapped_text(c, analysis, PAGE_MARGIN, y_position, width=95, font="Roboto", font_size=11)
    y_position -= 20

    # Start a new page for investor trading charts
    c.showPage()
    y_position = HEIGHT - 40

    # Draw investor trading charts
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    
    c.drawString(PAGE_MARGIN, y_position, "KHỚP LỆNH NHÀ ĐẦU TƯ")
    c.drawString(105 * mm, y_position, "THỎA THUẬN NHÀ ĐẦU TƯ")
    
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, 100 * mm, y_position - 4)
    c.line(105 * mm, y_position - 4, 200 * mm, y_position - 4)

    # Draw trading charts side by side
    y_position -= 10
    chart_width = 90 * mm
    chart_height = 60 * mm
    
    c.drawImage(ImageReader("chart_image/Khop_lenhNĐT.png"), 
                PAGE_MARGIN, y_position - chart_height,
                width=chart_width, height=chart_height)
    c.drawImage(ImageReader("chart_image/Thoa_thuanNĐT.png"), 
                105 * mm, y_position - chart_height,
                width=chart_width, height=chart_height)
    
    y_position -= (chart_height + 20)
    
    # Add AI analysis for investor trading
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(PAGE_MARGIN, y_position, "AI PHÂN TÍCH")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, 200 * mm, y_position - 4)
    y_position -= 20

    analysis = analyze_chart("chart_image/Khop_lenhNĐT.png")
    c.setFillColor(black)  # Set text color to black
    c.setFont("Roboto", 11)
    y_position -= draw_wrapped_text(c, analysis, PAGE_MARGIN, y_position, width=95, font="Roboto", font_size=11)
    y_position -= 20

    # Draw pie chart
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    
    c.drawString(PAGE_MARGIN, y_position, "PHÂN TÍCH CƠ CẤU")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, 200 * mm, y_position - 4)

    # Draw pie chart
    y_position -= 10
    chart_width = WIDTH - 2 * PAGE_MARGIN
    chart_height = 80 * mm
    
    c.drawImage(ImageReader(CHART_PATH_PIE), 
                PAGE_MARGIN, y_position - chart_height,
                width=chart_width, height=chart_height)
    
    y_position -= (chart_height + 10)
    
    # Add AI analysis for pie chart
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(PAGE_MARGIN, y_position, "AI PHÂN TÍCH")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, 200 * mm, y_position - 4,)
    y_position -= 20

    analysis = analyze_chart(CHART_PATH_PIE)
    c.setFillColor(black)  # Set text color to black
    c.setFont("Roboto", 11)
    y_position -= draw_wrapped_text(c, analysis, PAGE_MARGIN, y_position, width=95, font="Roboto", font_size=11)
    y_position -= 20

    # Start new page for market cap section
    c.showPage()
    y_position = HEIGHT - 40

    # Draw market cap chart
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(PAGE_MARGIN, y_position, "GIÁ TRỊ THỊ TRƯỜNG")
    c.setStrokeColor(HexColor("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, 200 * mm, y_position - 4)

    
    # Draw market cap chart
    y_position -= 10
    chart_width = WIDTH - 2 * PAGE_MARGIN
    chart_height = 80 * mm
    
    c.drawImage(ImageReader(CHART_PATH_MARKET), 
                PAGE_MARGIN, y_position - chart_height,
                width=chart_width, height=chart_height)
    
    y_position -= (chart_height + 20)
    
    # Add AI analysis for market cap chart
    c.setFont("Roboto-Bold", 12)
    c.setFillColor(HexColor("#E6B800"))
    c.drawString(PAGE_MARGIN, y_position, "AI PHÂN TÍCH")
    c.setStrokeColor(("#E6B800"))
    c.setLineWidth(1.2)
    c.line(PAGE_MARGIN, y_position - 4, 200* mm, y_position - 4)
    y_position -= 20

    analysis = analyze_chart(CHART_PATH_MARKET)
    c.setFillColor(black)  # Set text color to black
    c.setFont("Roboto", 11)
    y_position -= draw_wrapped_text(c, analysis, PAGE_MARGIN +20, y_position, width=95, font="Roboto", font_size=11)
    
    c.save()

if __name__ == "__main__":
    main()
