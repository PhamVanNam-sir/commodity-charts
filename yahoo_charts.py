import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from bokeh.plotting import figure
from bokeh.models import HoverTool
from bokeh.embed import file_html
from bokeh.resources import INLINE
from openpyxl import Workbook
import os
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as OpenpyxlImage


# Set style cho matplotlib
plt.style.use('seaborn-v0_8-darkgrid')

# Mapping tên commodity
COMMODITY_NAMES = {
    'HRC=F': 'Hot Rolled Coil',
    'CL=F': 'Crude Oil (WTI)',
    'BZ=F': 'Brent Crude',
    'NG=F': 'Natural Gas',
    'RB=F': 'Gasoline',
    'HO=F': 'Heating Oil',
    'EH=F': 'Ethanol',
    'GC=F': 'Gold',
    'SI=F': 'Silver',
    'HG=F': 'Copper',
    'PL=F': 'Platinum',
    'PA=F': 'Palladium',
    'ALI=F': 'Aluminum',
    'DX=F': 'Dollar Index',
}

def calculate_returns(prices):
    """Tính các loại returns"""
    df = pd.DataFrame({'Close': prices})
    
    # Daily return
    df['Daily'] = df['Close'].pct_change() * 100
    
    # Weekly return (5 trading days)
    df['Weekly'] = df['Close'].pct_change(periods=5) * 100
    
    # Monthly return (21 trading days)
    df['Monthly'] = df['Close'].pct_change(periods=21) * 100
    
    # YoY return (252 trading days)
    df['YoY'] = df['Close'].pct_change(periods=252) * 100
    
    # YTD return
    year_start = df.index.to_series().apply(lambda x: pd.Timestamp(year=x.year, month=1, day=1))
    year_start_prices = df.groupby(year_start)['Close'].transform('first')
    df['YTD'] = ((df['Close'] - year_start_prices) / year_start_prices * 100)
    
    return df

def create_bokeh_chart(commodity_data, commodity_name, output_html):
    """Tạo biểu đồ interactive đẹp với Bokeh"""
    from bokeh.models import ColumnDataSource, CrosshairTool, Range1d
    from bokeh.models.formatters import DatetimeTickFormatter
    
    # Prepare data
    dates = commodity_data.index.to_pydatetime()
    prices = commodity_data['Close'].values
    daily_pct = commodity_data['Daily'].values
    
    # Format dates for display
    date_strings = [d.strftime('%Y-%m-%d') for d in dates]
    
    # Create data source
    source = ColumnDataSource(data={
        'x': dates,
        'y': prices,
        'date_str': date_strings,
        'daily_pct': daily_pct,
        'daily_pct_str': [f"{x:+.2f}%" if pd.notna(x) else "N/A" for x in daily_pct]
    })
    
    # Calculate price range with 5% padding
    price_min = prices.min()
    price_max = prices.max()
    price_range = price_max - price_min
    y_start = price_min - price_range * 0.05
    y_end = price_max + price_range * 0.05
    
    # Create figure với styling đẹp
    p = figure(
        title=f"{commodity_name}",
        x_axis_type='datetime',
        width=1400,
        height=700,
        tools="pan,wheel_zoom,box_zoom,reset,save,crosshair",
        toolbar_location="right",
        sizing_mode='scale_width',
        y_range=Range1d(y_start, y_end)
    )
    
    # Vẽ line chính - KHÔNG CÓ CIRCLE
    line = p.line('x', 'y', source=source, line_width=2.5, color='#2962FF', alpha=0.9)
    
    # Thêm fill area dưới line
    p.varea(x='x', y1=y_start, y2='y', source=source, alpha=0.1, color='#2962FF')
    
    # Hover tool với tooltip đẹp như TradingView
    hover = HoverTool(
        tooltips="""
        <div style="background-color: #1E222D; padding: 12px; border-radius: 4px; border: 1px solid #363A45; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto;">
            <div style="color: #B2B5BE; font-size: 11px; margin-bottom: 6px;">@date_str</div>
            <div style="display: flex; justify-content: space-between; gap: 20px;">
                <div>
                    <div style="color: #787B86; font-size: 11px;">Price</div>
                    <div style="color: #D1D4DC; font-size: 14px; font-weight: 600;">$@y{0,0.00}</div>
                </div>
                <div>
                    <div style="color: #787B86; font-size: 11px;">Daily Change</div>
                    <div style="color: @daily_pct_color; font-size: 14px; font-weight: 600;">@daily_pct_str</div>
                </div>
            </div>
        </div>
        """,
        formatters={'@x': 'datetime'},
        renderers=[line],
        mode='vline',
        line_policy='nearest'
    )
    
    # Add color for daily_pct in tooltip
    source.data['daily_pct_color'] = ['#26A69A' if pd.notna(x) and x >= 0 else '#EF5350' if pd.notna(x) else '#787B86' for x in daily_pct]
    
    p.add_tools(hover)
    
    # Crosshair styling
    crosshair = p.select_one(CrosshairTool)
    crosshair.line_color = '#787B86'
    crosshair.line_alpha = 0.6
    
    # Title styling
    p.title.text_font_size = '18pt'
    p.title.text_color = '#D1D4DC'
    p.title.text_font = 'Helvetica Neue, Arial'
    p.title.align = 'left'
    
    # Axes styling giống TradingView
    p.xaxis.axis_label_text_font_size = '0pt'  # Hide label
    p.yaxis.axis_label = 'Price ($)'
    p.yaxis.axis_label_text_font_size = '11pt'
    p.yaxis.axis_label_text_color = '#787B86'
    p.yaxis.axis_label_standoff = 10
    
    # Format datetime axis
    p.xaxis.formatter = DatetimeTickFormatter(
        hours='%H:%M',
        days='%d %b',
        months='%b %Y',
        years='%Y'
    )
    
    # Grid styling
    p.xgrid.grid_line_color = '#363A45'
    p.xgrid.grid_line_alpha = 0.5
    p.xgrid.grid_line_dash = 'dotted'
    p.ygrid.grid_line_color = '#363A45'
    p.ygrid.grid_line_alpha = 0.5
    p.ygrid.grid_line_dash = 'dotted'
    
    # Background colors - dark theme như TradingView
    p.background_fill_color = '#1E222D'
    p.border_fill_color = '#1E222D'
    p.outline_line_color = '#363A45'
    
    # Axis colors
    p.xaxis.axis_line_color = '#363A45'
    p.yaxis.axis_line_color = '#363A45'
    p.xaxis.major_tick_line_color = '#363A45'
    p.yaxis.major_tick_line_color = '#363A45'
    p.xaxis.minor_tick_line_color = None
    p.yaxis.minor_tick_line_color = None
    p.xaxis.major_label_text_color = '#787B86'
    p.yaxis.major_label_text_color = '#787B86'
    p.xaxis.major_label_text_font_size = '11pt'
    p.yaxis.major_label_text_font_size = '11pt'
    
    # Toolbar styling
    p.toolbar.logo = None  # Remove Bokeh logo
    
    # Save to HTML as a 100% self-contained file
    print(f"    Đang tạo file HTML tự chứa (self-contained) cho: {commodity_name}")
    
    # Dùng file_html và INLINE để nhúng toàn bộ JS/CSS vào file
    html_content = file_html(p, resources=INLINE, title=commodity_name)
    
    # Tự tay ghi nội dung đã nhúng ra file
    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"    Đã tạo file HTML tự chứa thành công: {output_html}")

    return output_html

def create_commodity_charts(df, 
                            output_file='commodity_charts.xlsx', 
                            period_years=1, 
                            
                            # --- CÔNG TẮC ĐIỀU KHIỂN ---
                            upload_mode=False,
                            
                            # --- CÁC THAM SỐ CHO CHẾ ĐỘ ---
                            local_html_folder='charts_html', # (Dùng khi upload_mode=False)
                            github_repo_local_path=None,   # (Dùng khi upload_mode=True)
                            github_pages_url=None          # (Dùng khi upload_mode=True)
                           ):
    """
    Vẽ biểu đồ giá cho từng commodity và xuất vào Excel.
    
    Nâng cấp (Hỗ trợ 2 chế độ):
    - Nếu upload_mode=False:
      - Sẽ lưu HTML vào 'local_html_folder'.
      - Sẽ tạo hyperlink trong Excel trỏ đến file local (file:///...)
      
    - Nếu upload_mode=True:
      - Sẽ lưu HTML vào 'github_repo_local_path'.
      - Sẽ tạo hyperlink trỏ đến link GitHub Pages (https://...)
    """
    
    # Đảm bảo date là index
    if 'date' in df.columns:
        df = df.set_index('date')
    
    commodities = df['name'].unique()
    cutoff_date = df.index.max() - pd.DateOffset(years=period_years)
    
    wb = Workbook()
    wb.remove(wb.active)
    
    colors = ['#E74C3C', '#3498DB', '#2ECC71', '#F39C12', '#9B59B6', 
              '#1ABC9C', '#E67E22', '#34495E', '#16A085', '#D35400']
    
    # Kiểm tra cấu hình dựa trên chế độ
    if upload_mode:
        if not (github_repo_local_path and github_pages_url):
            raise ValueError("LỖI: 'upload_mode=True' nhưng thiếu 'github_repo_local_path' hoặc 'github_pages_url'.")
        if not github_pages_url.endswith('/'):
            github_pages_url += '/'
        print("--- Đang chạy ở chế độ UPLOAD ---")
    else:
        os.makedirs(local_html_folder, exist_ok=True)
        print("--- Đang chạy ở chế độ LOCAL ---")

    for idx, commodity_code in enumerate(commodities):
        commodity_name = COMMODITY_NAMES.get(commodity_code, commodity_code)
        full_name = f"{commodity_name} ({commodity_code})"
        
        print(f"Đang xử lý {full_name}...")
        
        commodity_data_full = df[df['name'] == commodity_code].copy().sort_index()
        commodity_data_full = calculate_returns(commodity_data_full['Close'])
        commodity_data = commodity_data_full[commodity_data_full.index >= cutoff_date].copy()
        
        # === 1. TẠO MATPLOTLIB CHART ===
        # (Không thay đổi - giữ nguyên code cũ của bạn)
        fig, ax = plt.subplots(figsize=(14, 7), dpi=100)
        dates = commodity_data.index
        prices = commodity_data['Close'].values
        color = colors[idx % len(colors)]
        ax.plot(dates, prices, color=color, linewidth=2.5, label=full_name, 
                marker='o', markersize=3, markevery=max(1, len(dates)//50))
        ax.fill_between(dates, prices, alpha=0.2, color=color)
        price_min = prices.min(); price_max = prices.max()
        price_range = price_max - price_min
        y_min = price_min - price_range * 0.1; y_max = price_max + price_range * 0.1
        ax.set_ylim(y_min, y_max)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        plt.xticks(rotation=45, ha='right')
        ax.set_title(f'{full_name} - Last {period_years} Year(s)', 
                    fontsize=20, fontweight='bold', pad=20, color='#2C3E50')
        ax.set_xlabel('Date', fontsize=14, fontweight='bold', color='#34495E')
        ax.set_ylabel('Close Price ($)', fontsize=14, fontweight='bold', color='#34495E')
        ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.7)
        ax.set_facecolor('#FAFAFA'); fig.patch.set_facecolor('white')
        ax.legend(loc='upper left', framealpha=0.9, fontsize=12)
        plt.tight_layout()
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight', 
                   facecolor='white', edgecolor='none')
        img_buffer.seek(0)
        plt.close(fig)
        
        # === 2. TẠO BOKEH INTERACTIVE CHART (VÀ LINK) ===
        
        html_filename = f"{commodity_code.replace('=', '_')}.html"
        
        html_save_path = ""
        excel_hyperlink = ""
        excel_link_text = ""

        if upload_mode:
            # --- CHẾ ĐỘ UPLOAD: Lưu vào repo, link tới GitHub Pages ---
            html_save_path = os.path.join(github_repo_local_path, html_filename)
            excel_hyperlink = github_pages_url + html_filename
            excel_link_text = "Click to open (GitHub Page)"
        else:
            # --- CHẾ ĐỘ LOCAL: Lưu vào 'charts_html', link tới file local ---
            html_save_path = os.path.join(local_html_folder, html_filename)
            excel_hyperlink = os.path.abspath(html_save_path)
            excel_link_text = "Click to open (Local File)"
            
        # Tạo file HTML tự chứa (self-contained) tại đường dẫn đã chọn
        create_bokeh_chart(commodity_data, full_name, html_save_path)
        
        # === 3. TẠO EXCEL SHEET ===
        sheet_name = commodity_code.replace('=', '').replace('/', '-')[:31]
        ws = wb.create_sheet(title=sheet_name)
        
        ws['A1'] = full_name
        ws['A1'].font = Font(bold=True, size=16, color='2C3E50')
        ws.merge_cells('A1:G1'); ws['A1'].alignment = Alignment(horizontal='center')
        
        # Link to interactive chart (SỬ DỤNG LINK ĐÃ XỬ LÝ)
        ws['A2'] = 'Interactive Chart:'
        ws['B2'].hyperlink = excel_hyperlink
        ws['B2'].value = excel_link_text
        ws['B2'].font = Font(color='0563C1', underline='single')
        ws['B2'].style = 'Hyperlink'
        
        # (Phần còn lại của code tạo Excel giữ nguyên...)
        ws['A3'] = f'Period: Last {period_years} year(s) | Min: ${price_min:,.2f} | Max: ${price_max:,.2f} | Avg: ${prices.mean():,.2f}'
        ws['A3'].font = Font(size=10, color='7F8C8D')
        ws.merge_cells('A3:G3')
        
        headers = ['Date', 'Close', 'Daily %', 'Weekly %', 'Monthly %', 'YoY %', 'YTD %']
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=5, column=col_idx, value=header)
            cell.fill = header_fill; cell.font = header_font; cell.alignment = Alignment(horizontal='center')
        
        recent_data = commodity_data.sort_index(ascending=False).head(10)
        
        for row_idx, (date, row) in enumerate(recent_data.iterrows(), start=6):
            ws.cell(row=row_idx, column=1, value=date.strftime('%Y-%m-%d'))
            ws.cell(row=row_idx, column=2, value=float(row['Close'])).number_format = '#,##0.00'
            # (Thêm 5 cột return)
            ws.cell(row=row_idx, column=3, value=float(row['Daily']) if pd.notna(row['Daily']) else None).number_format = '0.00'
            ws.cell(row=row_idx, column=4, value=float(row['Weekly']) if pd.notna(row['Weekly']) else None).number_format = '0.00'
            ws.cell(row=row_idx, column=5, value=float(row['Monthly']) if pd.notna(row['Monthly']) else None).number_format = '0.00'
        
            ws.cell(row=row_idx, column=6, value=float(row['YoY']) if pd.notna(row['YoY']) else None).number_format = '0.00'
            ws.cell(row=row_idx, column=7, value=float(row['YTD']) if pd.notna(row['YTD']) else None).number_format = '0.00'
            
            for col in range(3, 8):
                cell = ws.cell(row=row_idx, column=col)
                if cell.value and cell.value > 0: cell.font = Font(color='00B050')
                elif cell.value and cell.value < 0: cell.font = Font(color='FF0000')
        
        ws.column_dimensions['A'].width = 12; ws.column_dimensions['B'].width = 12
        for col in ['C', 'D', 'E', 'F', 'G']: ws.column_dimensions[col].width = 11
        
        img = OpenpyxlImage(img_buffer)
        img.width = 900; img.height = 450
        ws.add_image(img, 'H1')
    
    # Save Excel (luôn lưu local)
    wb.save(output_file)
    print(f"\\n✅ Đã xuất thành công file Excel (local): {output_file}")
    print(f"📊 Tổng số commodity: {len(commodities)}")
    print(f"📁 Excel file: {output_file}")