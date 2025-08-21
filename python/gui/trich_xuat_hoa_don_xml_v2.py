import os
import xml.etree.ElementTree as ET
import csv
import glob
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import pandas as pd
from tkcalendar import DateEntry
# lọc hoá đơn theo thời gian, có copy
def extract_invoice_data(xml_file_path):
    """
    Trích xuất dữ liệu từ file XML hóa đơn
    
    Args:
        xml_file_path (str): Đường dẫn đến file XML
        
    Returns:
        list: Danh sách các dòng dữ liệu đã trích xuất
    """
    try:
        # Phân tích file XML
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        
        # Namespace mặc định (nếu có)
        namespaces = {'': ''}
        
        # Trích xuất thông tin chung
        date_str = root.find('.//NLap').text if root.find('.//NLap') is not None else ""
        invoice_number = root.find('.//SHDon').text if root.find('.//SHDon') is not None else ""
        invoice_series = root.find('.//KHHDon').text if root.find('.//KHHDon') is not None else ""
        
        # Định dạng lại ngày tháng (từ YYYY-MM-DD thành DD/MM/YYYY)
        try:
            if date_str:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                formatted_date = date_obj.strftime('%d/%m/%Y')
                # Lưu cả đối tượng datetime để lọc dễ dàng
                date_obj_for_filter = date_obj
            else:
                formatted_date = ""
                date_obj_for_filter = None
        except ValueError:
            formatted_date = date_str
            date_obj_for_filter = None
        
        # Tìm thông tin người bán
        seller_name = root.find('.//NBan/Ten').text if root.find('.//NBan/Ten') is not None else ""
        seller_tax_code = root.find('.//NBan/MST').text if root.find('.//NBan/MST') is not None else ""
        
        # Tìm thông tin người mua
        buyer_name = root.find('.//NMua/Ten').text if root.find('.//NMua/Ten') is not None else ""
        buyer_tax_code = root.find('.//NMua/MST').text if root.find('.//NMua/MST') is not None else ""
        
        # Tìm danh sách hàng hóa dịch vụ
        products = root.findall('.//HHDVu')
        
        result_rows = []
        
        # Nếu không có sản phẩm, trả về một dòng với thông tin chung
        if not products:
            result_rows.append({
                'invoice_date': formatted_date,
                'invoice_date_obj': date_obj_for_filter,  # Thêm trường này để tiện lọc
                'invoice_number': invoice_number,
                'invoice_series': invoice_series,
                'seller_name': seller_name,
                'seller_tax_code': seller_tax_code,
                'buyer_name': buyer_name,
                'buyer_tax_code': buyer_tax_code,
                'stt': '',
                'product_name': '',
                'unit': '',
                'quantity': '',
                'unit_price': '',
                'amount': ''
            })
        
        # Xử lý từng sản phẩm trong hóa đơn
        for product in products:
            stt = product.find('STT').text if product.find('STT') is not None else ""
            product_name = product.find('THHDVu').text if product.find('THHDVu') is not None else ""
            unit = product.find('DVTinh').text if product.find('DVTinh') is not None else ""
            quantity = product.find('SLuong').text if product.find('SLuong') is not None else ""
            unit_price = product.find('DGia').text if product.find('DGia') is not None else ""
            amount = product.find('ThTien').text if product.find('ThTien') is not None else ""
            
            # Định dạng lại số lượng, đơn giá và thành tiền (loại bỏ phần thập phân không cần thiết)
            try:
                if quantity and quantity.endswith(".000000"):
                    quantity = quantity.replace(".000000", "")
                if unit_price and unit_price.endswith(".000000"):
                    unit_price = unit_price.replace(".000000", "")
                if amount and amount.endswith(".000000"):
                    amount = amount.replace(".000000", "")
            except:
                pass
            
            # Tạo dòng dữ liệu
            row = {
                'invoice_date': formatted_date,
                'invoice_date_obj': date_obj_for_filter,  # Thêm trường này để tiện lọc
                'invoice_number': invoice_number,
                'invoice_series': invoice_series,
                'seller_name': seller_name,
                'seller_tax_code': seller_tax_code,
                'buyer_name': buyer_name,
                'buyer_tax_code': buyer_tax_code,
                'stt': stt,
                'product_name': product_name,
                'unit': unit,
                'quantity': quantity,
                'unit_price': unit_price,
                'amount': amount
            }
            
            result_rows.append(row)
        
        return result_rows, None
        
    except Exception as e:
        return [], f"Lỗi khi xử lý file {os.path.basename(xml_file_path)}: {str(e)}"

def process_invoices(input_folder, progress_callback=None, status_callback=None, data_callback=None):
    """
    Xử lý tất cả các file XML trong thư mục và lưu kết quả vào file CSV
    
    Args:
        input_folder (str): Thư mục chứa các file XML
        progress_callback (function): Hàm callback để cập nhật tiến trình
        status_callback (function): Hàm callback để cập nhật trạng thái
        data_callback (function): Hàm callback để trả về dữ liệu đã trích xuất
    
    Returns:
        bool: True nếu thành công, False nếu thất bại
        str: Thông báo kết quả
        list: Danh sách các dòng dữ liệu đã trích xuất
        list: Danh sách các trường dữ liệu
    """
    # Tìm tất cả các file XML trong thư mục
    xml_files = glob.glob(os.path.join(input_folder, "*.xml"))
    
    if not xml_files:
        message = f"Không tìm thấy file XML trong thư mục {input_folder}"
        if status_callback:
            status_callback(message)
        return False, message, [], []
    
    # Danh sách chứa tất cả các dòng dữ liệu
    all_rows = []
    error_messages = []
    
    # Cập nhật tiến trình
    total_files = len(xml_files)
    
    # Xử lý từng file XML
    for i, xml_file in enumerate(xml_files):
        file_name = os.path.basename(xml_file)
        if status_callback:
            status_callback(f"Đang xử lý file: {file_name} ({i+1}/{total_files})")
        
        rows, error = extract_invoice_data(xml_file)
        if error:
            error_messages.append(error)
        all_rows.extend(rows)
        
        # Cập nhật thanh tiến trình
        if progress_callback:
            progress_value = (i + 1) / total_files * 100
            progress_callback(progress_value)
    
    # Nếu không có dữ liệu, thoát
    if not all_rows:
        message = "Không có dữ liệu được trích xuất"
        if status_callback:
            status_callback(message)
        return False, message, [], []
    
    # Xác định các trường dữ liệu
    fieldnames = [
        'invoice_date', 'invoice_number', 'invoice_series', 
        'seller_name', 'seller_tax_code', 'buyer_name', 'buyer_tax_code',
        'stt', 'product_name', 'unit', 'quantity', 'unit_price', 'amount'
    ]
    
    result = f"Đã trích xuất dữ liệu từ {total_files} file XML.\n"
    result += f"Tổng số dòng dữ liệu: {len(all_rows)}"
    
    if error_messages:
        result += "\n\nCác lỗi gặp phải:\n" + "\n".join(error_messages)
    
    if status_callback:
        status_callback(result)
        
    # Trả về dữ liệu để hiển thị trong bảng
    if data_callback:
        data_callback(fieldnames, all_rows)
    
    return True, result, all_rows, fieldnames

class InvoiceProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Trích xuất dữ liệu hóa đơn XML")
        self.root.geometry("900x650")  # Tăng kích thước để có không gian cho bảng dữ liệu
        self.root.minsize(800, 600)    # Tăng kích thước tối thiểu
        self.root.resizable(True, True)
        
        # Biến để lưu đường dẫn
        self.input_folder = tk.StringVar()
        
        # Biến để lưu trữ dữ liệu đã trích xuất
        self.extracted_data = []
        self.filtered_data = []  # Thêm biến để lưu dữ liệu đã lọc
        self.column_names = []
        
        # Khởi tạo giao diện
        self.create_widgets()
        
        # Thiết lập vị trí mặc định của cột
        self.setup_default_column_widths()
        
        # Thiết lập popup menu cho bảng dữ liệu
        self.setup_context_menu()
    
    def normalize_path(self, path):
        """Chuyển đổi đường dẫn sang định dạng chuẩn Windows"""
        if path:
            return path.replace('/', '\\')
        return path
    
    def create_widgets(self):
        # Tạo frame chính
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tiêu đề
        title_label = ttk.Label(main_frame, text="CÔNG CỤ TRÍCH XUẤT DỮ LIỆU HÓA ĐƠN XML", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Frame chứa các điều khiển
        control_frame = ttk.LabelFrame(main_frame, text="Cấu hình", padding="10")
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Thư mục đầu vào
        input_frame = ttk.Frame(control_frame)
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        input_label = ttk.Label(input_frame, text="Thư mục chứa file XML:", width=20)
        input_label.pack(side=tk.LEFT, padx=5)
        
        input_entry = ttk.Entry(input_frame, textvariable=self.input_folder, width=50)
        input_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        input_button = ttk.Button(input_frame, text="Chọn thư mục", command=self.select_input_folder)
        input_button.pack(side=tk.LEFT, padx=5)
        
        # Các nút thực thi và đóng
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Nút Đóng (bên trái)
        self.close_button = ttk.Button(button_frame, text="Đóng", command=self.root.destroy, width=15)
        self.close_button.pack(side=tk.LEFT, padx=5)
        
        # Nút xuất dữ liệu (không hiển thị ban đầu - sẽ hiển thị sau khi có dữ liệu)
        self.export_frame = ttk.Frame(button_frame)
        self.export_frame.pack(side=tk.RIGHT)
        
        self.export_xlsx_button = ttk.Button(
            self.export_frame, text="Xuất Excel (XLSX)", 
            command=lambda: self.export_data("xlsx"), width=20, state=tk.DISABLED
        )
        self.export_xlsx_button.pack(side=tk.RIGHT, padx=5)
        
        self.export_csv_button = ttk.Button(
            self.export_frame, text="Xuất CSV", 
            command=lambda: self.export_data("csv"), width=15, state=tk.DISABLED
        )
        self.export_csv_button.pack(side=tk.RIGHT, padx=5)
        
        # Thêm nút Sao chép tất cả dữ liệu
        self.copy_all_button = ttk.Button(
            self.export_frame, text="Sao chép tất cả", 
            command=self.copy_all_data, width=15, state=tk.DISABLED
        )
        self.copy_all_button.pack(side=tk.RIGHT, padx=5)
        
        # Nút Run (bên phải)
        self.run_button = ttk.Button(button_frame, text="Run", command=self.process, width=15)
        self.run_button.pack(side=tk.RIGHT, padx=5)
        
        # Tạo notebook để chứa nhiều tab
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Tab tiến trình
        progress_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(progress_frame, text="Tiến trình")
        
        self.progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        # Trạng thái
        status_frame = ttk.Frame(progress_frame)
        status_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.status_text = tk.Text(status_frame, height=15, wrap=tk.WORD)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        # Tab bảng dữ liệu
        data_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(data_frame, text="Dữ liệu")
        
        # Thêm phần lọc dữ liệu theo ngày tháng
        filter_frame = ttk.LabelFrame(data_frame, text="Lọc dữ liệu theo thời gian", padding="10")
        filter_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Tạo frame cho chọn ngày
        date_filter_frame = ttk.Frame(filter_frame)
        date_filter_frame.pack(fill=tk.X, pady=5)
        
        # Từ ngày
        from_date_label = ttk.Label(date_filter_frame, text="Từ ngày:", width=10)
        from_date_label.pack(side=tk.LEFT, padx=5)
        
        # Sử dụng DateEntry thay vì Entry thường
        self.from_date = DateEntry(date_filter_frame, width=12, background='darkblue',
                                  foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.from_date.pack(side=tk.LEFT, padx=5)
        
        # Đến ngày
        to_date_label = ttk.Label(date_filter_frame, text="Đến ngày:", width=10)
        to_date_label.pack(side=tk.LEFT, padx=5)
        
        self.to_date = DateEntry(date_filter_frame, width=12, background='darkblue',
                                foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.to_date.pack(side=tk.LEFT, padx=5)
        
        # Nút áp dụng bộ lọc
        self.apply_filter_button = ttk.Button(
            date_filter_frame, text="Áp dụng lọc", 
            command=self.apply_date_filter, width=15, state=tk.DISABLED
        )
        self.apply_filter_button.pack(side=tk.LEFT, padx=10)
        
        # Nút hiển thị tất cả
        self.show_all_button = ttk.Button(
            date_filter_frame, text="Hiển thị tất cả", 
            command=self.show_all_data, width=15, state=tk.DISABLED
        )
        self.show_all_button.pack(side=tk.LEFT, padx=5)
        
        # Thông tin dữ liệu đang hiển thị
        self.data_info_frame = ttk.Frame(filter_frame)
        self.data_info_frame.pack(fill=tk.X, pady=5)
        
        self.data_info_label = ttk.Label(self.data_info_frame, text="Dữ liệu: 0 dòng")
        self.data_info_label.pack(side=tk.LEFT, padx=5)
        
        # Tạo bảng dữ liệu với thanh cuộn
        self.create_data_table(data_frame)
        
        # Khởi tạo giá trị mặc định
        self.update_status("Hãy chọn thư mục chứa file XML để bắt đầu")
    
    def create_data_table(self, parent_frame):
        """Tạo bảng dữ liệu với thanh cuộn"""
        # Frame chứa bảng và thanh cuộn
        table_frame = ttk.Frame(parent_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tạo thanh cuộn
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Tạo bảng dữ liệu (Treeview)
        self.data_table = ttk.Treeview(
            table_frame,
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            show="headings"  # Chỉ hiển thị phần header và dữ liệu, không hiển thị cột đầu tiên (icon)
        )
        
        # Thiết lập thanh cuộn
        v_scrollbar.config(command=self.data_table.yview)
        h_scrollbar.config(command=self.data_table.xview)
        
        # Hiển thị bảng dữ liệu
        self.data_table.pack(fill=tk.BOTH, expand=True)
        
        # Thiết lập định dạng các cột (ban đầu không có cột)
        self.setup_table_columns([])
        
        # Liên kết sự kiện nhấp chuột
        self.data_table.bind("<Button-3>", self.show_context_menu)  # Chuột phải
        self.data_table.bind("<Button-1>", self.on_click)  # Chuột trái
        self.data_table.bind("<<TreeviewSelect>>", self.on_selection_changed)  # Khi thay đổi lựa chọn
        
        # Thêm hỗ trợ phím tắt
        self.data_table.bind("<Control-c>", lambda event: self.copy_row())  # Ctrl+C để sao chép dòng
        self.root.bind("<Control-a>", lambda event: self.select_all_rows())  # Ctrl+A để chọn tất cả
    
    def setup_context_menu(self):
        """Thiết lập menu ngữ cảnh khi nhấp chuột phải vào bảng"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Sao chép dòng", command=self.copy_row)
        self.context_menu.add_command(label="Sao chép giá trị ô", command=self.copy_cell)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Xuất dòng đã chọn", command=self.export_selected_rows)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Sao chép tất cả", command=self.copy_all_data)
    
    def show_context_menu(self, event):
        """Hiển thị menu ngữ cảnh khi nhấp chuột phải"""
        # Lấy dòng được chọn
        selected_iid = self.data_table.identify_row(event.y)
        if selected_iid:
            # Đánh dấu dòng được chọn
            self.data_table.selection_set(selected_iid)
            # Lưu vị trí chuột
            self.context_menu_position = (event.x_root, event.y_root)
            # Lưu thông tin ô được chọn
            self.selected_column = self.data_table.identify_column(event.x)
            # Hiển thị menu ngữ cảnh
            self.context_menu.post(event.x_root, event.y_root)
    
    def on_click(self, event):
        """Xử lý sự kiện click chuột trái"""
        # Lưu vị trí ô được chọn
        self.selected_column = self.data_table.identify_column(event.x)
    
    def on_selection_changed(self, event):
        """Xử lý sự kiện khi thay đổi lựa chọn"""
        # Cập nhật thông tin về số dòng đã chọn
        self.update_data_info()
    
    def select_all_rows(self):
        """Chọn tất cả các dòng trong bảng"""
        for item in self.data_table.get_children():
            self.data_table.selection_add(item)
        # Cập nhật thông tin
        self.update_data_info()
    
    def copy_row(self):
        """Sao chép dòng đang được chọn vào clipboard"""
        selected_items = self.data_table.selection()
        if not selected_items:
            return
        
        item_id = selected_items[0]
        values = self.data_table.item(item_id, 'values')
        
        if values:
            # Tạo chuỗi từ các giá trị, phân cách bởi tab
            row_text = '\t'.join(str(v) for v in values)
            
            # Sao chép vào clipboard
            self.root.clipboard_clear()
            self.root.clipboard_append(row_text)
            
            # Thông báo thành công
            messagebox.showinfo("Sao chép", "Đã sao chép dòng vào clipboard")
    
    def copy_cell(self):
        """Sao chép giá trị ô đang được chọn vào clipboard"""
        selected_items = self.data_table.selection()
        if not selected_items or not self.selected_column:
            return
        
        item_id = selected_items[0]
        column_index = int(self.selected_column.replace('#', '')) - 1
        values = self.data_table.item(item_id, 'values')
        
        if values and 0 <= column_index < len(values):
            # Lấy giá trị ô
            cell_value = values[column_index]
            
            # Sao chép vào clipboard
            self.root.clipboard_clear()
            self.root.clipboard_append(str(cell_value))
            
            # Thông báo thành công
            messagebox.showinfo("Sao chép", "Đã sao chép giá trị ô vào clipboard")
    
    def copy_all_data(self):
        """Sao chép tất cả dữ liệu đang hiển thị vào clipboard"""
        if not self.filtered_data or not self.column_names:
            messagebox.showerror("Lỗi", "Không có dữ liệu để sao chép")
            return
        
        # Tạo đối tượng DataFrame từ dữ liệu
        df = pd.DataFrame(self.filtered_data)
        
        # Tạo tiêu đề cho các cột
        column_display_names = {
            'invoice_date': 'Ngày hóa đơn',
            'invoice_number': 'Số hóa đơn',
            'invoice_series': 'Ký hiệu',
            'seller_name': 'Tên người bán',
            'seller_tax_code': 'MST người bán',
            'buyer_name': 'Tên người mua',
            'buyer_tax_code': 'MST người mua',
            'stt': 'STT',
            'product_name': 'Tên sản phẩm',
            'unit': 'Đơn vị tính',
            'quantity': 'Số lượng',
            'unit_price': 'Đơn giá',
            'amount': 'Thành tiền'
        }
        
        # Đổi tên cột
        df_renamed = df[self.column_names].copy()
        df_renamed.columns = [column_display_names.get(col, col) for col in self.column_names]
        
        # Chuyển đổi DataFrame thành chuỗi định dạng tab
        clipboard_str = df_renamed.to_csv(sep='\t', index=False)
        
        # Sao chép vào clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(clipboard_str)
        
        # Thông báo thành công
        messagebox.showinfo("Sao chép", f"Đã sao chép {len(self.filtered_data)} dòng dữ liệu vào clipboard")
    
    def export_selected_rows(self):
        """Xuất các dòng đã chọn sang clipboard hoặc file"""
        selected_items = self.data_table.selection()
        if not selected_items:
            messagebox.showinfo("Thông báo", "Vui lòng chọn ít nhất một dòng để xuất")
            return
        
        # Lấy dữ liệu từ các dòng đã chọn
        selected_rows = []
        for item_id in selected_items:
            item_index = self.data_table.index(item_id)
            if 0 <= item_index < len(self.filtered_data):
                selected_rows.append(self.filtered_data[item_index])
        
        if not selected_rows:
            return
        
        # Tạo menu con để chọn cách xuất
        export_menu = tk.Menu(self.context_menu, tearoff=0)
        export_menu.add_command(label="Sao chép vào clipboard", 
                               command=lambda: self.copy_rows_to_clipboard(selected_rows))
        export_menu.add_command(label="Xuất sang CSV", 
                               command=lambda: self.export_rows_to_file(selected_rows, "csv"))
        export_menu.add_command(label="Xuất sang Excel", 
                               command=lambda: self.export_rows_to_file(selected_rows, "xlsx"))
        
        # Hiển thị menu tại vị trí chuột
        export_menu.post(self.context_menu_position[0], self.context_menu_position[1])
    
    def copy_rows_to_clipboard(self, rows):
        """Sao chép các dòng dữ liệu đã chọn vào clipboard"""
        if not rows:
            return
        
        # Tạo DataFrame từ dữ liệu
        df = pd.DataFrame(rows)
        df = df[self.column_names]  # Chỉ lấy các cột hiển thị
        
        # Chuyển đổi sang định dạng tab
        clipboard_str = df.to_csv(sep='\t', index=False)
        
        # Sao chép vào clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(clipboard_str)
        
        # Thông báo
        messagebox.showinfo("Sao chép", f"Đã sao chép {len(rows)} dòng dữ liệu vào clipboard")
    
    def export_rows_to_file(self, rows, file_type):
        """Xuất các dòng dữ liệu đã chọn sang file"""
        if not rows:
            return
        
        # Loại bỏ trường invoice_date_obj khi xuất dữ liệu
        export_data = []
        for row in rows:
            export_row = {k: v for k, v in row.items() if k != 'invoice_date_obj'}
            export_data.append(export_row)
        
        # Tạo tên file mặc định
        folder_name = os.path.basename(self.input_folder.get())
        default_filename = f"{folder_name}_selected_{len(rows)}_rows"
        
        if file_type == "csv":
            self.export_rows_to_csv(export_data, default_filename)
        elif file_type == "xlsx":
            self.export_rows_to_excel(export_data, default_filename)
    
    def export_rows_to_csv(self, data, default_filename):
        """Xuất dữ liệu sang file CSV"""
        file_path = filedialog.asksaveasfilename(
            title="Lưu file CSV",
            filetypes=[("CSV files", "*.csv")],
            defaultextension=".csv",
            initialfile=f"{default_filename}.csv"
        )
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=self.column_names)
                writer.writeheader()
                writer.writerows(data)
            
            messagebox.showinfo("Thành công", f"Đã xuất {len(data)} dòng dữ liệu ra file CSV: {file_path}")
            # Mở thư mục chứa file CSV
            os.startfile(os.path.dirname(file_path))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi xuất file CSV: {str(e)}")
    
    def export_rows_to_excel(self, data, default_filename):
        """Xuất dữ liệu sang file Excel"""
        file_path = filedialog.asksaveasfilename(
            title="Lưu file Excel",
            filetypes=[("Excel files", "*.xlsx")],
            defaultextension=".xlsx",
            initialfile=f"{default_filename}.xlsx"
        )
        if not file_path:
            return
        
        try:
            # Tạo DataFrame từ dữ liệu
            df = pd.DataFrame(data)
            df = df[self.column_names]  # Đảm bảo thứ tự cột
            
            # Tạo ExcelWriter để format dữ liệu
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Dữ liệu Hóa đơn')
                
                # Lấy worksheet để format
                worksheet = writer.sheets['Dữ liệu Hóa đơn']
                
                # Tự động điều chỉnh độ rộng cột
                for idx, col in enumerate(df.columns):
                    max_length = 0
                    column = worksheet.cell(row=1, column=idx+1).column_letter
                    for i in range(1, len(df) + 2):
                        try:
                            cell_length = len(str(worksheet.cell(row=i, column=idx+1).value))
                            if cell_length > max_length:
                                max_length = cell_length
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column].width = adjusted_width
            
            messagebox.showinfo("Thành công", f"Đã xuất {len(data)} dòng dữ liệu ra file Excel: {file_path}")
            # Mở thư mục chứa file Excel
            os.startfile(os.path.dirname(file_path))
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi xuất file Excel: {str(e)}")
    
    def setup_table_columns(self, column_names):
        """Thiết lập cột cho bảng dữ liệu"""
        # Xóa tất cả các cột cũ (nếu có)
        for col in self.data_table["columns"]:
            self.data_table.heading(col, text="")
        
        # Đặt lại cột mới
        self.data_table["columns"] = column_names
        
        # Thiết lập tiêu đề và định dạng cho mỗi cột
        column_display_names = {
            'invoice_date': 'Ngày hóa đơn',
            'invoice_number': 'Số hóa đơn',
            'invoice_series': 'Ký hiệu',
            'seller_name': 'Tên người bán',
            'seller_tax_code': 'MST người bán',
            'buyer_name': 'Tên người mua',
            'buyer_tax_code': 'MST người mua',
            'stt': 'STT',
            'product_name': 'Tên sản phẩm',
            'unit': 'Đơn vị tính',
            'quantity': 'Số lượng',
            'unit_price': 'Đơn giá',
            'amount': 'Thành tiền'
        }
        
        for col in column_names:
            display_name = column_display_names.get(col, col)
            self.data_table.heading(col, text=display_name)
            self.data_table.column(col, width=100, minwidth=50)  # Độ rộng mặc định
    
    def setup_default_column_widths(self):
        """Thiết lập độ rộng mặc định cho các cột thường gặp"""
        column_widths = {
            'invoice_date': 100,
            'invoice_number': 80,
            'invoice_series': 80,
            'seller_name': 200,
            'seller_tax_code': 120,
            'buyer_name': 200,
            'buyer_tax_code': 120,
            'stt': 50,
            'product_name': 250,
            'unit': 80,
            'quantity': 80,
            'unit_price': 100,
            'amount': 100
        }
        
        # Lưu lại để sử dụng khi cần thiết lập cột
        self.column_widths = column_widths
    
    def populate_table(self, column_names, data, is_filtered=False):
        """Cập nhật dữ liệu cho bảng"""
        # Lưu lại dữ liệu và tên cột
        self.column_names = column_names
        
        if not is_filtered:
            self.extracted_data = data
            self.filtered_data = data.copy()  # Ban đầu, dữ liệu lọc giống dữ liệu gốc
        else:
            self.filtered_data = data  # Cập nhật dữ liệu đã lọc
        
        # Thiết lập lại cột cho bảng
        self.setup_table_columns(column_names)
        
        # Xóa dữ liệu cũ
        for item in self.data_table.get_children():
            self.data_table.delete(item)
        
        # Thêm dữ liệu mới
        for row in data:
            values = [row.get(col, "") for col in column_names]
            self.data_table.insert("", tk.END, values=values)
        
        # Áp dụng độ rộng mặc định cho các cột
        for col in column_names:
            width = self.column_widths.get(col, 100)
            self.data_table.column(col, width=width)
        
        # Kích hoạt các nút xuất dữ liệu, sao chép dữ liệu và lọc dữ liệu
        self.export_csv_button.config(state=tk.NORMAL)
        self.export_xlsx_button.config(state=tk.NORMAL)
        self.apply_filter_button.config(state=tk.NORMAL)
        self.show_all_button.config(state=tk.NORMAL)
        self.copy_all_button.config(state=tk.NORMAL)
        
        # Cập nhật thông tin dữ liệu đang hiển thị
        self.update_data_info()
        
        # Chuyển đến tab Dữ liệu
        self.notebook.select(1)  # Chọn tab thứ hai (index 1)

    def update_data_info(self):
        """Cập nhật thông tin về dữ liệu đang hiển thị"""
        total_rows = len(self.filtered_data)
        total_original = len(self.extracted_data)
        
        if total_rows == total_original:
            self.data_info_label.config(text=f"Đang hiển thị: {total_rows} dòng dữ liệu")
        else:
            self.data_info_label.config(text=f"Đang hiển thị: {total_rows}/{total_original} dòng dữ liệu")
        
        # Thêm thông tin về số dòng đã chọn
        selected_count = len(self.data_table.selection())
        if selected_count > 0:
            self.data_info_label.config(text=f"Đang hiển thị: {total_rows}/{total_original} dòng dữ liệu | Đã chọn: {selected_count} dòng")
    
    def apply_date_filter(self):
        """Áp dụng bộ lọc thời gian"""
        try:
            # Lấy ngày từ DateEntry
            from_date = self.from_date.get_date()
            to_date = self.to_date.get_date()
            
            # Điều chỉnh to_date để bao gồm cả ngày kết thúc
            to_date = datetime.combine(to_date, datetime.max.time())
            
            # Lọc dữ liệu
            filtered_data = []
            for row in self.extracted_data:
                date_obj = row.get('invoice_date_obj')
                if date_obj and from_date <= date_obj.date() <= to_date.date():
                    filtered_data.append(row)
            
            # Hiển thị dữ liệu đã lọc
            self.populate_table(self.column_names, filtered_data, is_filtered=True)
            
            # Cập nhật thông tin
            if not filtered_data:
                messagebox.showinfo("Thông báo", "Không có dữ liệu trong khoảng thời gian đã chọn")
                
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi lọc dữ liệu: {str(e)}")
    
    def show_all_data(self):
        """Hiển thị tất cả dữ liệu (bỏ bộ lọc)"""
        self.populate_table(self.column_names, self.extracted_data)
    
    def select_input_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa file XML")
        if folder:
            # Chuyển đổi sang định dạng đường dẫn Windows
            folder = self.normalize_path(folder)
            self.input_folder.set(folder)
    
    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.root.update_idletasks()
    
    def update_status(self, message):
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, message)
        self.status_text.see(tk.END)
    
    def export_data(self, file_type):
        """Xuất dữ liệu sang file CSV hoặc XLSX"""
        if not self.filtered_data:
            messagebox.showerror("Lỗi", "Không có dữ liệu để xuất")
            return
        
        # Kiểm tra nếu có dòng đã chọn, đề xuất xuất các dòng được chọn
        selected_items = self.data_table.selection()
        if selected_items and len(selected_items) > 0:
            result = messagebox.askyesnocancel(
                "Xuất dữ liệu", 
                f"Bạn đã chọn {len(selected_items)} dòng.\nBạn muốn xuất chỉ các dòng đã chọn?\n\n- Chọn YES để xuất các dòng đã chọn\n- Chọn NO để xuất tất cả\n- Chọn CANCEL để hủy"
            )
            
            if result is None:  # Cancel
                return
            
            if result:  # Yes - xuất các dòng đã chọn
                selected_rows = []
                for item_id in selected_items:
                    item_index = self.data_table.index(item_id)
                    if 0 <= item_index < len(self.filtered_data):
                        selected_rows.append(self.filtered_data[item_index])
                
                if file_type == "csv":
                    self.export_rows_to_csv(selected_rows, f"{os.path.basename(self.input_folder.get())}_selected")
                else:
                    self.export_rows_to_excel(selected_rows, f"{os.path.basename(self.input_folder.get())}_selected")
                return
        
        # Xuất tất cả dữ liệu
        # Tạo folder name mặc định từ thư mục XML
        folder_name = os.path.basename(self.input_folder.get())
        
        # Loại bỏ trường invoice_date_obj khi xuất dữ liệu
        export_data = []
        for row in self.filtered_data:
            export_row = {k: v for k, v in row.items() if k != 'invoice_date_obj'}
            export_data.append(export_row)
        
        if file_type == "csv":
            file_path = filedialog.asksaveasfilename(
                title="Lưu file CSV",
                filetypes=[("CSV files", "*.csv")],
                defaultextension=".csv",
                initialfile=f"{folder_name}.csv"
            )
            if not file_path:
                return
            
            try:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=self.column_names)
                    writer.writeheader()
                    writer.writerows(export_data)
                
                messagebox.showinfo("Thành công", f"Đã xuất dữ liệu ra file CSV: {file_path}")
                # Mở thư mục chứa file CSV
                os.startfile(os.path.dirname(file_path))
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi xuất file CSV: {str(e)}")
                
        elif file_type == "xlsx":
            file_path = filedialog.asksaveasfilename(
                title="Lưu file Excel",
                filetypes=[("Excel files", "*.xlsx")],
                defaultextension=".xlsx",
                initialfile=f"{folder_name}.xlsx"
            )
            if not file_path:
                return
            
            try:
                # Tạo DataFrame từ dữ liệu
                df = pd.DataFrame(export_data)
                
                # Tạo ExcelWriter để format dữ liệu
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Dữ liệu Hóa đơn')
                    
                    # Lấy worksheet để format
                    worksheet = writer.sheets['Dữ liệu Hóa đơn']
                    
                    # Tự động điều chỉnh độ rộng cột
                    for idx, col in enumerate(df.columns):
                        max_length = 0
                        column = worksheet.cell(row=1, column=idx+1).column_letter
                        for i in range(1, len(df) + 2):
                            try:
                                cell_length = len(str(worksheet.cell(row=i, column=idx+1).value))
                                if cell_length > max_length:
                                    max_length = cell_length
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column].width = adjusted_width
                
                messagebox.showinfo("Thành công", f"Đã xuất dữ liệu ra file Excel: {file_path}")
                # Mở thư mục chứa file Excel
                os.startfile(os.path.dirname(file_path))
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi xuất file Excel: {str(e)}")
    
    def process(self):
        input_folder = self.input_folder.get()
        
        if not input_folder:
            messagebox.showerror("Lỗi", "Vui lòng chọn thư mục đầu vào")
            return
        
        # Kiểm tra thư mục đầu vào tồn tại
        if not os.path.exists(input_folder):
            messagebox.showerror("Lỗi", f"Thư mục {input_folder} không tồn tại")
            return
        
        # Vô hiệu hóa nút Run trong quá trình xử lý
        self.run_button.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        
        # Vô hiệu hóa các nút xuất dữ liệu
        self.export_csv_button.config(state=tk.DISABLED)
        self.export_xlsx_button.config(state=tk.DISABLED)
        self.copy_all_button.config(state=tk.DISABLED)
        
        # Chuyển đến tab tiến trình
        self.notebook.select(0)  # Chọn tab đầu tiên (index 0)
        
        self.update_status("Đang bắt đầu xử lý...")
        
        # Tạo và bắt đầu luồng xử lý
        def process_thread():
            success, message, data, fieldnames = process_invoices(
                input_folder, 
                progress_callback=self.update_progress,
                status_callback=self.update_status,
                data_callback=self.populate_table
            )
            
            # Kích hoạt lại nút Run sau khi hoàn thành
            self.root.after(0, lambda: self.run_button.config(state=tk.NORMAL))
            
            if success:
                self.root.after(0, lambda: messagebox.showinfo("Thành công", "Đã hoàn thành xử lý dữ liệu!"))
            else:
                self.root.after(0, lambda: messagebox.showerror("Lỗi", message))
        
        threading.Thread(target=process_thread, daemon=True).start()

    # Thêm các phương thức mới để tăng cường khả năng sao chép dữ liệu
    
    def format_table_for_copy(self, rows):
        """
        Định dạng dữ liệu bảng để sao chép sang Excel/LibreOffice
        với định dạng đẹp hơn
        """
        if not rows:
            return ""
        
        # Tạo tiêu đề cho các cột
        column_display_names = {
            'invoice_date': 'Ngày hóa đơn',
            'invoice_number': 'Số hóa đơn',
            'invoice_series': 'Ký hiệu',
            'seller_name': 'Tên người bán',
            'seller_tax_code': 'MST người bán',
            'buyer_name': 'Tên người mua',
            'buyer_tax_code': 'MST người mua',
            'stt': 'STT',
            'product_name': 'Tên sản phẩm',
            'unit': 'Đơn vị tính',
            'quantity': 'Số lượng',
            'unit_price': 'Đơn giá',
            'amount': 'Thành tiền'
        }
        
        # Tạo tiêu đề
        headers = [column_display_names.get(col, col) for col in self.column_names]
        result = '\t'.join(headers) + '\n'
        
        # Thêm dữ liệu
        for row in rows:
            values = [str(row.get(col, "")) for col in self.column_names]
            result += '\t'.join(values) + '\n'
        
        return result
    
    def copy_data_and_notify(self, data_str, count):
        """Sao chép dữ liệu vào clipboard và thông báo"""
        self.root.clipboard_clear()
        self.root.clipboard_append(data_str)
        
        # Thông báo thành công
        messagebox.showinfo("Sao chép", f"Đã sao chép {count} dòng dữ liệu vào clipboard")
    
    def copy_selected_columns(self):
        """Sao chép chỉ các cột đã chọn cho tất cả các dòng"""
        # Lấy các cột được chọn
        if not hasattr(self, 'selected_column') or not self.selected_column:
            messagebox.showinfo("Thông báo", "Vui lòng chọn ít nhất một cột để sao chép")
            return
        
        column_index = int(self.selected_column.replace('#', '')) - 1
        if column_index < 0 or column_index >= len(self.column_names):
            return
        
        selected_column_name = self.column_names[column_index]
        column_display_name = {
            'invoice_date': 'Ngày hóa đơn',
            'invoice_number': 'Số hóa đơn',
            'invoice_series': 'Ký hiệu',
            'seller_name': 'Tên người bán',
            'seller_tax_code': 'MST người bán',
            'buyer_name': 'Tên người mua',
            'buyer_tax_code': 'MST người mua',
            'stt': 'STT',
            'product_name': 'Tên sản phẩm',
            'unit': 'Đơn vị tính',
            'quantity': 'Số lượng',
            'unit_price': 'Đơn giá',
            'amount': 'Thành tiền'
        }.get(selected_column_name, selected_column_name)
        
        # Sao chép giá trị từ cột đã chọn
        result = f"{column_display_name}\n"
        for row in self.filtered_data:
            result += f"{row.get(selected_column_name, '')}\n"
        
        # Sao chép vào clipboard
        self.copy_data_and_notify(result, len(self.filtered_data))
    
    def copy_selected_columns_for_selected_rows(self):
        """Sao chép cột đã chọn chỉ cho các dòng đã chọn"""
        # Lấy các dòng đã chọn
        selected_items = self.data_table.selection()
        if not selected_items:
            messagebox.showinfo("Thông báo", "Vui lòng chọn ít nhất một dòng")
            return
        
        # Lấy cột được chọn
        if not hasattr(self, 'selected_column') or not self.selected_column:
            messagebox.showinfo("Thông báo", "Vui lòng chọn một cột để sao chép")
            return
        
        column_index = int(self.selected_column.replace('#', '')) - 1
        if column_index < 0 or column_index >= len(self.column_names):
            return
        
        selected_column_name = self.column_names[column_index]
        column_display_name = {
            'invoice_date': 'Ngày hóa đơn',
            'invoice_number': 'Số hóa đơn',
            'invoice_series': 'Ký hiệu',
            'seller_name': 'Tên người bán',
            'seller_tax_code': 'MST người bán',
            'buyer_name': 'Tên người mua',
            'buyer_tax_code': 'MST người mua',
            'stt': 'STT',
            'product_name': 'Tên sản phẩm',
            'unit': 'Đơn vị tính',
            'quantity': 'Số lượng',
            'unit_price': 'Đơn giá',
            'amount': 'Thành tiền'
        }.get(selected_column_name, selected_column_name)
        
        # Tạo danh sách các dòng đã chọn
        selected_rows = []
        for item_id in selected_items:
            item_index = self.data_table.index(item_id)
            if 0 <= item_index < len(self.filtered_data):
                selected_rows.append(self.filtered_data[item_index])
        
        # Sao chép giá trị từ cột đã chọn cho các dòng đã chọn
        result = f"{column_display_name}\n"
        for row in selected_rows:
            result += f"{row.get(selected_column_name, '')}\n"
        
        # Sao chép vào clipboard
        self.copy_data_and_notify(result, len(selected_rows))
    
    def filter_by_cell_value(self):
        """Lọc dữ liệu theo giá trị ô đã chọn"""
        selected_items = self.data_table.selection()
        if not selected_items or not self.selected_column:
            messagebox.showinfo("Thông báo", "Vui lòng chọn một ô để lọc")
            return
        
        # Lấy chỉ số cột
        column_index = int(self.selected_column.replace('#', '')) - 1
        if column_index < 0 or column_index >= len(self.column_names):
            return
        
        item_id = selected_items[0]
        values = self.data_table.item(item_id, 'values')
        
        if values and 0 <= column_index < len(values):
            # Lấy giá trị ô và tên cột
            cell_value = values[column_index]
            column_name = self.column_names[column_index]
            
            # Lọc dữ liệu
            filtered_data = [row for row in self.extracted_data if str(row.get(column_name, '')) == str(cell_value)]
            
            # Hiển thị dữ liệu đã lọc
            self.populate_table(self.column_names, filtered_data, is_filtered=True)
            
            # Thông báo
            messagebox.showinfo("Lọc dữ liệu", f"Đã lọc theo {cell_value} - {len(filtered_data)} kết quả")
    
    def copy_with_custom_format(self):
        """Sao chép dữ liệu với định dạng tùy chỉnh"""
        if not self.filtered_data:
            messagebox.showerror("Lỗi", "Không có dữ liệu để sao chép")
            return
        
        # Tạo cửa sổ tùy chỉnh
        format_window = tk.Toplevel(self.root)
        format_window.title("Tùy chỉnh định dạng sao chép")
        format_window.geometry("600x400")
        format_window.resizable(True, True)
        
        # Tạo frame chứa các điều khiển
        control_frame = ttk.Frame(format_window, padding="10")
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Chọn các cột cần sao chép
        columns_frame = ttk.LabelFrame(control_frame, text="Chọn cột cần sao chép", padding="10")
        columns_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Tạo biến để lưu trạng thái chọn
        column_vars = {}
        
        # Tạo frame cuộn cho danh sách cột
        columns_canvas = tk.Canvas(columns_frame, height=150)
        columns_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        columns_scrollbar = ttk.Scrollbar(columns_frame, orient=tk.VERTICAL, command=columns_canvas.yview)
        columns_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns_canvas.configure(yscrollcommand=columns_scrollbar.set)
        columns_canvas.bind('<Configure>', lambda e: columns_canvas.configure(scrollregion=columns_canvas.bbox("all")))
        
        columns_frame_inner = ttk.Frame(columns_canvas)
        columns_canvas.create_window((0, 0), window=columns_frame_inner, anchor="nw")
        
        # Tạo checkbutton cho mỗi cột
        column_display_names = {
            'invoice_date': 'Ngày hóa đơn',
            'invoice_number': 'Số hóa đơn',
            'invoice_series': 'Ký hiệu',
            'seller_name': 'Tên người bán',
            'seller_tax_code': 'MST người bán',
            'buyer_name': 'Tên người mua',
            'buyer_tax_code': 'MST người mua',
            'stt': 'STT',
            'product_name': 'Tên sản phẩm',
            'unit': 'Đơn vị tính',
            'quantity': 'Số lượng',
            'unit_price': 'Đơn giá',
            'amount': 'Thành tiền'
        }
        
        for i, col in enumerate(self.column_names):
            column_vars[col] = tk.BooleanVar(value=True)  # Mặc định chọn tất cả
            display_name = column_display_names.get(col, col)
            chk = ttk.Checkbutton(columns_frame_inner, text=display_name, variable=column_vars[col])
            chk.grid(row=i//2, column=i%2, sticky="w", padx=5, pady=2)
        
        # Nút chọn/bỏ chọn tất cả
        buttons_frame = ttk.Frame(control_frame)
        buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        def select_all():
            for var in column_vars.values():
                var.set(True)
        
        def deselect_all():
            for var in column_vars.values():
                var.set(False)
        
        select_all_button = ttk.Button(buttons_frame, text="Chọn tất cả", command=select_all, width=15)
        select_all_button.pack(side=tk.LEFT, padx=5)
        
        deselect_all_button = ttk.Button(buttons_frame, text="Bỏ chọn tất cả", command=deselect_all, width=15)
        deselect_all_button.pack(side=tk.LEFT, padx=5)
        
        # Tùy chọn định dạng
        format_frame = ttk.LabelFrame(control_frame, text="Tùy chọn định dạng", padding="10")
        format_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Lựa chọn phân cách
        separator_frame = ttk.Frame(format_frame)
        separator_frame.pack(fill=tk.X, pady=5)
        
        separator_label = ttk.Label(separator_frame, text="Phân cách:", width=15)
        separator_label.pack(side=tk.LEFT, padx=5)
        
        separator_var = tk.StringVar(value="tab")
        
        tab_radio = ttk.Radiobutton(separator_frame, text="Tab (cho Excel)", variable=separator_var, value="tab")
        tab_radio.pack(side=tk.LEFT, padx=5)
        
        comma_radio = ttk.Radiobutton(separator_frame, text="Dấu phẩy", variable=separator_var, value="comma")
        comma_radio.pack(side=tk.LEFT, padx=5)
        
        semicolon_radio = ttk.Radiobutton(separator_frame, text="Dấu chấm phẩy", variable=separator_var, value="semicolon")
        semicolon_radio.pack(side=tk.LEFT, padx=5)
        
        # Tùy chọn khác
        options_frame = ttk.Frame(format_frame)
        options_frame.pack(fill=tk.X, pady=5)
        
        include_header_var = tk.BooleanVar(value=True)
        include_header_check = ttk.Checkbutton(options_frame, text="Bao gồm tiêu đề cột", variable=include_header_var)
        include_header_check.pack(side=tk.LEFT, padx=5)
        
        selected_only_var = tk.BooleanVar(value=False)
        selected_only_check = ttk.Checkbutton(options_frame, text="Chỉ dòng đã chọn", variable=selected_only_var)
        selected_only_check.pack(side=tk.LEFT, padx=5)
        
        # Hiển thị xem trước
        preview_frame = ttk.LabelFrame(format_window, text="Xem trước", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        preview_text = tk.Text(preview_frame, wrap=tk.WORD, height=10)
        preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=preview_text.yview)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        preview_text.config(yscrollcommand=preview_scrollbar.set)
        
        def update_preview():
            # Xóa nội dung cũ
            preview_text.delete(1.0, tk.END)
            
            # Lấy các cột đã chọn
            selected_columns = [col for col in self.column_names if column_vars[col].get()]
            
            if not selected_columns:
                preview_text.insert(tk.END, "Chưa chọn cột nào")
                return
            
            # Xác định dữ liệu để xem trước
            preview_data = []
            if selected_only_var.get():
                selected_items = self.data_table.selection()
                if not selected_items:
                    preview_text.insert(tk.END, "Chưa chọn dòng nào")
                    return
                
                for item_id in selected_items[:5]:  # Lấy tối đa 5 dòng để xem trước
                    item_index = self.data_table.index(item_id)
                    if 0 <= item_index < len(self.filtered_data):
                        preview_data.append(self.filtered_data[item_index])
            else:
                preview_data = self.filtered_data[:5]  # Lấy 5 dòng đầu tiên
            
            # Xác định phân cách
            separator = {
                "tab": "\t",
                "comma": ",",
                "semicolon": ";"
            }[separator_var.get()]
            
            # Tạo tiêu đề
            if include_header_var.get():
                headers = [column_display_names.get(col, col) for col in selected_columns]
                preview_text.insert(tk.END, separator.join(headers) + "\n")
            
            # Thêm dữ liệu
            for row in preview_data:
                values = [str(row.get(col, "")) for col in selected_columns]
                preview_text.insert(tk.END, separator.join(values) + "\n")
            
            preview_text.insert(tk.END, "\n(Đây chỉ là bản xem trước với tối đa 5 dòng)")
        
        # Cập nhật xem trước khi thay đổi tùy chọn
        for var in column_vars.values():
            var.trace_add("write", lambda *args: update_preview())
        
        separator_var.trace_add("write", lambda *args: update_preview())
        include_header_var.trace_add("write", lambda *args: update_preview())
        selected_only_var.trace_add("write", lambda *args: update_preview())
        
        # Cập nhật xem trước ban đầu
        update_preview()
        
        # Nút áp dụng và hủy
        button_frame = ttk.Frame(format_window)
        button_frame.pack(fill=tk.X, padx=5, pady=10)
        
        def apply_format():
            # Lấy các cột đã chọn
            selected_columns = [col for col in self.column_names if column_vars[col].get()]
            
            if not selected_columns:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một cột để sao chép")
                return
            
            # Xác định dữ liệu để sao chép
            copy_data = []
            if selected_only_var.get():
                selected_items = self.data_table.selection()
                if not selected_items:
                    messagebox.showwarning("Cảnh báo", "Không có dòng nào được chọn")
                    return
                
                for item_id in selected_items:
                    item_index = self.data_table.index(item_id)
                    if 0 <= item_index < len(self.filtered_data):
                        copy_data.append(self.filtered_data[item_index])
            else:
                copy_data = self.filtered_data
            
            # Xác định phân cách
            separator = {
                "tab": "\t",
                "comma": ",",
                "semicolon": ";"
            }[separator_var.get()]
            
            # Tạo chuỗi để sao chép
            result = ""
            
            # Thêm tiêu đề nếu cần
            if include_header_var.get():
                column_display_names = {
                    'invoice_date': 'Ngày hóa đơn',
                    'invoice_number': 'Số hóa đơn',
                    'invoice_series': 'Ký hiệu',
                    'seller_name': 'Tên người bán',
                    'seller_tax_code': 'MST người bán',
                    'buyer_name': 'Tên người mua',
                    'buyer_tax_code': 'MST người mua',
                    'stt': 'STT',
                    'product_name': 'Tên sản phẩm',
                    'unit': 'Đơn vị tính',
                    'quantity': 'Số lượng',
                    'unit_price': 'Đơn giá',
                    'amount': 'Thành tiền'
                }
                headers = [column_display_names.get(col, col) for col in selected_columns]
                result += separator.join(headers) + "\n"
            
            # Thêm dữ liệu
            for row in copy_data:
                values = [str(row.get(col, "")) for col in selected_columns]
                result += separator.join(values) + "\n"
            
            # Sao chép vào clipboard
            format_window.clipboard_clear()
            format_window.clipboard_append(result)
            
            # Thông báo và đóng cửa sổ
            messagebox.showinfo("Thành công", f"Đã sao chép {len(copy_data)} dòng dữ liệu vào clipboard")
            format_window.destroy()
        
        def cancel_format():
            format_window.destroy()
        
        # Nút áp dụng
        apply_button = ttk.Button(button_frame, text="Sao chép", command=apply_format, width=15)
        apply_button.pack(side=tk.RIGHT, padx=5)
        
        # Nút hủy
        cancel_button = ttk.Button(button_frame, text="Hủy", command=cancel_format, width=15)
        cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # Canh giữa cửa sổ
        format_window.update_idletasks()
        width = format_window.winfo_width()
        height = format_window.winfo_height()
        x = (format_window.winfo_screenwidth() // 2) - (width // 2)
        y = (format_window.winfo_screenheight() // 2) - (height // 2)
        format_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # Đặt focus vào cửa sổ mới
        format_window.transient(self.root)
        format_window.grab_set()
        self.root.wait_window(format_window)

    def find_in_table(self):
        """Tìm kiếm giá trị trong bảng dữ liệu"""
        if not self.filtered_data:
            messagebox.showinfo("Thông báo", "Không có dữ liệu để tìm kiếm")
            return
        
        # Tạo cửa sổ tìm kiếm
        search_window = tk.Toplevel(self.root)
        search_window.title("Tìm kiếm")
        search_window.geometry("400x250")
        search_window.resizable(False, False)
        
        # Tạo các điều khiển
        main_frame = ttk.Frame(search_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Nhập từ khóa tìm kiếm
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=10)
        
        search_label = ttk.Label(search_frame, text="Tìm kiếm:")
        search_label.pack(side=tk.LEFT, padx=5)
        
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        search_entry.focus_set()  # Đặt focus vào ô nhập liệu
        
        # Tùy chọn tìm kiếm
        options_frame = ttk.LabelFrame(main_frame, text="Tùy chọn tìm kiếm", padding="10")
        options_frame.pack(fill=tk.X, pady=10)
        
        # Tìm kiếm trong những cột nào
        columns_frame = ttk.Frame(options_frame)
        columns_frame.pack(fill=tk.X, pady=5)
        
        columns_label = ttk.Label(columns_frame, text="Tìm trong:")
        columns_label.pack(side=tk.LEFT, padx=5)
        
        # Chọn cột
        search_in_var = tk.StringVar(value="all")
        all_radio = ttk.Radiobutton(columns_frame, text="Tất cả các cột", variable=search_in_var, value="all")
        all_radio.pack(side=tk.LEFT, padx=5)
        
        selected_radio = ttk.Radiobutton(columns_frame, text="Cột đã chọn", variable=search_in_var, value="selected")
        selected_radio.pack(side=tk.LEFT, padx=5)
        
        # Tùy chọn khớp
        match_frame = ttk.Frame(options_frame)
        match_frame.pack(fill=tk.X, pady=5)
        
        exact_match_var = tk.BooleanVar(value=False)
        exact_match_check = ttk.Checkbutton(match_frame, text="Khớp chính xác", variable=exact_match_var)
        exact_match_check.pack(side=tk.LEFT, padx=5)
        
        case_sensitive_var = tk.BooleanVar(value=False)
        case_sensitive_check = ttk.Checkbutton(match_frame, text="Phân biệt hoa thường", variable=case_sensitive_var)
        case_sensitive_check.pack(side=tk.LEFT, padx=5)
        
        # Kết quả tìm kiếm
        result_frame = ttk.LabelFrame(main_frame, text="Kết quả", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        result_var = tk.StringVar(value="Nhập từ khóa và nhấn Tìm kiếm")
        result_label = ttk.Label(result_frame, textvariable=result_var, wraplength=350)
        result_label.pack(fill=tk.BOTH, expand=True)
        
        # Các nút thực thi
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        def search():
            keyword = search_var.get().strip()
            if not keyword:
                result_var.set("Vui lòng nhập từ khóa tìm kiếm")
                return
            
            # Xác định phạm vi tìm kiếm
            search_columns = []
            if search_in_var.get() == "all":
                search_columns = self.column_names
            elif search_in_var.get() == "selected":
                if not hasattr(self, 'selected_column') or not self.selected_column:
                    result_var.set("Vui lòng chọn một cột trong bảng dữ liệu trước")
                    return
                
                column_index = int(self.selected_column.replace('#', '')) - 1
                if column_index < 0 or column_index >= len(self.column_names):
                    result_var.set("Không thể xác định cột đã chọn")
                    return
                
                search_columns = [self.column_names[column_index]]
            
            # Thực hiện tìm kiếm
            matches = []
            match_indices = []
            
            for i, row in enumerate(self.filtered_data):
                for col in search_columns:
                    value = str(row.get(col, ""))
                    
                    # Xử lý điều kiện phân biệt hoa thường
                    if not case_sensitive_var.get():
                        value = value.lower()
                        search_keyword = keyword.lower()
                    else:
                        search_keyword = keyword
                    
                    # Kiểm tra khớp
                    if exact_match_var.get():
                        if value == search_keyword:
                            matches.append(row)
                            match_indices.append(i)
                            break
                    else:
                        if search_keyword in value:
                            matches.append(row)
                            match_indices.append(i)
                            break
            
            # Hiển thị kết quả
            if not matches:
                result_var.set(f"Không tìm thấy kết quả nào khớp với từ khóa '{keyword}'")
            else:
                result_var.set(f"Tìm thấy {len(matches)} kết quả khớp với từ khóa '{keyword}'")
                
                # Bỏ chọn tất cả trước
                self.data_table.selection_remove(self.data_table.selection())
                
                # Chọn các dòng khớp
                for i in match_indices:
                    try:
                        item_id = self.data_table.get_children()[i]
                        self.data_table.selection_add(item_id)
                        self.data_table.see(item_id)  # Cuộn đến dòng đầu tiên khớp
                    except IndexError:
                        pass
                
                # Cập nhật thông tin
                self.update_data_info()
                
                # Đóng cửa sổ tìm kiếm nếu có kết quả
                search_window.after(1000, search_window.destroy)
        
        def cancel():
            search_window.destroy()
        
        # Nút tìm kiếm
        search_button = ttk.Button(button_frame, text="Tìm kiếm", command=search, width=15)
        search_button.pack(side=tk.RIGHT, padx=5)
        
        # Nút hủy
        cancel_button = ttk.Button(button_frame, text="Đóng", command=cancel, width=15)
        cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # Xử lý phím Enter để tìm kiếm
        search_window.bind("<Return>", lambda event: search())
        
        # Canh giữa cửa sổ
        search_window.update_idletasks()
        width = search_window.winfo_width()
        height = search_window.winfo_height()
        x = (search_window.winfo_screenwidth() // 2) - (width // 2)
        y = (search_window.winfo_screenheight() // 2) - (height // 2)
        search_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # Đặt focus vào cửa sổ mới
        search_window.transient(self.root)
        search_window.grab_set()
        self.root.wait_window(search_window)
    
    def show_column_statistics(self):
        """Hiển thị thống kê cơ bản cho cột số liệu được chọn"""
        if not self.filtered_data:
            messagebox.showinfo("Thông báo", "Không có dữ liệu để phân tích")
            return
        
        # Kiểm tra cột đã chọn
        if not hasattr(self, 'selected_column') or not self.selected_column:
            messagebox.showinfo("Thông báo", "Vui lòng chọn một cột số liệu để hiển thị thống kê")
            return
        
        column_index = int(self.selected_column.replace('#', '')) - 1
        if column_index < 0 or column_index >= len(self.column_names):
            return
        
        selected_column_name = self.column_names[column_index]
        column_display_name = {
            'invoice_date': 'Ngày hóa đơn',
            'invoice_number': 'Số hóa đơn',
            'invoice_series': 'Ký hiệu',
            'seller_name': 'Tên người bán',
            'seller_tax_code': 'MST người bán',
            'buyer_name': 'Tên người mua',
            'buyer_tax_code': 'MST người mua',
            'stt': 'STT',
            'product_name': 'Tên sản phẩm',
            'unit': 'Đơn vị tính',
            'quantity': 'Số lượng',
            'unit_price': 'Đơn giá',
            'amount': 'Thành tiền'
        }.get(selected_column_name, selected_column_name)
        
        # Kiểm tra nếu cột là loại số liệu
        numeric_columns = ['quantity', 'unit_price', 'amount']
        
        if selected_column_name not in numeric_columns:
            messagebox.showinfo("Thông báo", f"Cột '{column_display_name}' không phải là cột số liệu.\nVui lòng chọn một trong các cột: Số lượng, Đơn giá, Thành tiền.")
            return
        
        # Thu thập dữ liệu từ cột đã chọn
        values = []
        for row in self.filtered_data:
            try:
                # Lấy giá trị và chuyển đổi sang số
                value_str = str(row.get(selected_column_name, "0")).replace(",", ".")
                value = float(value_str)
                values.append(value)
            except ValueError:
                # Bỏ qua giá trị không phải số
                pass
        
        if not values:
            messagebox.showinfo("Thông báo", f"Không có dữ liệu số trong cột '{column_display_name}'")
            return
        
        # Tính toán thống kê
        total = sum(values)
        count = len(values)
        average = total / count if count > 0 else 0
        minimum = min(values) if values else 0
        maximum = max(values) if values else 0
        
        # Hiển thị thống kê
        stats_window = tk.Toplevel(self.root)
        stats_window.title(f"Thống kê cột {column_display_name}")
        stats_window.geometry("400x300")
        stats_window.resizable(False, False)
        
        main_frame = ttk.Frame(stats_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tiêu đề
        title_label = ttk.Label(main_frame, text=f"Thống kê cột: {column_display_name}", font=("Arial", 12, "bold"))
        title_label.pack(pady=10)
        
        # Khung thống kê
        stats_frame = ttk.LabelFrame(main_frame, text="Các chỉ số thống kê", padding="10")
        stats_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Định dạng số
        def format_number(value):
            if value == int(value):
                return f"{int(value):,}"
            return f"{value:,.2f}"
        
        # Hiển thị từng chỉ số
        stats_data = [
            ("Tổng số dòng dữ liệu:", f"{count:,}"),
            ("Tổng cộng:", format_number(total)),
            ("Trung bình:", format_number(average)),
            ("Giá trị nhỏ nhất:", format_number(minimum)),
            ("Giá trị lớn nhất:", format_number(maximum))
        ]
        
        for i, (label_text, value_text) in enumerate(stats_data):
            row_frame = ttk.Frame(stats_frame)
            row_frame.pack(fill=tk.X, pady=5)
            
            label = ttk.Label(row_frame, text=label_text, width=20)
            label.pack(side=tk.LEFT, padx=5)
            
            value = ttk.Label(row_frame, text=value_text)
            value.pack(side=tk.LEFT, padx=5)
        
        # Nút Đóng
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        close_button = ttk.Button(button_frame, text="Đóng", command=stats_window.destroy, width=15)
        close_button.pack(side=tk.RIGHT, padx=5)
        
        # Nút sao chép
        def copy_stats():
            stats_text = f"Thống kê cột: {column_display_name}\n\n"
            for label, value in stats_data:
                stats_text += f"{label} {value}\n"
            
            stats_window.clipboard_clear()
            stats_window.clipboard_append(stats_text)
            messagebox.showinfo("Sao chép", "Đã sao chép thống kê vào clipboard")
        
        copy_button = ttk.Button(button_frame, text="Sao chép", command=copy_stats, width=15)
        copy_button.pack(side=tk.RIGHT, padx=5)
        
        # Canh giữa cửa sổ
        stats_window.update_idletasks()
        width = stats_window.winfo_width()
        height = stats_window.winfo_height()
        x = (stats_window.winfo_screenwidth() // 2) - (width // 2)
        y = (stats_window.winfo_screenheight() // 2) - (height // 2)
        stats_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # Đặt focus vào cửa sổ mới
        stats_window.transient(self.root)
        stats_window.grab_set()
        self.root.wait_window(stats_window)
    
    def update_context_menu(self):
        """Cập nhật menu ngữ cảnh với các tùy chọn nâng cao"""
        # Xóa menu cũ
        self.context_menu.delete(0, tk.END)
        
        # Tùy chọn sao chép
        self.context_menu.add_command(label="Sao chép dòng", command=self.copy_row)
        self.context_menu.add_command(label="Sao chép giá trị ô", command=self.copy_cell)
        self.context_menu.add_command(label="Sao chép tất cả các dòng đã chọn", command=lambda: self.copy_rows_to_clipboard(
            [self.filtered_data[self.data_table.index(item_id)] for item_id in self.data_table.selection() 
             if 0 <= self.data_table.index(item_id) < len(self.filtered_data)]
        ))
        
        self.context_menu.add_separator()
        
        # Tùy chọn xuất
        self.context_menu.add_command(label="Xuất dòng đã chọn...", command=self.export_selected_rows)
        
        # Thêm tùy chọn định dạng tùy chỉnh
        self.context_menu.add_command(label="Sao chép với định dạng tùy chỉnh...", command=self.copy_with_custom_format)
        
        self.context_menu.add_separator()
        
        # Tùy chọn lọc và tìm kiếm
        if hasattr(self, 'selected_column') and self.selected_column:
            column_index = int(self.selected_column.replace('#', '')) - 1
            if 0 <= column_index < len(self.column_names):
                column_name = self.column_names[column_index]
                display_name = {
                    'invoice_date': 'Ngày hóa đơn',
                    'invoice_number': 'Số hóa đơn',
                    'invoice_series': 'Ký hiệu',
                    'seller_name': 'Tên người bán',
                    'seller_tax_code': 'MST người bán',
                    'buyer_name': 'Tên người mua',
                    'buyer_tax_code': 'MST người mua',
                    'stt': 'STT',
                    'product_name': 'Tên sản phẩm',
                    'unit': 'Đơn vị tính',
                    'quantity': 'Số lượng',
                    'unit_price': 'Đơn giá',
                    'amount': 'Thành tiền'
                }.get(column_name, column_name)
                
                self.context_menu.add_command(label=f"Lọc theo giá trị ô này", command=self.filter_by_cell_value)
                
                # Thêm phân tích thống kê cho cột số liệu
                if column_name in ['quantity', 'unit_price', 'amount']:
                    self.context_menu.add_command(label=f"Thống kê cột {display_name}", command=self.show_column_statistics)
        
        self.context_menu.add_command(label="Tìm kiếm...", command=self.find_in_table)
        
        self.context_menu.add_separator()
        
        # Tùy chọn chọn
        self.context_menu.add_command(label="Chọn tất cả", command=self.select_all_rows)
        self.context_menu.add_command(label="Bỏ chọn tất cả", command=lambda: self.data_table.selection_remove(
            self.data_table.selection()
        ))
        
        self.context_menu.add_separator()
        
        # Tùy chọn khác
        self.context_menu.add_command(label="Sao chép tất cả dữ liệu", command=self.copy_all_data)
    
    def show_context_menu(self, event):
        """Hiển thị menu ngữ cảnh khi nhấp chuột phải"""
        # Lấy dòng được chọn
        selected_iid = self.data_table.identify_row(event.y)
        if selected_iid:
            # Đánh dấu dòng được chọn
            self.data_table.selection_set(selected_iid)
            # Lưu vị trí chuột
            self.context_menu_position = (event.x_root, event.y_root)
            # Lưu thông tin ô được chọn
            self.selected_column = self.data_table.identify_column(event.x)
            # Cập nhật menu ngữ cảnh
            self.update_context_menu()
            # Hiển thị menu ngữ cảnh
            self.context_menu.post(event.x_root, event.y_root)
            
# Hàm main để khởi động ứng dụng
def main():
    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
