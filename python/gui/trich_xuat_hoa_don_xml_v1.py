import os
import xml.etree.ElementTree as ET
import csv
import glob
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import pandas as pd
from tkcalendar import DateEntry  # Thêm thư viện calendar để chọn ngày
# lọc hoá đơn theo thời gian, không có copy
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
        
        # Kích hoạt các nút xuất dữ liệu và lọc dữ liệu
        self.export_csv_button.config(state=tk.NORMAL)
        self.export_xlsx_button.config(state=tk.NORMAL)
        self.apply_filter_button.config(state=tk.NORMAL)
        self.show_all_button.config(state=tk.NORMAL)
        
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

# Hàm main để khởi động ứng dụng
def main():
    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
