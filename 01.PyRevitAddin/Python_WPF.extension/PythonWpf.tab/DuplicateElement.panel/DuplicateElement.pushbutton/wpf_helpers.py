# -*- coding: utf-8 -*-
import clr

# Add necessary .NET references
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference("System.Windows.Forms")
clr.AddReference('System.Data')

from System.Windows import Visibility, Thickness
from System.Windows.Media import Brushes, SolidColorBrush, Colors
from System.Windows.Data import IValueConverter, Binding, RelativeSource, RelativeSourceMode
from System.Globalization import CultureInfo
from System.Data import DataView, DataRowView

# ================================================================
# 1️⃣ CONVERTERS
# ================================================================

class StatusToVisibilityConverter(IValueConverter):
    """Hiển thị hoặc ẩn button dựa theo Status"""
    def Convert(self, value, target_type, parameter, culture):
        if value is None or parameter is None:
            return Visibility.Collapsed

        status = str(value).lower()  # Convert to lowercase for case-insensitive comparison
        target_status = str(parameter).lower()  # Convert to lowercase for case-insensitive comparison

        # Cho phép nhiều điều kiện 'New|Update'
        if "|" in target_status:
            target_statuses = [s.strip() for s in target_status.split("|")]  # Split and trim values
            if status in target_statuses:
                return Visibility.Visible
        elif status == target_status:
            return Visibility.Visible

        return Visibility.Collapsed

    def ConvertBack(self, value, target_type, parameter, culture):
        return None


class StatusToColorConverter(IValueConverter):
    """Chuyển Status sang màu nền button"""
    def Convert(self, value, target_type, parameter, culture):
        if value is None:
            return Brushes.LightGray

        status = str(value).lower()
        if status == "update":
            return Brushes.Gold
        elif status == "new":
            return Brushes.LightGreen
        elif status == "existing":
            return Brushes.LightBlue
        else:
            return Brushes.LightGray

    def ConvertBack(self, value, target_type, parameter, culture):
        return None


class StatusToActionTextConverter(IValueConverter):
    """Chuyển Status sang text hiển thị trên button"""
    def Convert(self, value, target_type, parameter, culture):
        if value is None:
            return "Action"

        status = str(value).lower()
        if status == "update":
            return "Update"
        elif status == "new":
            return "Create"
        elif status == "existing":
            return "Select"
        else:
            return "Action"

    def ConvertBack(self, value, target_type, parameter, culture):
        return None


# ================================================================
# 2️⃣ DATAGRID SETUP
# ================================================================

def setup_datagrid_columns(datagrid, data_table, view_model):
    """Tạo các cột từ DataTable, thêm cột Status (và Action nếu cần)"""
    if not datagrid or not data_table:
        return

    from System.Windows.Controls import DataGridTextColumn, DataGridTemplateColumn, StackPanel, TextBlock, Button
    from System.Windows import FrameworkElementFactory, DataTemplate
    from System.Windows.Controls import Orientation

    # Xóa cột cũ
    datagrid.Columns.Clear()

    # 1️⃣ Tạo các cột từ Excel
    for i in range(data_table.Columns.Count):
        col_name = data_table.Columns[i].ColumnName
        if col_name not in ["Status", "Action"]:
            col = DataGridTextColumn()
            col.Header = col_name
            col.Binding = Binding(col_name)
            col.Width = 200 if i == 0 else 120
            datagrid.Columns.Add(col)

    # 2️⃣ Tạo cột Status (hiển thị text)
    if "Status" in [c.ColumnName for c in data_table.Columns]:
        status_col = DataGridTextColumn()
        status_col.Header = "Status"
        status_col.Binding = Binding("Status")
        status_col.Width = 100
        datagrid.Columns.Add(status_col)

    # 3️⃣ Tạo cột Action (có button)
    action_col = DataGridTemplateColumn()
    action_col.Header = "Action"
    action_col.Width = 100

    # Template button
    template = DataTemplate()
    factory = FrameworkElementFactory(Button)
    factory.SetValue(Button.HeightProperty, 25.0)
    factory.SetValue(Button.WidthProperty, 80.0)
    factory.SetValue(Button.MarginProperty, Thickness(3))
    factory.SetBinding(Button.ContentProperty, Binding("Status", Converter=StatusToActionTextConverter()))
    factory.SetBinding(Button.BackgroundProperty, Binding("Status", Converter=StatusToColorConverter()))

    # Gán command nếu ViewModel có lệnh ActionCommand
    cmd_binding = Binding("DataContext.ActionCommand")
    cmd_binding.RelativeSource = RelativeSource(RelativeSourceMode.FindAncestor, datagrid.GetType(), 1)
    factory.SetBinding(Button.CommandProperty, cmd_binding)
    factory.SetBinding(Button.CommandParameterProperty, Binding())

    template.VisualTree = factory
    action_col.CellTemplate = template
    datagrid.Columns.Insert(0, action_col)

    datagrid.UpdateLayout()


# ================================================================
# 3️⃣ HÀM ĐỊNH DẠNG MÀU HÀNG
# ================================================================

def apply_row_formatting(datagrid):
    """Tô màu hàng dựa vào giá trị cột Status"""
    if not datagrid:
        return

    existing_brush = SolidColorBrush(Colors.LightGreen)
    update_brush = SolidColorBrush(Colors.LightYellow)
    new_brush = SolidColorBrush(Colors.LightBlue)

    def update_row_style(sender, args):
        if args.Row.Item is None:
            return
        try:
            status = str(args.Row.Item["Status"]).lower()
            if status == "existing":
                args.Row.Background = existing_brush
            elif status == "update":
                args.Row.Background = update_brush
            elif status == "new":
                args.Row.Background = new_brush
        except:
            pass

    try:
        datagrid.LoadingRow -= update_row_style
    except:
        pass
    datagrid.LoadingRow += update_row_style


# ================================================================
# 4️⃣ HÀM CẬP NHẬT UI THEO HÀNG ĐƯỢC CHỌN
# ================================================================

def update_ui_based_on_selection(window, selected_item, categorized_types):
    """Cập nhật UI dựa vào hàng được chọn và loại dữ liệu"""
    try:
        # Ẩn các nút action mặc định
        if hasattr(window, 'btnUpdateSelected'):
            window.btnUpdateSelected.Visibility = Visibility.Collapsed
        if hasattr(window, 'btnCreateSelected'):
            window.btnCreateSelected.Visibility = Visibility.Collapsed
            
        # Nếu không có hàng nào được chọn, không làm gì thêm
        if not selected_item:
            return
            
        # Lấy status từ hàng được chọn
        try:
            row = selected_item.Row
            status = str(row["Status"]).lower()
            
            # Hiển thị nút tương ứng với status
            if status == "update":
                if hasattr(window, 'btnUpdateSelected'):
                    window.btnUpdateSelected.Visibility = Visibility.Visible
            elif status == "new":
                if hasattr(window, 'btnCreateSelected'):
                    window.btnCreateSelected.Visibility = Visibility.Visible
        except Exception as ex:
            print("Error getting row status: {}".format(str(ex)))
            
        # Cập nhật thông tin tổng quan nếu có categorized_types
        if categorized_types:
            try:
                if hasattr(window, 'btnExisting'):
                    window.btnExisting.Content = "Existing: {}".format(len(categorized_types.get('existing', [])))
                if hasattr(window, 'btnUpdate'):
                    window.btnUpdate.Content = "Need Update: {}".format(len(categorized_types.get('update', [])))
                if hasattr(window, 'btnNew'):
                    window.btnNew.Content = "New: {}".format(len(categorized_types.get('new', [])))
                
                # Hiển thị nút "Create All" nếu có các loại mới
                if hasattr(window, 'btnCreateAll'):
                    if len(categorized_types.get('new', [])) > 0:
                        window.btnCreateAll.Visibility = Visibility.Visible
                    else:
                        window.btnCreateAll.Visibility = Visibility.Collapsed
            except Exception as ex:
                print("Error updating type counts: {}".format(str(ex)))
    except Exception as ex:
        print("Error updating UI based on selection: {}".format(str(ex)))

# ================================================================
# 5️⃣ HÀM THÊM BUTTON CỘT ACTION
# ================================================================

def replace_action_column(datagrid, view_model):
    """Tạo lại cột Action có chứa button test"""
    try:
        from System.Windows.Controls import DataGridTemplateColumn, Button, DataGrid
        from System.Windows import DataTemplate, FrameworkElementFactory, Thickness
        from System.Windows.Data import Binding, RelativeSource, RelativeSourceMode

        # Xóa cột cũ
        for i in range(datagrid.Columns.Count - 1, -1, -1):
            if str(datagrid.Columns[i].Header) == "Action":
                datagrid.Columns.RemoveAt(i)

        # Tạo template button
        template = DataTemplate()
        factory = FrameworkElementFactory(Button)
        factory.SetValue(Button.ContentProperty, "Run")
        factory.SetValue(Button.WidthProperty, 80.0)
        factory.SetValue(Button.HeightProperty, 25.0)
        factory.SetValue(Button.BackgroundProperty, Brushes.LightBlue)
        factory.SetValue(Button.MarginProperty, Thickness(3))

        # Bind Command to VM.ActionCommand and pass current row as CommandParameter
        cmd_binding = Binding("DataContext.ActionCommand")
        cmd_binding.RelativeSource = RelativeSource(RelativeSourceMode.FindAncestor, DataGrid, 1)
        factory.SetBinding(Button.CommandProperty, cmd_binding)
        factory.SetBinding(Button.CommandParameterProperty, Binding())  # pass DataRowView

        template.VisualTree = factory

        # Tạo cột Action và chèn vào vị trí đầu
        col = DataGridTemplateColumn()
        col.Header = "Action"
        col.CellTemplate = template
        datagrid.Columns.Insert(0, col)

        datagrid.UpdateLayout()
        print("✅ Action column added successfully at index 0")
        return True

    except Exception as ex:
        import traceback
        print("❌ replace_action_column ERROR:", ex)
        print(traceback.format_exc())
        return False

# ================================================================
# 6️⃣ HÀM ĐĂNG KÝ BUTTON EVENTS
# ================================================================

def register_button_events(datagrid, click_handler):
    """Đăng ký sự kiện click cho các buttons trong DataGrid"""
    try:
        print("Đang đăng ký sự kiện cho buttons...")
        if not datagrid or not click_handler:
            print("DataGrid hoặc handler không hợp lệ")
            return 0

        from System.Windows.Controls import Button
        from System.Windows.Media import VisualTreeHelper

        def find_buttons(parent):
            found = []
            if parent is None:
                return found
            # If the visual is a Button
            if isinstance(parent, Button):
                found.append(parent)
                return found
            # Traverse children
            try:
                count = VisualTreeHelper.GetChildrenCount(parent)
                for i in range(count):
                    child = VisualTreeHelper.GetChild(parent, i)
                    found.extend(find_buttons(child))
            except:
                pass
            return found

        button_count = 0
        datagrid.UpdateLayout()
        for row_idx in range(datagrid.Items.Count):
            try:
                container = datagrid.ItemContainerGenerator.ContainerFromIndex(row_idx)
                if not container:
                    continue
                for col_idx in range(datagrid.Columns.Count):
                    col = datagrid.Columns[col_idx]
                    if str(col.Header) != "Action":
                        continue
                    cell_content = col.GetCellContent(container)
                    if cell_content is None:
                        continue
                    buttons = find_buttons(cell_content)
                    for btn in buttons:
                        # Only attach fallback Click if no Command is set
                        try:
                            if getattr(btn, "Command", None) is None:
                                btn.Click += click_handler
                                button_count += 1
                                print("✓ Đăng ký button tại dòng {}".format(row_idx))
                        except:
                            # Safe-attach anyway
                            btn.Click += click_handler
                            button_count += 1
                            print("✓ Đăng ký button (force) tại dòng {}".format(row_idx))
            except Exception as ex:
                print("Lỗi khi xử lý dòng {}: {}".format(row_idx, str(ex)))

        print("Đã đăng ký tổng cộng {} buttons".format(button_count))
        return button_count
    except Exception as ex:
        print("Lỗi đăng ký sự kiện: {}".format(str(ex)))
        import traceback
        print(traceback.format_exc())
        return 0