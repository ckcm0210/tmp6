# dependency_converter.py

import os
import re
import colorsys
import urllib.parse

def _format_formula_for_display(formula, max_line_length=50):
    """
    將長公式換行以便在 pyvis 節點中更好地顯示。
    只在有意義的位置斷行，保持可讀性。
    """
    if not formula or formula == 'N/A':
        return 'N/A'
    
    # 移除開頭的等號（如果有的話）
    display_formula = formula[1:] if formula.startswith('=') else formula
    
    if len(display_formula) <= max_line_length:
        return display_formula
    
    # 只在有意義的位置斷行：加減乘除運算符
    break_after = ['+', '-', '*', '/', ',']
    
    formatted_formula = ''
    current_line = ''
    
    for char in display_formula:
        current_line += char
        # 只有在行長度超過限制且遇到運算符時才斷行
        if len(current_line) >= max_line_length and char in break_after:
            formatted_formula += current_line + '\n'
            current_line = ''
    
    # 添加剩餘的字符
    if current_line:
        formatted_formula += current_line
    
    return formatted_formula

def _create_short_formula(formula):
    """
    創建簡化的公式顯示，隱藏完整路徑，只保留檔案名和工作表
    例如：='C:\path\[file.xlsx]Sheet1'!A1 -> =[file.xlsx]Sheet1!A1
    """
    if not formula or formula == 'N/A':
        return 'N/A'
    
    # 移除開頭的等號（如果有的話）
    display_formula = formula[1:] if formula.startswith('=') else formula
    
    # 使用正則表達式找到並簡化路徑
    import re
    
    # 匹配完整路徑格式：'C:\path\[file.xlsx]Sheet'!A1
    pattern = r"'([^']*\\)?\[([^\]]+)\]([^']*)'!"
    
    def replace_path(match):
        # 只保留檔案名和工作表名
        filename = match.group(2)  # [file.xlsx] 中的 file.xlsx
        sheet = match.group(3)     # 工作表名
        return f"[{filename}]{{sheet}}!"
    
    # 替換所有匹配的路徑
    simplified_formula = re.sub(pattern, replace_path, display_formula)
    
    # 格式化顯示（使用較短的行長度）
    return _format_formula_for_display('=' + simplified_formula, max_line_length=40)

def _generate_unique_colors_for_files(filenames):
    """確保每個檔案都有唯一顏色，避免顏色衝突"""
    import colorsys
    
    # 擴展的顏色調色板
    color_palette = [
        "#007bff", "#28a745", "#ff8c00", "#dc3545", "#6f42c1", "#20c997",
        "#fd7e14", "#e83e8c", "#6610f2", "#17a2b8", "#ffc107", "#198754",
        "#0d6efd", "#d63384", "#adb5bd", "#495057", "#f8f9fa", "#343a40"
    ]
    
    # 如果檔案數量超過調色板，生成額外顏色
    if len(filenames) > len(color_palette):
        for i in range(len(color_palette), len(filenames)):
            hue = (i * 137.508) % 360  # 使用黃金角度確保顏色分散
            saturation = 0.7 + (i % 3) * 0.1
            value = 0.8 + (i % 2) * 0.1
            
            rgb = colorsys.hsv_to_rgb(hue/360, saturation, value)
            hex_color = "#{:02x}{:02x}{:02x}".format(
                int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255)
            )
            color_palette.append(hex_color)
    
    # 分配唯一顏色
    file_colors = {}
    unique_filenames = list(set(filenames))
    
    # 確保 Current File 使用第一個顏色（藍色）
    if 'Current File' in unique_filenames:
        file_colors['Current File'] = color_palette[0]
        unique_filenames.remove('Current File')
        remaining_colors = color_palette[1:]
    else:
        remaining_colors = color_palette
    
    # 為其他檔案分配顏色
    for i, filename in enumerate(sorted(unique_filenames)):
        file_colors[filename] = remaining_colors[i % len(remaining_colors)]
    
    return file_colors

def _create_short_address(address):
    """創建簡化的地址顯示，避免工作表名重複"""
    if not address or address == 'N/A':
        return 'N/A'
    
    # 如果是外部引用，只顯示檔案名和工作表
    if '[' in address and ']' in address:
        # 解析格式：[filename]sheet!cell
        bracket_end = address.find(']')
        if bracket_end != -1:
            filename_part = address[1:bracket_end]  # 提取檔案名（去掉 [ ]）
            remaining_part = address[bracket_end + 1:]  # 工作表!儲存格
            return f"[{filename_part}]{{remaining_part}}"
    
    return address

def _create_enhanced_node_label(address, formula, value, node_type):
    """創建增強的節點標籤，對齊冒號，長公式換行對齊"""
    short_address = _create_short_address(address)
    short_formula = _create_short_formula(formula) if formula != 'N/A' else ''
    formatted_value = _format_value_display(value)
    
    # 三層結構，冒號對齊，長公式換行對齊
    parts = []
    
    # 第一行：地址（8個字符對齊）
    parts.append(f"Address : {short_address}")
    
    # 第二行：公式（如果有），處理長公式換行對齊
    if short_formula and short_formula != 'N/A':
        formatted_formula = _format_long_formula_with_alignment(short_formula)
        parts.append(f"Formula : {formatted_formula}")
    
    # 第三行：值（8個字符對齊）
    parts.append(f"Value   : {formatted_value}")
    
    return "\n".join(parts)

def _format_long_formula_with_alignment(formula, max_length=35):
    """格式化長公式，換行時保持對齊"""
    if not formula or len(formula) <= max_length:
        return formula
    
    # 找到適當的斷行位置（運算符後）
    break_chars = ['+', '-', '*', '/', ',', '(', ')']
    
    lines = []
    current_line = ''
    
    i = 0
    while i < len(formula):
        char = formula[i]
        current_line += char
        
        # 如果當前行長度超過限制，尋找斷行點
        if len(current_line) >= max_length:
            # 向後尋找最近的斷行字符
            break_pos = -1
            for j in range(len(current_line) - 1, max(0, len(current_line) - 10), -1):
                if current_line[j] in break_chars:
                    break_pos = j
                    break
            
            if break_pos > 0:
                # 在斷行字符後斷行
                lines.append(current_line[:break_pos + 1])
                current_line = '          ' + current_line[break_pos + 1:]  # 10個空格對齊
            else:
                # 找不到合適斷行點，強制斷行
                lines.append(current_line)
                current_line = '          '  # 10個空格對齊
        
        i += 1
    
    # 添加剩餘部分
    if current_line.strip():
        lines.append(current_line)
    
    return '\n'.join(lines)

def _format_value_display(value):
    """格式化值的顯示，支援hash值自動換行和Error處理"""
    if value is None or value == 'N/A':
        return 'N/A'
    
    try:
        if isinstance(value, (int, float)):
            if isinstance(value, float) and value.is_integer():
                return f"{int(value):,}"
            else:
                return f"{value:,.2f}"
    except:
        pass
    
    str_value = str(value)
    
    # === 修復問題2：處理Error值，嘗試提供更多信息 ===
    if str_value == 'Error':
        return 'Error (計算失敗)'
    
    # === 新增：檢查是否為hash值格式並自動換行 ===
    # 檢查是否包含hash值模式 (如: "9Rx1C | Hash: abc123def456")
    if 'Hash:' in str_value and '|' in str_value:
        parts = str_value.split('|')
        if len(parts) >= 2:
            dimension_part = parts[0].strip()
            hash_part = parts[1].strip()
            
            # 如果hash值很長，進行換行
            if len(hash_part) > 20:
                # 將hash值分成多行，每行最多16個字符
                hash_value = hash_part.replace('Hash: ', '')
                formatted_hash_lines = []
                for i in range(0, len(hash_value), 16):
                    formatted_hash_lines.append(hash_value[i:i+16])
                
                # 重新組合，第一行包含"Hash:"，後續行縮進對齊
                result = dimension_part + '\n'
                result += 'Hash: ' + formatted_hash_lines[0]
                for line in formatted_hash_lines[1:]:
                    result += '\n      ' + line  # 6個空格對齊"Hash: "
                return result
    
    # 一般值的處理
    if len(str_value) > 25:
        return f"{str_value[:22]}..."
    
    return str_value

def _create_enhanced_tooltip(node_data):
    """創建簡潔的純文字 tooltip，避免 HTML 顯示問題"""
    address = node_data.get('address', 'N/A')
    formula = node_data.get('formula', 'N/A')
    value = node_data.get('value', 'N/A')
    node_type = node_data.get('type', 'unknown')
    filename = node_data.get('filename', 'Unknown')
    
    tooltip_parts = []
    
    # 地址部分
    tooltip_parts.append(f"Address: {address}")
    tooltip_parts.append("***")
    
    # 類型部分
    tooltip_parts.append(f"Type: {node_type.title()}")
    tooltip_parts.append("***")
    
    # 公式部分（如果有）
    if formula and formula != 'N/A' and formula.strip():
        tooltip_parts.append(f"Formula: {formula}")
        tooltip_parts.append("***")
    
    # 值部分
    formatted_value = _format_value_display(value)
    tooltip_parts.append(f"Value: {formatted_value}")
    tooltip_parts.append("***")
    
    # 檔案信息
    tooltip_parts.append(f"File: {filename}")
    
    return "\n".join(tooltip_parts)

def _format_formula_for_tooltip(formula):
    """格式化公式以便在 tooltip 中顯示，在適當位置換行"""
    if not formula or len(formula) <= 60:
        return formula
    
    # 在適當位置添加換行，但保持外部引用的完整性
    import re
    
    # 保護外部引用
    external_refs = re.findall(r"'[^']*'![A-Z]+\d+", formula)
    temp_formula = formula
    placeholders = {}
    
    for i, ref in enumerate(external_refs):
        placeholder = f"__EXT_REF_{i}__"
        placeholders[placeholder] = ref
        temp_formula = temp_formula.replace(ref, placeholder)
    
    # 在運算符處換行
    formatted = re.sub(r'([+\-*/,])', r'\1<br>&nbsp;&nbsp;', temp_formula)
    
    # 恢復外部引用
    for placeholder, original in placeholders.items():
        formatted = formatted.replace(placeholder, original)
    
    return formatted

def convert_tree_to_graph_data(dependency_tree_data):
    """
    將從 explode_cell_dependencies 得到的樹狀資料，轉換為 pyvis 需要的格式。
    改進版本：每個主節點左邊添加小的標籤節點，避免對齊問題
    """
    import colorsys
    
    nodes_data = []
    edges_data = []
    processed_nodes = set()
    
    # 首先收集所有檔案名稱以生成唯一顏色
    all_filenames = set()
    
    def collect_filenames(node):
        address = node.get('address', '')
        workbook_path = node.get('workbook_path', '')
        
        if '[' in address and ']' in address:
            # 外部引用
            match = re.search(r'\[([^\]]+)\]', address)
            if match:
                all_filenames.add(match.group(1))
        elif workbook_path:
            # 本地引用：從workbook_path提取檔案名
            filename = os.path.basename(workbook_path)
            if filename.endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                filename = os.path.splitext(filename)[0] + '.xlsx'
            all_filenames.add(filename)
        else:
            # 備用
            all_filenames.add('Current File')
        
        for child in node.get('children', []):
            collect_filenames(child)
    
    collect_filenames(dependency_tree_data)
    
    # 生成唯一顏色映射
    file_colors = _generate_unique_colors_for_files(list(all_filenames))

    def traverse_tree(node, parent_id=None):
        node_id = node.get('address', str(hash(str(node))))
        
        if node_id not in processed_nodes:
            processed_nodes.add(node_id)
            
            # 提取基本信息
            address = node.get('address', 'N/A')
            raw_formula = node.get('formula', 'N/A')
            resolved_formula = node.get('resolved_formula', '')
            value = node.get('value', 'N/A')
            node_type = node.get('type', 'unknown')
            has_resolved = node.get('has_resolved', False)
            
            # === 修復：正確確定檔案名和顏色 + 清理URL編碼 ===
            filename = 'Current File'
            workbook_path = node.get('workbook_path', '')
            
            import urllib.parse
            clean_address = urllib.parse.unquote(address) if address else address
            
            if '[' in clean_address and ']' in clean_address:
                match = re.search(r'\[([^\]]+)\]', clean_address)
                if match:
                    filename = urllib.parse.unquote(match.group(1))
            elif workbook_path:
                filename = os.path.basename(urllib.parse.unquote(workbook_path))
                if filename.endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                    filename = os.path.splitext(filename)[0] + '.xlsx'
            else:
                filename = 'Current File'
            
            color = file_colors.get(filename, "#808080")
            
            short_address = node.get('short_address', _create_short_address(address))
            full_address = node.get('full_address', address)
            short_formula = _create_short_formula(raw_formula)
            full_formula = raw_formula
            formatted_value = _format_value_display(value)
            
            if short_formula and short_formula != 'N/A':
                display_formula = short_formula if short_formula.startswith('=') else f"={short_formula}"
                simple_label = f"Address : <b>{short_address}</b>\n\nFormula : <i>{display_formula}</i>"
                
                if has_resolved and resolved_formula and resolved_formula != raw_formula:
                    display_resolved = resolved_formula if resolved_formula.startswith('=') else f"={resolved_formula}"
                    simple_label += f"\n\nResolved : <i>{display_resolved}</i>"
                
                simple_label += f"\n\nValue     : {formatted_value}"
            else:
                simple_label = f"Address : <b>{short_address}</b>\n\nValue     : {formatted_value}"
            
            enhanced_tooltip = _create_enhanced_tooltip({
                'address': address,
                'formula': raw_formula,
                'value': value,
                'type': node_type,
                'filename': filename
            })

            nodes_data.append({
                "id": node_id,
                "label": simple_label,
                "color": color,
                "filename": filename,
                "level": node.get('depth', 0),
                "title": enhanced_tooltip,
                "shape": "box",
                "short_address_label": short_address,
                "full_address_label": full_address,
                "short_formula_label": short_formula,
                "full_formula_label": full_formula,
                "value_label": formatted_value,
                "resolved_formula": resolved_formula,
                "has_resolved": has_resolved
            })

        if parent_id is not None:
            edges_data.append((parent_id, node_id))

        for child in node.get('children', []):
            traverse_tree(child, parent_id=node_id)

    traverse_tree(dependency_tree_data)
    return nodes_data, edges_data
