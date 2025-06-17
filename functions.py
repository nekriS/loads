def set_column_autowidth(ws, columns, reserve=1.2):
    """
    Устанавливает оптимальную ширину столбцов на основе содержимого.
    """
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter  # Получаем букву столбца (A, B, C, ...)
        
        if column in columns:
            # Находим максимальную длину текста в столбце
            for cell in column_cells:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
        
            # Устанавливаем ширину столбца с небольшим запасом
            adjusted_width = (max_length + 2) * reserve  # Можно изменить коэффициент для более комфортного отображения
            ws.column_dimensions[column].width = adjusted_width

def is_float(element: any) -> bool:
    if element is None: 
        return False
    try:
        float(element)
        return True
    except ValueError:
        return False
    
def getValue(current_str):
    current_str = str(current_str)
    current_str = current_str.lower()
    current_str = current_str.replace(' ','')
    current_str = current_str.replace('a','')
    #print(current_str)
    try:
        if 'm' in current_str:
            return float(current_str[:-1]) * pow(10, -3)
        elif 'u' in current_str:
            return float(current_str[:-1]) * pow(10, -6)
        elif 'n' in current_str:
            return float(current_str[:-1]) * pow(10, -9)
        elif 'p' in current_str:
            return float(current_str[:-1]) * pow(10, -12)
        elif is_float(current_str):
            return float(current_str)
        else:
            return ""
    except:
        return ""