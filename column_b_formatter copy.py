import re

class ColumnBFormatter:
    """
    Класс для форматирования текста в столбце B (имена)
    """
    
    def process_excel_column(self, sheet, column_index=2):
        """
        Обрабатывает все ячейки в столбце B
        
        Args:
            sheet: Лист Excel
            column_index (int): Индекс столбца (по умолчанию 2 для столбца B)
            
        Returns:
            dict: Статистика обработки
        """
        stats = {
            'cells_processed': 0,
            'names_split': 0,
            'names_moved': 0,
            'names_formatted': 0  # Для совместимости
        }
        
        # Определяем, с какой строки начать (обычно с 1, но может быть с 2 если первая строка заголовок)
        start_row = 1
        
        for row in range(start_row, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            if cell.value:
                original_text = str(cell.value)
                
                # Разделяем слитные имена (например, "ЖумановИсабекМаратбекович")
                formatted_text = self._split_joined_names(original_text)
                
                if formatted_text != original_text:
                    cell.value = formatted_text
                    stats['cells_processed'] += 1
                    stats['names_split'] += 1
                    stats['names_formatted'] += 1
                
                # Перенос имени в столбец N
                if self._move_name_to_column_n(sheet, row, cell.value):
                    stats['names_moved'] += 1
        
        return stats
    
    def _split_joined_names(self, text):
        """
        Разделяет слитные имена с заглавными буквами внутри слова
        
        Args:
            text (str): Исходный текст
            
        Returns:
            str: Текст с разделенными именами
        """
        # Не обрабатываем специальные случаи с "СПб"
        if re.search(r'СПб|[гГ]\.?\s*СПб', text):
            return text
        
        # Удаляем лишние пробелы
        text = re.sub(r'\s+', ' ', text).strip()
            
        # Разбиваем текст на слова
        words = text.split()
        result_words = []
        
        # Берем первые три слова или все, если их меньше трех
        for word in words[:min(3, len(words))]:
            # Проверяем, есть ли внутри слова заглавные буквы (не в начале)
            if len(word) > 1 and re.search(r'[А-ЯЁA-Z]', word[1:]):
                # Находим все позиции заглавных букв
                capitals = [0] + [i for i in range(1, len(word)) if word[i].isupper()]
                
                # Если нашли заглавные буквы внутри слова
                if len(capitals) > 1:
                    # Разделяем слово по позициям заглавных букв
                    split_words = []
                    for i in range(len(capitals)):
                        start = capitals[i]
                        end = capitals[i+1] if i+1 < len(capitals) else len(word)
                        split_words.append(word[start:end])
                    
                    # Добавляем разделенные слова
                    result_words.extend(split_words)
                    continue
            
            # Если слово не нужно разделять, добавляем как есть
            result_words.append(word)
        
        # Добавляем оставшиеся слова, если они есть
        if len(words) > 3:
            result_words.extend(words[3:])
            
        return ' '.join(result_words)
    
    def _check_for_patronymic(self, sheet, row):
        """
        Проверяет, остались ли в столбце B отчества, и переносит их в столбец N
        
        Args:
            sheet: Лист Excel
            row (int): Номер строки
        """
        cell_b = sheet.cell(row=row, column=2)
        cell_n = sheet.cell(row=row, column=14)
        
        if not cell_b.value:
            return
            
        text = str(cell_b.value).strip()
        words = text.split()
        
        if not words:
            return
            
        # Проверяем первое слово на признаки отчества
        first_word = words[0]
        
        # Типичные окончания отчеств
        patronymic_endings = ['вич', 'вна', 'ич', 'ична', 'ична', 'овна', 'евна', 'ович', 'евич']
        
        is_patronymic = False
        
        # Проверяем, содержит ли первое слово типичные окончания отчеств
        for ending in patronymic_endings:
            if first_word.lower().endswith(ending) and len(first_word) > len(ending) + 1:
                is_patronymic = True
                break
                
        # Если первое слово - отчество
        if is_patronymic:
            # Проверяем, не является ли следующее слово "СПб" или "г."
            stop_words = ['СПб', 'г.', 'г.СПб', 'г. СПб', 'пр.', 'ул.', 'д.']
            
            if len(words) > 1 and any(words[1] == stop_word for stop_word in stop_words):
                # Переносим только отчество в столбец N
                patronymic = first_word
                
                # Добавляем отчество в столбец N
                if cell_n.value:
                    cell_n.value = str(cell_n.value) + ' ' + patronymic
                else:
                    cell_n.value = patronymic
                
                # Удаляем отчество из столбца B
                remaining_text = ' '.join(words[1:]).strip()
                cell_b.value = remaining_text if remaining_text else None
    
    def _move_name_to_column_n(self, sheet, row, name_text):
        """
        Переносит имя в столбец N и удаляет только имя из B, оставляя остальной текст
        
        Args:
            sheet: Лист Excel
            row (int): Номер строки
            name_text (str): Текст имени
            
        Returns:
            bool: True если имя было перенесено, False в противном случае
        """
        # Получаем ячейку в столбце N
        cell_n = sheet.cell(row=row, column=14)
        cell_b = sheet.cell(row=row, column=2)
        
        # Если текст пустой, выходим
        if not name_text:
            return False
            
        name_text = str(name_text).strip()
        
        # Разбиваем текст на слова
        words = name_text.split()
        
        # Если текст пустой после разбиения, выходим
        if not words:
            return False
        
        # Стоп-слова, которые указывают на конец имени
        stop_words = ['СПб', 'г.', 'г.СПб', 'г. СПб', 'пр.', 'ул.', 'д.', 'р-н', 'обл.', 'респ.']
        
        # Проверяем, начинается ли текст со стоп-слова
        if words[0] in stop_words:
            return False
        
        # Берем первые три слова или до первого стоп-слова
        name_parts = []
        max_words = min(3, len(words))
        
        for i in range(max_words):
            # Если встретили стоп-слово, останавливаемся
            if words[i] in stop_words:
                break
            name_parts.append(words[i])
        
        # Если не нашли имени, выходим
        if not name_parts:
            return False
        
        # Формируем имя
        formatted_name = ' '.join(name_parts)
        
        # Если в ячейке N уже есть текст, добавляем имя в начало
        if cell_n.value:
            cell_n.value = formatted_name + '. ' + str(cell_n.value)
        else:
            cell_n.value = formatted_name
        
        # Удаляем первые три слова (имя) из столбца B и оставляем остальной текст
        # Определяем, сколько слов мы взяли для имени
        name_word_count = len(name_parts)
        
        # Собираем оставшиеся слова
        remaining_words = words[name_word_count:]
        
        # Преобразуем обратно в текст
        remaining_text = ' '.join(remaining_words).strip()
        
        # Обновляем ячейку B
        cell_b.value = remaining_text if remaining_text else None
        
        return True 