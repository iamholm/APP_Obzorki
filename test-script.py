import unittest
from docx_to_excel_processor import DocxToExcelProcessor

class TestDateProcessing(unittest.TestCase):
    def setUp(self):
        self.processor = DocxToExcelProcessor()
    
    def test_extract_dates_with_spaces(self):
        """Тест на извлечение дат с пробелами"""
        text = "Постановлением суда от 13. 05. 2024 осужден по ст. 158"
        dates = self.processor._extract_all_dates_from_text(text)
        self.assertEqual(len(dates), 1)
        self.assertIn("13. 05. 2024", dates)
        
        text_multiple = "Задержан 10. 01. 2024, освобожден 15. 02. 2024"
        dates = self.processor._extract_all_dates_from_text(text_multiple)
        self.assertEqual(len(dates), 2)
        self.assertIn("10. 01. 2024", dates)
        self.assertIn("15. 02. 2024", dates)
    
    def test_normalize_dates_with_spaces(self):
        """Тест на нормализацию дат с пробелами"""
        # Даты с пробелами
        self.assertEqual(self.processor._parse_and_normalize_date("13. 05. 2024"), "13.05.2024")
        self.assertEqual(self.processor._parse_and_normalize_date("1. 5. 24"), "01.05.2024")
        self.assertEqual(self.processor._parse_and_normalize_date("10 . 12 . 2023"), "10.12.2023")
        
        # Двузначные годы с пробелами
        self.assertEqual(self.processor._parse_and_normalize_date("12. 09. 95"), "12.09.1995")
        self.assertEqual(self.processor._parse_and_normalize_date("03. 02. 22"), "03.02.2022")
        
        # Другие форматы с пробелами
        self.assertEqual(self.processor._parse_and_normalize_date("14 / 07 / 2023"), "14.07.2023")
        self.assertEqual(self.processor._parse_and_normalize_date("19 - 11 - 2022"), "19.11.2022")
    
    def test_date_detection(self):
        """Тест на определение строки как даты"""
        # Стандартные форматы
        self.assertTrue(self.processor._is_date("14.07.2023"))
        self.assertTrue(self.processor._is_date("01.05.22"))
        
        # Форматы с пробелами
        self.assertTrue(self.processor._is_date("15. 08. 2023"))
        self.assertTrue(self.processor._is_date("3. 4. 21"))
        self.assertTrue(self.processor._is_date("10 . 12 . 2022"))
        
        # Не даты
        self.assertFalse(self.processor._is_date("Не дата"))
        self.assertFalse(self.processor._is_date(""))
        self.assertFalse(self.processor._is_date("14,07,2023"))  # Запятые вместо точек
    
    def test_court_info_date_normalization(self):
        """Тест на нормализацию дат в информации о судах"""
        # Создаем простой мок-объект листа
        class MockSheet:
            def __init__(self):
                self.cells = {}
                self.max_row = 1
            
            def cell(self, row, column):
                key = (row, column)
                if key not in self.cells:
                    self.cells[key] = MockCell()
                return self.cells[key]
        
        class MockCell:
            def __init__(self):
                self.value = None
        
        # Создаем тестовый лист с данными
        sheet = MockSheet()
        sheet.cell(1, 9).value = "Постановлением суда от 13. 05. 2024 осужден по ст. 158"
        
        # Нормализуем даты
        normalized_count = self.processor._normalize_dates_in_court_info(sheet, 9)
        
        # Проверяем результаты
        self.assertEqual(normalized_count, 1)
        self.assertEqual(
            sheet.cell(1, 9).value, 
            "Постановлением суда от 13.05.2024 осужден по ст. 158"
        )
        
        # Проверяем случай с несколькими датами
        sheet.cell(1, 9).value = "Задержан 10. 01. 2023, освобожден 15. 02. 2023"
        normalized_count = self.processor._normalize_dates_in_court_info(sheet, 9)
        self.assertEqual(normalized_count, 2)
        self.assertEqual(
            sheet.cell(1, 9).value, 
            "Задержан 10.01.2023, освобожден 15.02.2023"
        )


if __name__ == "__main__":
    unittest.main()
