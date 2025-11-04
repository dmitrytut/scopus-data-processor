"""
Конфигурация по умолчанию для приложения обработки Scopus данных
"""

# Настройки поиска
DEFAULT_FUZZY_MATCH_THRESHOLD = 90  # Порог совпадения для fuzzy matching (0-100)

# Ключевые слова для поиска аффиляций пользователя
AFFILIATION_KEYWORDS = [
    'Khazar University',
    'Khazar',
    'Xəzər Universiteti'
]

# Подстроки для исключения статей по Title (по умолчанию пусто)
DEFAULT_TITLE_EXCLUDE_KEYWORDS = ['Correction:', 'Correction to:', 'Erratum to', 'Corrigendum to', '<FOR VERIFICATION>']

# Настройки форматирования Excel
HIGHLIGHT_COLOR_MULTIPLE_DEPTS = 'FFFF00'  # Желтый для множественных департаментов
HIGHLIGHT_COLOR_NO_DEPT = 'FFFF00'  # Желтый для отсутствующих департаментов

# Название листа по умолчанию в United файле
DEFAULT_UNITED_SHEET_NAME = 'Last'
