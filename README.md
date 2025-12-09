# Authorization Service

[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Authorization Service is an simple service for users check.

## XLSX Hyperlink Converter

Преобразует Excel файлы из 1C с формулами HYPERLINK в файлы с кликабельными гиперссылками, совместимые с MS Office и LibreOffice.

### Использование:

```javascript
const xlsxService = require('./src/server/services/xlsxService.js');

// Основной метод (автоматически выберет лучший способ конвертации)
const result = await xlsxService.doConvertXlsx('path/to/input.xlsx');

// Прямые методы
const result1 = xlsxService.convertWithSheetJS('input.xlsx', 'output.xlsx'); // Рекомендуемый
const result2 = xlsxService.convertToHyperlinks('input.xlsx', 'output.xlsx'); // Fallback
```

### Зависимости:
- `xlsx` (SheetJS) - основная библиотека для максимальной MS Office совместимости
- `adm-zip` - для работы с XLSX архивами

### Особенности:
- ✅ Динамическое извлечение гиперссылок из любых файлов
- ✅ Поддержка MS Office и LibreOffice  
- ✅ Сохранение всех данных из оригинального файла
- ✅ Автоматический выбор наиболее совместимого метода конвертации

## Features

- Make requests and return data ...

## Requirements

- [UNIX)
- [Node.js >= 18.x](https://nodejs.org/en/download/)

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/.../AuthorizationService.git
   ```
