# t.ly statistics parser
Python script that can help you collect statistics on t.ly short links. Uses Excel tables as an interface
### How to use(EN):
1. Install necessary python modules: `pip install -r requirements.txt`
2. Get your API token on https://t.ly/settings#/api
3. Сreate an Excel file where statistics will be collected
4. Insert your API token and the path to the Excel file into `TOKEN` and `FILEPATH` in `config.py`
5. Enter your short t.ly links in `A` column. __Skip first row, there will be headers.__
6. Run `main.py`. __Your Excel file should not be opened in any other program.__


### How to use(RU):
1. Установите необходимые модули: `pip install -r requirements.txt`
2. Получите API токен по адресу https://t.ly/settings#/api
3. Создайте Excel файл, куда будет собираться статистика
4. Вставьте ваш API токен и путь до Excel файла в переменные `TOKEN` и `FILEPATH` в файле `config.py`
5. Введите ваши короткие t.ly ссылки в столбец `A`. __Пропустите первую строку, там будут заголовки.__
6. Запустите `main.py`. __Другие программы не должны использовать ваш Excel файл в этот момент.__