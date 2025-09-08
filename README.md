## Инструмент для автоматического формирования таблицы дефектов мостовых сооружений по ОДМ 218.3.042-2014 в формате Microsoft Word

**Возможности**
* Группировка дефектов: автоматическая группировка записей по разделам сооружения (в соответствии с классификацией в ОДМ)
* Редактирование записей в процессе составления таблицы дефектов
* Генерация таблицы: создание структурированного отчета в формате Word (.docx).

## Установка
### Клонирование репозитория:
``` git clone https://github.com/alexkit78/bridge_defreport_tool.git ```

``` cd bridge_defreport_tool ```
### Установка зависимостей:

``` pip install -r requirements.txt ```

### Для создания .exe:
````pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db;." --add-data "report_template.docx;." --icon="icon.ico" --version-file "version.txt" --name "bdrt.exe" main.py````

### Для создания .pkg:
````pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db:." --add-data "report_template.docx:." --icon="icon.ico" --version-file "version.txt" --name "bdrt.exe" main.py````
