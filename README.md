# Bridge Defect Report Tool

## Automated tool for generating bridge structure defect tables according to ODM 218.3.042-2014 in Microsoft Word format

**Key Features**
* **Defect grouping:** Automatic grouping of records by structure sections (according to ODM classification)
* **Record editing:** Editing entries during the table creating process
* **Table generation:** Creation of structured reports in Word (.docx) format

## Installation

### Clone the repository:
```git clone https://github.com/alexkit78/bridge_defreport_tool.git ```

```cd bridge_defreport_tool```

### Install dependencies:

``` pip install -r requirements.txt ```
## Building executables
### For Windows (.exe):
````pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db;." --add-data "report_template.docx;." --icon="icon.ico" --version-file "version.txt" --name "bdrt.exe" main.py````

### For macOS/Linux (.pkg and other formats):
````pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db:." --add-data "report_template.docx:." --icon="icon.icns" --version-file "version.txt" --name "bdrt.pkg" main.py````

_____________________________________________________


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
## Формирование исполняемого файла
### Для создания .exe:
```pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db;." --add-data "report_template.docx;." --icon="icon.ico" --version-file "version.txt" --name "bdrt.exe" main.py```

### Для создания .pkg (macOS/Linux) :
```pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db:." --add-data "report_template.docx:." --icon="icon.icns" --version-file "version.txt" --name "bdrt.pkg" main.py```
