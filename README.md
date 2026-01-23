# Bridge Report Tool

## Automated tool for generating passport with bridge structure defect tables according to ODM 218.3.042-2014 in Microsoft Word format and Bridge Inspection Report.

**Key Features**
* **Defect grouping:** Automatic grouping of records by structure sections (according to ODM classification)
* **Record editing:** Editing entries during the table creating process
* **Documents generation:** Creation of structured reports in Word (.docx) format

## Installation

### Clone the repository:
```git clone https://github.com/alexkit78/bridge_defreport_tool.git ```

```cd bridge_defreport_tool```

### Install dependencies:

``` pip install -r requirements.txt ```
## Building executables




 ### For Windows (.exe):

```python -m PyInstaller bridge_reptool.spec```
<!--
````pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db;." --add-data "report_template.docx;." --icon="icon.ico" --version-file "version.txt" --name "bdrt.exe" main.py````
-->
### For macOS/Linux (.pkg and other formats):

```python3 -m PyInstaller bridge_reptool.spec```
<!--````pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db:." --add-data "report_template.docx:." --icon="icon.icns" --version-file "version.txt" --name "bdrt.pkg" main.py```` -->

_____________________________________________________


## Инструмент для автоматического формирования Паспорта мостового сооружения с таблицей дефектов по ОДМ 218.3.042-2014, а также отчётов по обследованию мостового сооружения в формате Microsoft Word

**Возможности**
* Группировка дефектов: автоматическая группировка записей по разделам сооружения (в соответствии с классификацией в ОДМ)
* Редактирование записей в процессе составления таблицы дефектов
* Формирование документов: создание структурированного паспорта и отчёта в формате Word (.docx) по введённым данным.

## Установка
### Клонирование репозитория:
``` git clone https://github.com/alexkit78/bridge_defreport_tool.git ```

``` cd bridge_defreport_tool ```
### Установка зависимостей:

``` pip install -r requirements.txt ```
## Формирование исполняемого файла

### Для создания .exe (Windows):
<!--
````pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db;." --add-data "report_template.docx;." --icon="icon.ico" --version-file "version.txt" --name "bdrt.exe" main.py````
-->
```python -m PyInstaller bridge_reptool.spec```

### Для создания .pkg (macOS/Linux) :

```python3 -m PyInstaller bridge_reptool.spec```
<!--
```pyinstaller --noconsole --onefile -n "bridge defect report tool" --add-data "bridge_defects.db:." --add-data "report_template.docx:." --icon="icon.icns" --version-file "version.txt" --name "bdrt.pkg" main.py``` -->