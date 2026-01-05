# create_version.py
from PyInstaller.utils.win32 import versioninfo

# Данные версии
version_data = versioninfo.VSVersionInfo(
    ffi=versioninfo.FixedFileInfo(
        filevers=(1, 3, 0, 0),
        prodvers=(1, 3, 0, 0),
        mask=0x3f,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(0, 0)
    ),
    kids=[
        versioninfo.StringFileInfo([
            versioninfo.StringTable(
                '041904e3',
                [
                    versioninfo.StringStruct('CompanyName', ''),
                    versioninfo.StringStruct('FileDescription', 'Оценка состояния мостовых сооружений'),
                    versioninfo.StringStruct('FileVersion', '1.3.0.0'),
                    versioninfo.StringStruct('InternalName', 'Bridge defects report tool'),
                    versioninfo.StringStruct('LegalCopyright', 'Чечеткин '
                                                               'Александр © '
                                                               '2025'),
                    versioninfo.StringStruct('OriginalFilename', 'bdrt.exe'),
                    versioninfo.StringStruct('ProductName', 'Оценка состояния мостовых сооружений'),
                    versioninfo.StringStruct('ProductVersion', '1.3.0.0')
                ]
            )
        ]),
        versioninfo.VarFileInfo([versioninfo.VarStruct('Translation', [1049, 65001])])
    ]
)

# Сохраняем в файл
with open('version.txt', 'w', encoding='utf-8') as f:
    f.write(version_data.__str__())

print("Файл version.txt создан успешно!")