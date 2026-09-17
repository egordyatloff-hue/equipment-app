[app]
title = Учёт поверки
package.name = equipment_calibration
package.domain = org.equipment
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 1.0
requirements = python3,kivy==2.3.1,filetype,openpyxl==3.1.5,et_xmlfile==2.0.0
p4a.branch = release-2024.01.21
orientation = portrait
fullscreen = 0
icon.filename = icon.png
android.api = 34
android.minapi = 24
android.ndk = 25b
android.archs = arm64-v8a,armeabi-v7a
android.allow_backup = True
android.permissions = WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE
android.extra_pip_args = --trusted-host pypi.org --trusted-host files.pythonhosted.org

[buildozer]
log_level = 2
warn_on_root = 1
