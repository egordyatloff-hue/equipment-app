[app]
title = Учёт поверки
package.name = equipment_calibration
package.domain = org.equipment
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 1.0
requirements = python3,kivy==2.3.1
p4a.branch = release-2024.01.21
orientation = portrait
fullscreen = 0
icon.filename = icon.png
android.api = 34
android.minapi = 24
android.ndk = 26b
android.archs = arm64-v8a,armeabi-v7a
android.allow_backup = True

[buildozer]
log_level = 2
warn_on_root = 1
