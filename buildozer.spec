[app]
title = Horas RDP
package.name = horasrdp
package.domain = org.rdp
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 10.0
# sqlite3 debe ir ANTES de python3, o la app falla al abrir con "No module named _sqlite3"
requirements = sqlite3,python3,kivy==2.3.0,openpyxl,et_xmlfile,plyer,android,pyjnius
orientation = portrait
fullscreen = 0
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,READ_MEDIA_IMAGES,MANAGE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 24
# Solo 64 bits (compila mucho más rápido). Si tu teléfono es de 32 bits: arm64-v8a,armeabi-v7a
android.archs = arm64-v8a
android.accept_sdk_license = True

[buildozer]
log_level = 2
warn_on_root = 1
