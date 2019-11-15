# -*- mode: python ; coding: iso-8859-1 -*-

block_cipher = None


a = Analysis(['TalanqueraUi.py'],
             pathex=['C:\\Users\\Administrator\\Documents\\Talanquera\\Talanquera'],
             binaries=[],
             datas=[('talanqueraUi.ui', '.'), ('LS.ico', '.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='TalanqueraUi',
          debug=False,
          strip=False,
          upx=True,
          console=False , icon='LS.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='TalanqueraUi')