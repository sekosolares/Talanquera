# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['TalanqueraUi.py'],
             pathex=['C:\\Users\\Administrator\\Documents\\Talanquera\\Talanquera'],
             binaries=[],
             datas=[('LS.ico', '.'), ('talanqueraUi.ui', '.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='TalanqueraUi',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True , icon='LS.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='TalanqueraUi')
