# -*- mode: python -*-

block_cipher = None


a = Analysis(['pdf_toolkit.py'],
             pathex=['C:\\Users\\c41663\\Documents\\programming\\pdftools'],
             binaries=[],
             datas=[],
             hiddenimports=['pikepdf._cpphelpers'],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='pdf_toolkit',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
