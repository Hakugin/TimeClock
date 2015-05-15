from distutils.core import setup
import sys
import os
import py2exe
import sqlalchemy

if len(sys.argv) == 1:
    sys.argv.append("py2exe")
    sys.argv.append("-q")

manifest = """
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<assembly xmlns='urn:schemas-microsoft-com:asm.v1' manifestVersion='1.0'>
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
    <security>
      <requestedPrivileges>
        <requestedExecutionLevel level='asInvoker' uiAccess='false' />
      </requestedPrivileges>
    </security>
  </trustInfo>
  <dependency>
    <dependentAssembly>
      <assemblyIdentity
     type='win32'
     name='Microsoft.VC90.CRT'
     version='9.0.21022.8'
     processorArchitecture='*'
     publicKeyToken='1fc8b3b9a1e18e3b' />
    </dependentAssembly>
  </dependency>
  <dependency>
    <dependentAssembly>
      <assemblyIdentity
         type="win32"
         name="Microsoft.Windows.Common-Controls"
         version="6.0.0.0"
         processorArchitecture="*"
         publicKeyToken="6595b64144ccf1df"
         language="*" />
    </dependentAssembly>
  </dependency>
</assembly>
"""

excludes = [
# Normally excluded
    'Tkinter', 'doctest', 'unittest', 'pydoc', 'pdb', 'curses',
    'email', '_tkagg', '_gtkagg', 'bsddb', 'tcl', 'difflib',

# Additional excludess - Tested and program appears to work normally
    '_sha512', 'gc', 'win32evtlog', 'email.mime', 'msvcrt',
    'imp', '_win32sysloader', 'signal', 'cStringIO', 'zlib',
    '_sha', '_winreg', '_subprocess', 'math', '_ctypes',
    'exceptions', '_functools', '_locale', 'thread', 'itertools',
    '_collections', '_sre', '__builtin__', 'operator', 'array',
    'select', '_heapq', 'errno', 'binascii', '_sha256',
    '_warnings', 'cPickle', '_codecs', 'win32ui', '_struct',
    '_hashlib', '_random', 'parser', '_md5', '_io', 'strop',
    'nt', '_ssl', 'sys', 'bz2', '_symtable', '_weakref',
    'marshal', 'time', 'ssl', 'riscos', 'riscosenviron',
    'riscospath', 'hmac', 'compiler', 'getpass', 'pyreadline',
    #'hashlib',
    'uu', 'email.base64mime', 'pywin',

# Additional excludes - Testing
    'clr', 'grp', 'System', 'startup', 'fcntl', 'ce', '_emx_link',
    'IronPythonConsole', 'ctypes.cdll',  'os2', 'win32com.gen_py',
    'ctypes._SimpleCData', 'posix', 'pwd', '_scproxy', 'SOCKS',
    #'listenerimpl',
    'EasyDialogs', 'publishermixin', 'rourl2path',
    '_sysconfigdata', 'termios', 'topicargspecimpl', 'topicmgrimpl',
    'modes.editingmodes'
]


dll_excludes = [
   'API-MS-Win-Core-LocalRegistry-L1-1-0.dll',
   'MPR.dll',
   'MSWSOCK.DLL',
   'POWRPROF.dll',
   'profapi.dll',
   'userenv.dll',
   'w9xpopen.exe',
   'wtsapi32.dll',
   'libgdk-win32-2.0-0.dll',
   'libgobject-2.0-0.dll',
   'tcl84.dll',
   'tk84.dll'
]

package_includes = [
                    'sqlalchemy',
                    'sqlalchemy.dialects.sqlite',
                    'gzip',
                    'ctypes',
                    '_outlook',
                    '_dialogs',
                    '_widgets',
                    '_logger',
                    '_database',
                    'wx.lib.pubsub',
                    'win32com.server.util'
                   ]

py2exe_options = {
   'optimize': 2, # 0 (None), 1 (-O), 2 (-OO)
   'excludes': excludes,
   'dll_excludes': dll_excludes,
   'packages': package_includes,
   'xref': False,
   # bundle_files: 1|2|3
   #    1: executable and library.zip
   #    2: executable, Python DLL, library.zip
   #    3: executable, Python DLL, other DLLs and PYDs, library.zip
   'bundle_files': 1,
  }

setup(
      windows=[{
          'script' : 'ClockPuncher3.py',
          'other_resources' : [(24, 1, manifest)],
          'icon_resources'  : [(1, 'myicons.ico')]
      }],
      version='3.0.2',
      options={'py2exe': py2exe_options},
      zipfile = None
    )
