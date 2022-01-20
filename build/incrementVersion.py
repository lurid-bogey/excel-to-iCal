import re

VERSION_FILE_NAME = '..\\excel-to-iCal\\version.py'
VERSION_INFO = 'versionInfo.txt'


def incrementVersion():
    f = open(VERSION_FILE_NAME, 'r', encoding='utf8')
    inLines = f.readlines()
    outLines = []

    for line in inLines:
        match = re.search(r'VERSION_PATCH\s*=\s*(\d+)', line)
        if match is not None:
            verBuild = int(match.group(1)) + 1
            print('Build index erhöht auf %s' % verBuild)
            outLines.append('VERSION_PATCH = %s\n' % verBuild)
        else:
            outLines.append(line)

    f.close()

    # WRITE THE INCREMENTED VERSION
    f = open(VERSION_FILE_NAME, 'w', encoding='utf8')
    f.write(''.join(outLines))
    f.close()

    # RE-READ THE VERSION AND INCORPORATE IT INTO THE locals()
    f = open(VERSION_FILE_NAME, 'r', encoding='utf8')
    content = f.read()
    f.close()

    exec(content)

    # WRITE VERSION_INFO.TXT
    VER_INFO = '''VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=%(fileversC)s,
    prodvers=%(fileversC)s,
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
    ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct('CompanyName', 'Smoke and Mirrors'),
        StringStruct('FileDescription', '%(DESCRIPTION)s'),
        StringStruct('FileVersion', '%(VERSION)s'),
        StringStruct('InternalName', '%(FILE_NAME)s'),
        StringStruct('LegalCopyright', 'GPL3 – 2022 – Smoke and Mirrors'),
        StringStruct('OriginalFilename', '%(FILE_NAME)s.exe'),
        StringStruct('ProductName', 'Smoke and Mirrors Collection'),
        StringStruct('ProductVersion', '%(fileversC)s')])
      ]),
    VarFileInfo([VarStruct(u'Translation', [1031, 1033])])
  ]
)
''' % locals()

    f = open(VERSION_INFO, 'w', encoding='utf8')
    f.write(VER_INFO)
    f.close()


if __name__ == '__main__':
    incrementVersion()
