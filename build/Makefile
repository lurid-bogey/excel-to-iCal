CC = gcc
CYTHON = cython
PYTHON = py -3

PROJECT_NAME = excel-to-iCal
PROJECT_DIR = D:\Projekte\iCal-gen.pycharm\excel-to-iCal
SRC_DIR = $(PROJECT_DIR)\excel-to-iCal
VERSION_DIR = $(PROJECT_DIR)\bin
BUILD_DIR = $(VERSION_DIR)\build
VERSION_INFO_TXT = versionInfo.txt

all: exe

exe: increment-version create-spec-file
	pyinstaller --onefile --workpath $(BUILD_DIR) --distpath $(VERSION_DIR) -y -p $(SRC_DIR) $(PROJECT_NAME).spec
	rmdir $(BUILD_DIR) /s /q
	del $(VERSION_INFO_TXT)

create-spec-file:
	pyinstaller --version-file $(VERSION_INFO_TXT) --onefile --workpath $(BUILD_DIR) --distpath $(VERSION_DIR) -y -p $(SRC_DIR) $(SRC_DIR)\$(PROJECT_NAME).py

increment-version:
	$(PYTHON) incrementVersion.py

cleandir:
	rmdir $(BUILD_DIR) /s /q

clean:
	del $(VERSION_DIR)\$(PROJECT_NAME).exe
