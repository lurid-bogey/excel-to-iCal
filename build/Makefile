CC = gcc
CYTHON = cython

PROJECT_NAME = excel-to-iCal
PROJECT_DIR = D:\Projekte\iCal-gen.pycharm\excel-to-iCal
SRC_DIR = $(PROJECT_DIR)\excel-to-iCal
VERSION_DIR = $(PROJECT_DIR)\bin
BUILD_DIR = $(VERSION_DIR)\build


all: exe

exe: create-spec-file
	pyinstaller --onefile --workpath $(BUILD_DIR) --distpath $(VERSION_DIR) -y -p $(SRC_DIR) $(PROJECT_NAME).spec
	rmdir $(BUILD_DIR) /s /q

create-spec-file:
	pyinstaller --version-file versionInfo.txt --onefile --workpath $(BUILD_DIR) --distpath $(VERSION_DIR) -y -p $(SRC_DIR) $(SRC_DIR)\$(PROJECT_NAME).py

cleandir:
	rmdir $(BUILD_DIR) /s /q

clean:
	del $(VERSION_DIR)\$(PROJECT_NAME).exe
