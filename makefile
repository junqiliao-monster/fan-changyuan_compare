# 库的路径
LIB_PATH = C:/Users/Administrator/AppData/Local/Programs/Python/Python38/Lib/site-packages

# 要打包的文件名
FILE_NAME = compare.py

# 默认目标
all: clean package

# 使用PyInstaller打包程序
package: 
	pyinstaller \
	--hidden-import=pyexcel_xls \
	--hidden-import=pyexcel_xlsx \
	--hidden-import=pyexcel_xlsxw \
	--hidden-import pyexcel_io.readers.csv_in_file \
	--hidden-import pyexcel_io.readers.csv_in_memory \
	--hidden-import pyexcel_io.readers.csv_content \
	--hidden-import pyexcel_io.readers.csvz \
	--hidden-import pyexcel_io.writers.csv_in_file \
	--hidden-import pyexcel_io.writers.csv_in_memory \
	--hidden-import pyexcel_io.writers.csvz_writer \
	--hidden-import pyexcel_io.database.importers.django \
	--hidden-import pyexcel_io.database.importers.sqlalchemy \
	--hidden-import pyexcel_io.database.exporters.django \
	--hidden-import pyexcel_io.database.exporters.sqlalchemy \
	-p $(LIB_PATH) $(FILE_NAME) --onefile \

# 删除dist、build目录和.spec文件
clean:
	rm -rf ./dist ./build
	rm -f $(FILE_NAME:.py=.spec)