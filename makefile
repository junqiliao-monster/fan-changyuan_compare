# 库的路径
LIB_PATH = ./venv/Lib/site-packages

# 要打包的文件名
FILE_NAME = compare.py

# 默认目标
all: clean package

# 使用PyInstaller打包程序
package:
	pyinstaller -F -p $(LIB_PATH) $(FILE_NAME)

# 删除dist、build目录和.spec文件
clean:
	rm -rf ./dist ./build
	rm -f $(FILE_NAME:.py=.spec)
