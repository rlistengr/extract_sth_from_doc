# extract_sth_from_doc

# 1. 安装python
# 2. 安装python-docx 使用命令pip3 install python-docx
# 运行该python程序，python 该程序名字 文档路径 正则表达式 是否去除重复（默认去除，false为不去重）
# 比如改成保存为extract.py 去重的方式抽取所有的书名，程序和文档在同一个路径下，且名字叫test.docx，
# 运行python extract.py test.docx ".*(《.*》).*"   
# 注意上面的破折号是中文输入法的，其他均是英文输入法的


from docx import Document

import sys
import os
import re

if len(sys.argv) < 3:
    print ("usage: extract.py file_path key_word remove_duplicates.\nthe remove_duplicates's default value is true.")
    sys.exit(1)

file_path = os.path.abspath(sys.argv[1])
key_word = sys.argv[2]
remove_duplicates = True
if len(sys.argv) > 3 and (sys.argv[3] == "False" or sys.argv[3] == "false"):
    remove_duplicates = False

print(file_path, key_word, remove_duplicates)

doc_handle = Document(file_path)
tables = doc_handle.tables;
records = []
for table in tables:
    n_rows = len(table.rows)
    n_columns = len(table.columns)
    for i in range(0, n_rows):
        for j in range(0, n_columns):
            text = table.cell(i, j).text
            linebits = re.match(key_word, text)
            if linebits:
                value = linebits.groups()[0]
                if value not in records:
                    records.append(value)

for record in records:
    print(record)
            
