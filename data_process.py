from pathlib import Path
import xlrd
from collections import Counter
import operator
import time

print('If you have any problem, please visit https://github.com/StarSkyLight.\n\n')

print('请输入Excel文件名。\n注意，如果程序与文件不在同一路径，请输入完整的文件路径。\n例如：\n'
      'Windows：D:\\example\\example.xls\n'
      'Linux：/home/username/example.xls\n'
      'macOS：~/username/example.xls\n'
      '如果需要处理的Excel文件在程序运行的目录，则可以直接输入文件名。\n'
      '如果需要关闭程序，请输入exit。')
while True:
    code = input()

    if code != 'exit':
        file_name = Path(code)

        extension_name = str(file_name).split('.').pop()

        if extension_name == 'xls' or extension_name == 'xlsx':
            try:
                work_book = xlrd.open_workbook(file_name)
            except FileNotFoundError:
                print('\n文件不存在。\n请重新输入文件名或输入exit关闭程序。')
                continue
            except OSError:
                print('\n文件路径格式错误。\n请重新输入文件名或输入exit关闭程序。')
                continue
            else:
                print('\n开始处理数据...')

                sheet_1 = work_book.sheet_by_index(0)

                result_list = []

                for index in range(1, sheet_1.nrows):
                    temp = sheet_1.cell_value(index, 0)

                    temp_list = temp.split('"，"')

                    if len(temp_list) == 1 and temp_list[0] == '':  # 去除空行
                        continue

                    temp_list[0] = temp_list[0].split('["')[1]
                    temp_list.reverse()
                    temp_list[0] = temp_list[0].split('"]')[0]

                    result_list.extend(temp_list)

                temp_result = dict(Counter(result_list))
                result = sorted(temp_result.items(), key=operator.itemgetter(1), reverse=True)

                print('数据处理完成')

                print('\n请输入保存处理结果的文件名，文件将被以该文件名保存在当前目录。\n不需要输入文件扩展名，'
                      '默认使用csv格式储存。\n直接按Enter键在当前界面输出。')
                output_file_path = input()

                if output_file_path == '':
                    for t in result:
                        print(t[0] + '\t' + str(t[1]))
                else:
                    file_output = open(output_file_path + '.csv', 'w')
                    for t in result:
                        file_output.write(t[0] + ',' + str(t[1]) + '\n')
                    file_output.close()

                print('\n数据输出完毕。\n请重新输入文件名或输入exit关闭程序。')
        else:
            print('\n该文件不是Excel文件。\n请重新输入文件名或输入exit关闭程序。')
    else:
        print('\nSee You Next Time', end='')
        time.sleep(0.5)
        print('.', end='')
        time.sleep(0.5)
        print('.', end='')
        time.sleep(0.5)
        print('.')
        print('Bye~~~')
        time.sleep(2)

        break
