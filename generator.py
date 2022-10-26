from asyncio import ensure_future
from io import BufferedReader, BufferedWriter
from typing import Any, Dict, List, Union
from docx import Document

import rich
import os
import json
import re


lib_name_list: List[str] = ['primary', 'middle']
json_file_name = 'lib2class.json'

def log(*message: Any, color: Union[str, None] = None) -> None:
    message: str = ' '.join([str(single) for single in message])
    rich.print(f'[{color}]{message}[/{color}]' if color else message)

def handle_error(message: str) -> None:
    log(f'处理失败: {message}', color='red')
    exit(-1)

def parse_int(string: str) -> int:
    if re.match(r'^[0-9]*$', string):
        return int(string)
    else:
        handle_error(f'`{string}` 不是一个有效的数字!')

def main(year: int) -> None:
    path: str = f'./{year}'
    result: Dict[str, List[Dict[str, str]]] = {}
    for name in lib_name_list:
        file_path: str = f'{path}/{name}.docx'
        if os.path.exists(file_path): 
            file: BufferedReader = open(file_path, 'rb')
            document: Document = Document(file)
            log(f'成功读取 `{file_path}` 到文档对象: {document}')
            temp: List[Dict[str, str]] = []
            for para in document.paragraphs:
                text: str = para.text.strip()
                num: re.Match = re.match(r'^[0-9]*(、)', text)
                ans: re.Match = re.match(r'( *)(答案：)[A-Z]*', text)
                if para.style.name == 'Normal':
                    if num:
                        number: str = num.group()
                        removed: str = text.replace(number, '')
                        temp.append({ removed: None })
                    elif ans:
                        answer: str = ans.group().split('：')[-1]
                        temp[-1][[t for t in temp[-1].keys()][0]] = answer
            result[name] = temp
            log(f'({name}) 成功加载 {temp.__len__()} 个题目!', color='green')
        else:
            handle_error(f'无法找到 `{file_path}`!')
            exit(-1)
    json_file: BufferedWriter = open(json_file_name, 'w', encoding='utf-8')
    json.dump(result, json_file, ensure_ascii=False, indent=2)
    log(f'已写入文件 `{json_file_name}`!')

if __name__ == '__main__':
    year: int = parse_int(input('请输入学年 > '))
    if str(year).__len__() == 4 and year >= 2022:
        main(year)
    else:
        handle_error(f'`{year}` 不是一个有效的年份!')
