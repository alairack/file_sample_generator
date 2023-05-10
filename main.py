from file_writer import FileWriter
import os

SOURCE_FILE_LIMIT_SIZE = 11000    # 输入文件大小限制

task = {'pdf': 60, 'doc': 100, 'docx': 100, 'ppt': 30, 'pptx': 30, 'xls': 20, 'xlsx': 20, 'ofd': 30, "wps": 100,
        'et': 20}

secret_task = {'pdf': 15, 'doc': 20, 'docx': 20, 'ppt': 15, 'pptx': 15, 'xls': 5, 'xlsx': 5, 'ofd': 3, "wps": 20,
               'et': 5}


def get_input_file_list(folder_path: str):
    source_file_list = []

    for root, dirs, files in os.walk(os.path.abspath(folder_path)):
        size_list = []
        for name in files:
            file_path = os.path.join(root, name)
            file_size = os.path.getsize(file_path)
            if file_size > SOURCE_FILE_LIMIT_SIZE and file_size not in size_list:
                source_file_list.append(file_path)
                size_list.append(file_size)

    return source_file_list


def generate_sample_files(input_file_list: list):
    index = 0
    serial_number = 1
    for file_type, number_of_file in secret_task.items():
        generated_number = 1
        while generated_number <= number_of_file:
            with open(input_file_list[index], 'r', encoding='utf-8') as f:
                file_content = f.read()
            index = index + 1
            formatted_number = "{:03d}".format(serial_number)

            try:
                FileWriter(file_content, f'DL-模拟不相关涉密文档-{formatted_number}.{file_type}', file_type, True)
            except Exception:
                print(formatted_number)
            else:
                generated_number = generated_number + 1
                serial_number = serial_number + 1


if __name__ == "__main__":
    generate_sample_files(get_input_file_list('\\new'))
