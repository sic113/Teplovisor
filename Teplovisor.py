import os, shutil, openpyxl, re


builds_dict = {
        0: "АВК",
        1: "АТК",
        2: "ВДС",
        3: "ГКО",
        4: "Грузовой",
        5: "Котельная",
        6: "ОМТС",
        7: "ПБЗ",
        8: "СРТ",
        9: "DutyFree",
        10: "VIP"
    }

names_dict = {
            5: 'IR00000',
            6: 'IR0000',
            7: 'IR000',
            8: 'IR00',
            9: 'IR0',
            10: 'IR'
        }


teplak_path = "Teplak"
list_photo_names = os.listdir(teplak_path)
teplak_photo_names_string = '_'.join(list_photo_names)

def folders_create():
    tree_folder = "Tree"
    if not os.path.exists(tree_folder):
        os.mkdir(tree_folder)
        print(f"Директория '{tree_folder}' успешно создана!")

    for i in range(0, len(builds_dict)):
        if not os.path.exists(f"{tree_folder}/{builds_dict[i]}"):
            os.mkdir(f"{tree_folder}/{builds_dict[i]}")
            print(f"Директория '{tree_folder}/{builds_dict[i]}' успешно создана!")


def zipper(list_number):
    wb = openpyxl.load_workbook('list.xlsx')
    wb.active = list_number
    sheet = wb.active

    A1 = []
    for cell in sheet['A']:
        A1.append(cell.value)

    B1 = []
    for cell in sheet['B']:
        B1.append(cell.value)

    A1_B1 = dict(zip(A1, B1))
    return A1_B1


def copyer(build_num, n):

    print("- - - - - "+str(builds_dict[build_num])+" - - - - -")
    build = zipper(build_num)

    if not os.path.exists('Tree/' + str(builds_dict[build_num])):
        os.mkdir('Tree/' + str(builds_dict[build_num]))

    for k in build:

        excel_string = str(build[k])
        excel_key = excel_string.split(',')

        i = len(excel_key)

        for i in range(0,i):

            photo_key_2 = str(excel_key[i]) + '.BMT'
            sum = len(photo_key_2)
            photo_key_1 = names_dict[sum]
            full_photo_name = photo_key_1+photo_key_2
            match = re.search(photo_key_2,teplak_photo_names_string)

            if match:

                file_path = 'Teplak/'+full_photo_name
                papka_path = 'Tree/' + str(builds_dict[build_num]) + '/' + str(k)

                if not os.path.exists(papka_path):
                    os.mkdir('Tree/' + str(builds_dict[build_num]) + '/' + str(k))
                try:
                    # shutil.copy(file, papka)  'copy' - функция модуля 'shutil' с аргументами 'file' - локальный путь к файлу, 'papka' - локальный путь к папке, функция копирует файл(file) в папку(papka)
                    shutil.move(file_path, papka_path)  # 'move' для перемещения
                    n += 1
                    print(f"Фото: {full_photo_name} перемещено в {papka_path}")
                    print(f"Фотографий перемещено: {n}")
                except:
                    print(f"- - - - - - - Ошибка перемещения {full_photo_name} - - - - - - -")
    return n


def main():
    folders_create()
    count = 0
    for i in range(0, 11):
        count = copyer(i,count)
        print('Готово!')


if __name__ == '__main__':
    main()
