import json
import os

all_list = []


def add_data_01(ke1, ke2, ke3, ke4, ke5):
    using_data = add_json_data(ke1, ke2, ke3, ke4, ke5)
    creat_json_file(using_data)


def modify_file(ke1, ke2, ke3, ke4, ke5):
    using_data = add_json_data(ke1, ke2, ke3, ke4, ke5)
    old_real_filename = f"{ke1}{ke2}.json"  # 添加了扩展名
    modify_json_file(old_real_filename, using_data)


def remove_data_02(ke1, ke2, ke3, ke4, ke5):
    real_filename = f"{ke1}{ke2}.json"  # 添加了扩展名
    delete_json_file(real_filename)


def add_json_data(key1, key2, key3, key4, key5):
    data = {
        "矿物序号": str(key1),
        "物品编号": str(key2),
        "持有人": key3,
        "采集地": key4,
        "薄片描述": key5,
        "real_file_name": f"{key1}{key2}.json"  # 关键修正：添加扩展名
    }
    return data


def creat_json_file(data):
    os.makedirs('for_json', exist_ok=True)
    filename = data["real_file_name"]
    all_list.append(filename)
    filepath = os.path.join('for_json', filename)
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    return filepath


def modify_json_file(old_filename, data):
    new_filename = data["real_file_name"]
    old_path = os.path.join('for_json', old_filename)
    new_path = os.path.join('for_json', new_filename)

    if not os.path.exists(old_path):
        raise FileNotFoundError(f"文件 {old_filename} 不存在")

    with open(old_path, 'r', encoding='utf-8') as f:
        file_data = json.load(f)

    file_data.update(data)

    if new_filename != old_filename:
        with open(new_path, 'w', encoding='utf-8') as f:
            json.dump(file_data, f, indent=4, ensure_ascii=False)
        os.remove(old_path)
        if old_filename in all_list:
            all_list.remove(old_filename)
            all_list.append(new_filename)
    else:
        with open(old_path, 'w', encoding='utf-8') as f:
            json.dump(file_data, f, indent=4, ensure_ascii=False)

    return new_path


def delete_json_file(filename):
    filepath = os.path.join('for_json', filename)
    if os.path.exists(filepath):
        os.remove(filepath)
        if filename in all_list:
            all_list.remove(filename)
        return True
    return False
