def add_or_update_item(item_list: list, new_item: dict):
    for item in item_list:
        if item['наименование'] == new_item['наименование']:
            item['количество_пришло'] += new_item['количество_пришло']
            item['остаток_на_начало_периода'] += new_item['остаток_на_начало_периода']
            return item_list
    item_list.append(new_item.copy())
    return item_list
