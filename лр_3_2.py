def find_common_participants(str1, str2, sep=','):
    return sorted(set(str1.split(sep)).intersection(str2.split(sep)))


participants_first_group = "Иванов|Петров|Сидоров"
participants_second_group = "Петров|Сидоров|Смирнов"

print(find_common_participants(participants_first_group, participants_second_group, sep='|'))
