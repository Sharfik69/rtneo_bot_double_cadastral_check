from openpyxl import load_workbook, Workbook

name = 'Аларский синхронизированные (2)'

wb = load_workbook('{}.xlsx'.format(name))
s = wb['выгрузка']

wb1 = Workbook()
ss = wb1.active

a = set()
d = {}

for i in range(2, 400000):
    if s.cell(row=i, column=1).value is None:
        break
    kn = s.cell(row=i, column=10).value
    if kn != None and kn[0].isdigit():
        info = []

        for j in range(1, 40):
            info.append(s.cell(row=i, column=j).value)

        if kn not in d:
            d[kn] = []

        d[kn].append(info)

row_ = 1
ans = 0

for kn in d:
    if len(d[kn]) == 1:
        continue

    street_set, house_set, apart_set = set(), set(), set()

    for j in d[kn]:
        street_set.add(j[4].lower().replace('-й', ''))
        house_set.add(j[5])
        apart_set.add(j[6])

    if len(street_set) == 1 and len(house_set) == 1 and len(apart_set) == 1:
        continue

    ans += 1
    for j in range(len(d[kn])):
        for col_ in range(len(d[kn][j])):
            ss.cell(row=row_, column=col_ + 1).value = d[kn][j][col_]
        row_ += 1
    row_ += 1

ss = wb1.create_sheet("список кадастров")

row_ = 1
for i in d.keys():
    if len(d[i]) == 1:
        continue
    ss.cell(row=row_, column=1).value = i
    row_ += 1

wb1.save('{} проверка.xlsx'.format(name))
print(ans)
