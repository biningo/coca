import os

import openpyxl

if __name__ == '__main__':
    base_dir = './excel'
    filenames = []
    for name in os.listdir(base_dir):
        filenames.append(name)
    filenames.sort()

    rows = []
    for name in filenames:
        sheet = openpyxl.load_workbook(os.path.join(base_dir, name))['Sheet1']
        rows.extend(sheet.rows)
    content = ''
    counter = 0
    for row in rows:
        counter += 1
        cols = [col.value for col in row]
        if len(cols) < 3:
            continue
        word = cols[0]
        ps = cols[1]
        translation = cols[2]
        content += '{}. **{}**\n{}\n{}\n\n'.format(counter, word, ps, translation)
        if counter % 20 == 0:
            part_dir = './md/{:02d}'.format(counter // 1000 + 1)
            if not os.path.exists(part_dir):
                os.mkdir(part_dir)
            with open(os.path.join(part_dir, "{:04d}.md".format(counter // 20)), 'w') as f:
                f.write(content)
                content = ''
