#!/home/b/projects/sari_store_prices/venv/bin/python3

def generateTagLocFile(sspDir):
    from openpyxl import load_workbook

    excel_file = load_workbook(sspDir + 'ssp.xlsx')
    sheet1 = excel_file.get_sheet_by_name('Sheet1')

    tag = []
    location = []

    i = 1
    while sheet1['A' + str(i+1)].value != None:
        current_tag = sheet1['C' + str(i+1)].value
        current_location = sheet1['D' + str(i+1)].value

        if current_tag not in tag: tag.append(str(current_tag))
        if current_location not in location: location.append(str(current_location))
        i += 1

    tag = sorted(tag)
    location = sorted(location)

    with open(sspDir + 'text_files/tags.txt', 'w') as tags:
        for t in tag:
            tags.write(f'{t}\n')
    with open(sspDir + 'text_files/locations.txt', 'w') as locations:
        for l in location:
            locations.write(f'{l}\n')

if __name__ == '__main__':
    generateTagLocFile('/home/b/projects/sari_store_prices/')