import openpyxl


def concatenate_cells(first, rest):
    output = []
    for num in first:
        for line in rest:
            output.append([num] + line)
    return output


def main():
    path = input('Path to workbook: ')
    # path = 'test.xlsx'  # Testing purposes
    num_sheet_name = 'FID'
    rest_sheet_name = 'MVtasks'
    result_sheet_name = 'result'
    headers = 'y' in input('Are there headers? ').lower()

    print('Loading {}...'.format(path))
    wb = openpyxl.load_workbook(path, guess_types=True)
    num_sheet = wb[num_sheet_name]
    rest_sheet = wb[rest_sheet_name]
    result_sheet = wb.create_sheet(result_sheet_name)

    nums = [row[0].value for row in num_sheet.rows]  # Only need the first cell
    rest = [[cell.value for cell in row] for row in rest_sheet.rows]
    if headers:
        nums = nums[1:]
        rest = rest[1:]

    print('Concatenating cells...')
    result = concatenate_cells(nums, rest)

    print('Writing changes...')
    for row in result:
        result_sheet.append(row)

    wb.save(path)


if __name__ == '__main__':
    main()
