# Exercise 1
# Lọc ra các từ có ít nhất 2 nguyên âm (EUOAI) từ sheet `Excercise1` trong file `ueoai.xlsx` và ghi vào sheet `Filtered` trong cùng file.

import pandas as pd

def has_at_least_two_vowels(word):
    vowels = 'ueoai'
    count = sum(1 for char in str(word).lower() if char in vowels)
    return count >=2

def write_to_excel(file_path, sheet_name, sheet_new):
    df = pd.read_excel(file_path, sheet_name= sheet_name, header= None)
    filtered_df = df[df.iloc[:,0].apply(has_at_least_two_vowels)]
    with pd.ExcelWriter(file_path, engine= 'openpyxl', mode='a', if_sheet_exists='replace') as writer:
        filtered_df.to_excel(writer, sheet_name= sheet_new, index= False, header= False)

def main():
    file_path = '/Users/admin/Documents/Python/python/euoai.xlsx'
    sheet_name = 'Excercise1'
    sheet_new = 'Filtered'
    write_to_excel(file_path,sheet_name, sheet_new)

if __name__ == "__main__":
    main()