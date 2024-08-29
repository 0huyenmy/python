# Exercise 1
# Lọc ra cac valid email từ sheet `Email` trong file `email.xlsx` và in ra.
# Với


import pandas as pd
def filter_valid_emails(file_path, sheet_name, pattern):
    df = pd.read_excel(file_path, sheet_name= sheet_name, header= None)

    valid_emails = df[df[0].str.contains(pattern, regex= True, na=False)]
    print("Valid Emails:")
    for email in valid_emails[0]:
        print(email)

def main():
    file_path = '/Users/admin/Documents/Python/python/email.xlsx'  
    sheet_name = 'Email'     
    pattern ='[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,4}'
    
    filter_valid_emails(file_path, sheet_name, pattern)

if __name__ == "__main__":
    main()