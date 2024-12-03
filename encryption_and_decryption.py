import pandas as pd
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import padding
import os
import base64
import datetime as dt
from dateutil import parser

key = os.urandom(32) 
iv = os.urandom(16)  

cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())

def encrypt_data(data):
    if not isinstance(data, str):
        data = str(data)
    padder = padding.PKCS7(128).padder()
    padded_data = padder.update(data.encode()) + padder.finalize()
    encryptor = cipher.encryptor()
    return encryptor.update(padded_data) + encryptor.finalize()

def decrypt_data(data):
    decryptor = cipher.decryptor()
    decrypted_padded_data = decryptor.update(data) + decryptor.finalize()
    unpadder = padding.PKCS7(128).unpadder()
    return unpadder.update(decrypted_padded_data) + unpadder.finalize()

def save_key_and_iv(key, iv, key_path):
    encrypted_key = base64.b64encode(key).decode('utf-8')
    encrypted_iv = base64.b64encode(iv).decode('utf-8')
    with open(key_path, 'w') as key_file:
        key_file.write(f"Key: {encrypted_key}\nIV: {encrypted_iv}\n")

def read_and_encrypt_excel(input_file_path, output_file_path, columns_to_encrypt):
    df = pd.read_excel(input_file_path)
    for column in columns_to_encrypt:
        if column in df.columns:
            df[column] = df[column].apply(lambda x: encrypt_data(x).hex())
    df.to_excel(output_file_path, index=False)

def time_extraction(input_file_path, output_file_path, date_columns):
    df = pd.read_excel(input_file_path)
    for column in date_columns:
        if column in df.columns:
            df[column] = df[column].apply(lambda x: parser.parse(str(x)).strftime('%m-%Y') if pd.notnull(x) and isinstance(x, str) else x)
    df.to_excel(output_file_path, index=False)

def masking_unrelated_data(input_file_path, output_file_path, unrelated_data):
    df = pd.read_excel(input_file_path)
    for column in unrelated_data:
        if column in df.columns:
            df[column] = '*'
    df.to_excel(output_file_path, index=False)

def date_shifting(input_file_path, output_file_path, year_columns):
    df = pd.read_excel(input_file_path)
    for column in year_columns:
        if column in df.columns:
            year = df[column]
            df[column] = (year // 10) * 10
    df.to_excel(output_file_path, index=False)

input_file_paths = ['Classified_Tables//case_visit.xlsx',
                    'Classified_Tables//demo_patient.xlsx',
                    'Classified_Tables//lab_results.xlsx',
                    'Classified_Tables//medication_order.xlsx',
                    'Classified_Tables//surgical_patient.xlsx']

output_file_paths = ['Encrypted_Tables//encrypt_case_visit.xlsx', 
                    'Encrypted_Tables//encrypt_demo_patient.xlsx', 
                    'Encrypted_Tables//encrypt_lab_results.xlsx',
                    'Encrypted_Tables//encrypt_medication_order.xlsx',
                    'Encrypted_Tables//encrypt_surgical_patient.xlsx']

key_path = 'key_collection//encrypt_key.txt'
columns_to_encrypt = ['Patient_No', 
                    'Case_No',
                    'Surgical_Code',
                    ]
date_columns = ['Order_Creation_Date', 'operation_date','Result_Test_Date','Adm_DateTime','Dis_DateTime','Dis_DateTime_Planned']
unrelated_data=['Religion','Marital_Status','Operation_Table_Code']
year_columns=['Patient_DOB','Death_Date']

for input_path, output_path in zip(input_file_paths, output_file_paths):
    time_extraction(input_path, output_path, date_columns)
    masking_unrelated_data(output_path, output_path, unrelated_data)
    date_shifting(output_path, output_path, year_columns)
    read_and_encrypt_excel(output_path, output_path, columns_to_encrypt)

save_key_and_iv(key, iv, key_path)

print("all encrypted.key saved.")