import os

try:
    old_file = "teste.txt"
    new_file = "teste123.txt"
    os.rename(old_file, new_file)
    print(f"O arquivo {old_file} foi renomeado para {new_file}")

except OSError as e:
    print(f"Erro: {e}")
