import pandas as pd
from tqdm import tqdm

# Wczytanie wartości od użytkownika
max_diff_litry = input("Podaj maksymalną dopuszczalną różnicę dla kolumny [Litry] (w litrach): ")
max_diff_wartosc = input("Podaj maksymalną dopuszczalną różnicę dla kolumny [Wartość] (w złotówkach): ")
max_diff_litry = float(max_diff_litry)
max_diff_wartosc = float(max_diff_wartosc)

# Odczyt plików
print("Wczytuję pliki...")
df_abs = pd.read_excel('abs.xlsx')
df_obi = pd.read_excel('obi.xlsx')

# Dataframes wynikowe
df_missing_abs = pd.DataFrame(columns=df_abs.columns)
df_missing_obi = pd.DataFrame(columns=df_obi.columns)
df_diff = pd.DataFrame(columns=['Numer', 'Ilość_abs', 'Litry_abs', 'Wartość_abs', 'Ilość_obi', 'Litry_obi', 'Wartość_obi'])

# Zadanie 1
for num in tqdm(df_abs['Numer'], desc="Przeszukuje plik ABS"):
    if num not in df_obi['Numer'].values:
        df_missing_obi = pd.concat([df_missing_obi, df_abs[df_abs['Numer'] == num]], ignore_index=True)

# Zadanie 2
for num in tqdm(df_obi['Numer'], desc="Przeszukuję plik OBI"):
    if num not in df_abs['Numer'].values:
        df_missing_abs = pd.concat([df_missing_abs, df_obi[df_obi['Numer'] == num]], ignore_index=True)

# Zadanie 3
# Szukanie nieprawidłowych wartości i dodawanie ich do ramki danych
common_nums = set(df_abs['Numer']).intersection(set(df_obi['Numer']))
data = []

for num in tqdm(common_nums, desc="Porównuję dane dla obu plików"):
    row_abs = df_abs[df_abs['Numer'] == num].iloc[0]
    row_obi = df_obi[df_obi['Numer'] == num].iloc[0]
    if not all(row_abs[['Ilość', 'Litry', 'Wartość']].values == row_obi[['Ilość', 'Litry', 'Wartość']].values) \
            and abs(row_abs['Litry'] - row_obi['Litry']) > max_diff_litry \
            or abs(row_abs['Wartość'] - row_obi['Wartość']) > max_diff_wartosc:
        data.append((num, row_abs['Ilość'], row_obi['Ilość'], row_abs['Litry'], row_obi['Litry'], row_abs['Wartość'], row_obi['Wartość']))
df_diff = pd.DataFrame(data, columns=['Numer', 'Ilość_abs', 'Litry_abs', 'Wartość_abs', 'Ilość_obi', 'Litry_obi', 'Wartość_obi'])

# Zapis wyników do pliku output.xlsx
with pd.ExcelWriter('output.xlsx') as writer:
    df_missing_abs.to_excel(writer, sheet_name='Nie znaleziono w ABS')
    df_missing_obi.to_excel(writer, sheet_name='Nie znaleziono w OBI')
    df_diff.to_excel(writer, sheet_name='Różne wartości')
