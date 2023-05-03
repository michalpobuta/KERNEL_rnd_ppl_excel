# This is a sample Python script.
import datetime
import math
import os
import random

import pandas as pd


def przetworz_plik_excel(nazwa_pliku):
    df = pd.read_excel(nazwa_pliku)

    wynik = []
    firmy = set()
    pozycje = ["1st", "2nd", "3rd", "4th", "5th", "6th"]
    # Zbieranie unikalnych nazw firm
    for pos in pozycje:
        for col in df.columns:
            if f"[{pos}]" in col:
                unikalne_firmy = df[col].dropna().unique()
                firmy.update(unikalne_firmy)

    print(firmy)
    # Tworzenie słownika wynikowego dla każdej firmy
    for firma in firmy:
        firma_dict = {firma: []}
        for pos in pozycje:
            pos_dict = {pos: []}

            for col in df.columns:
                if f"[{pos}]" in col:
                    mask = df[col] == firma
                    indeksy = df[mask]['indeks'].tolist()
                    pos_dict[pos].extend(indeksy)

            firma_dict[firma].append(pos_dict)

        wynik.append(firma_dict)

    return wynik

def tworz_pliki_excel(tablica_slownikow):
    nazwy_plikow = {}
    for slownik in tablica_slownikow:
        for nazwa_firmy, wartosci in slownik.items():
            godzina_start = datetime.datetime.strptime(wartosci[0], '%H:%M')
            godzina_koniec = datetime.datetime.strptime(wartosci[1], '%H:%M')
            ilosc_minut = int(wartosci[2])

            # Tworzenie nagłówka DataFrame
            df = pd.DataFrame(columns=['Godzina', nazwa_firmy])

            # Wypełnianie kolumny "Godzina"
            obecna_godzina = godzina_start
            while obecna_godzina <= godzina_koniec:
                df = df._append({'Godzina': obecna_godzina.strftime('%H:%M')}, ignore_index=True)

                obecna_godzina += datetime.timedelta(minutes=ilosc_minut)

            # Usunięcie istniejącego pliku Excel, jeśli istnieje
            nazwa_pliku = f"{nazwa_firmy}.xlsx"
            if os.path.exists(nazwa_pliku):
                os.remove(nazwa_pliku)

            # Zapis DataFrame do pliku Excel
            df.to_excel(nazwa_pliku, index=False)
            nazwy_plikow[nazwa_firmy] = nazwa_pliku
    return nazwy_plikow


def getslowniki():
    start = '18:50'
    tablica_slownikow = [
        {'Nokia C++': [start, '21:30', 10]},
        {'Nokia DevOps': [start, '20:30', 10]},
        {'Nokia Tester': [start, '20:30', 10]},
        {'Hitachi - QA': [start, '21:30', 5]},
        {'Deployed': [start, '21:30', 5]},
        {'Hitachi - Front, FPGA': [start, '21:30', 5]}

    ]
    return tablica_slownikow

def wpisz_uczestnikow(nazwy_plikow, wynik_przetworz_plik_excel):
    # Odczytanie plików Excel i zapisanie ich jako DataFrame w słowniku
    dataframes = {}
    for nazwa_firmy, nazwa_pliku in nazwy_plikow.items():
        dataframes[nazwa_firmy] = pd.read_excel(nazwa_pliku)

    # Utworzenie DataFrame do przechowywania indeksów wpisanych osób i ich przedziałów czasowych
    wpisane_osoby = pd.DataFrame(columns=['indeks', 'start', 'koniec'])
    wpisane_osoby['start'].apply(lambda x: datetime.datetime.strptime(x,'%H:%M').time())
    wpisane_osoby['koniec'].apply(lambda x: datetime.datetime.strptime(x,'%H:%M').time())
    # Iteracja po tablicy plików
    for nazwa_firmy, nazwa_pliku in nazwy_plikow.items():
        df = dataframes[nazwa_firmy]

        # Wyszukanie wartości w tablicy słowników utworzonej przez przetworz_plik_excel()
        for wynik in wynik_przetworz_plik_excel:
            if nazwa_firmy in wynik:
                pozycje = wynik[nazwa_firmy]

                # Wpisanie użytkowników do excela
                for pozycja in pozycje:
                    for priorytet, indeksy in pozycja.items():
                        random.shuffle(indeksy)
                        for indeks in indeksy:
                            wolne_miejsca = df[df[nazwa_firmy].isna() ].index.tolist()
                            # Sprawdzenie, czy użytkownik nie jest już na liście
                            if indeks not in df[nazwa_firmy].values:
                                if wolne_miejsca:
                                    wolne_miejsce = wolne_miejsca[0]
                                    start = df.at[wolne_miejsce, 'Godzina']
                                    # Pobranie czasu trwania spotkania z aktualnie iterowanego excela
                                    if (len(wolne_miejsca) != 1):
                                        czas_trwania_spotkan = (datetime.datetime.combine(datetime.datetime.min, datetime.time.fromisoformat(df.at[wolne_miejsce + 1, 'Godzina'])) - datetime.datetime.combine(datetime.datetime.min, datetime.time.fromisoformat(start))).seconds // 60
                                    else:
                                        czas_trwania_spotkan = (datetime.datetime.combine(datetime.datetime.min, datetime.time.fromisoformat(start)) - datetime.datetime.combine(datetime.datetime.min,datetime.time.fromisoformat(df.at[wolne_miejsce -1, 'Godzina']))).seconds  // 60
                                    koniec = (datetime.datetime.combine(datetime.datetime.min, datetime.time.fromisoformat(start)) + datetime.timedelta(minutes=czas_trwania_spotkan)).time()

                                    # Sprawdzenie kolizji czasowych
                                    kolizja = wpisane_osoby[(wpisane_osoby['indeks'] == indeks) &
                                                            (((wpisane_osoby['start'] >= datetime.datetime.strptime(start,'%H:%M').time()) & (wpisane_osoby['start'] < koniec)) |
                                                             ((wpisane_osoby['koniec'] > datetime.datetime.strptime(start,'%H:%M').time()) & (wpisane_osoby['koniec'] <= koniec)))]



                                    if kolizja.empty or len(wolne_miejsca) == 1:

                                        # Dodanie uczestnika do excela
                                        df.at[wolne_miejsce, nazwa_firmy] = indeks
                                        # Dodanie uczestnika do wpisane_osoby
                                        wpisane_osoby = wpisane_osoby._append({'indeks': indeks, 'start':  datetime.datetime.combine(datetime.datetime.min, datetime.time.fromisoformat(start)).time(), 'koniec': koniec}, ignore_index=True)
                                    else:
                                        print('kolizja użytkownika: ' + str(indeks) + ' ' + nazwa_firmy)
                                        # Przesunięcie uczestnika na inną godzinę
                                        for index, row in df.iterrows():
                                            if(index != 0):
                                                start = datetime.datetime.combine(datetime.datetime.min, datetime.time.fromisoformat(df.at[index, 'Godzina']))
                                                czas_trwania_spotkan = (datetime.datetime.combine(datetime.datetime.min, datetime.time.fromisoformat(df.at[index + 1, 'Godzina'])) - start).seconds // 60
                                                koniec = ( start+ datetime.timedelta(minutes=czas_trwania_spotkan)).time()
                                                kolizja = wpisane_osoby[(wpisane_osoby['indeks'] == indeks) &
                                                                (((wpisane_osoby['start'] >= start.time()) & (wpisane_osoby['start'] < koniec)) |
                                                                 ((wpisane_osoby['koniec'] > start.time()) & (wpisane_osoby['koniec'] <= koniec)))]

                                                if kolizja.empty and math.isnan(df.at[index, nazwa_firmy]) :
                                                    df.at[index, nazwa_firmy] = indeks
                                                    start = start.time()
                                                    wpisane_osoby = wpisane_osoby._append(
                                                        {'indeks': indeks, 'start': start, 'koniec': koniec},
                                                        ignore_index=True)
                                                    break
                                else:
                                    # Dodanie uczestnika do listy rezerwowej
                                    df = df._append({nazwa_firmy: indeks}, ignore_index=True)

                                            # Zapisanie zmodyfikowanego DataFrame do pliku Excel
        df.to_excel(nazwa_pliku, index=False)


    return nazwy_plikow


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    wynik = przetworz_plik_excel('test.xlsx')
    tab = tworz_pliki_excel(getslowniki())
    wpisz_uczestnikow(tab, wynik)
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
