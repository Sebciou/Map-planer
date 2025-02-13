import streamlit as st
import pandas as pd
import requests
import folium
from streamlit_folium import st_folium
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side  # Dodano do formatowania obrysów komórek

# Klucz API Google Geocoding i Directions
GOOGLE_API_KEY = "AIzaSyDkWOkJfwOHwZf83KNv-u-DmcDgnNslU9Q"  # Tutaj wklej swój klucz API

# Funkcja do geokodowania adresów za pomocą Google Geocoding API
def geocode_address(address):
    url = f"https://maps.googleapis.com/maps/api/geocode/json?address={address}&key={GOOGLE_API_KEY}"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            if data["status"] == "OK":
                location = data["results"][0]["geometry"]["location"]
                return location["lat"], location["lng"]
            else:
                st.warning(f"Błąd geokodowania dla adresu: {address}. Status: {data['status']}")
                return None, None
        else:
            st.warning(f"Błąd połączenia z API. Kod statusu: {response.status_code}")
            return None, None
    except Exception as e:
        st.warning(f"Wystąpił błąd: {e}")
        return None, None

# Funkcja do obliczania czasu przejazdu między punktami
def calculate_travel_time(origin, destination):
    url = f"https://maps.googleapis.com/maps/api/directions/json?origin={origin}&destination={destination}&key={GOOGLE_API_KEY}"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            if data["status"] == "OK":
                # Pobierz czas przejazdu i odległość
                duration = data["routes"][0]["legs"][0]["duration"]["text"]
                distance = data["routes"][0]["legs"][0]["distance"]["text"]
                return duration, distance
            else:
                st.warning(f"Błąd obliczania trasy. Status: {data['status']}")
        else:
            st.warning(f"Błąd połączenia z API. Kod statusu: {response.status_code}")
    except Exception as e:
        st.warning(f"Wystąpił błąd: {e}")
    return None, None

# Interfejs Streamlit
st.title("Mapowanie adresów i eksport do Excela")

# Dodaj logo w nagłówku
st.image("logo.png", width=200)  # Zmień "logo.png" na ścieżkę do swojego logo

# Inicjalizacja stanu aplikacji
if 'selected_points' not in st.session_state:
    st.session_state.selected_points = []

# Wczytaj plik Excel
uploaded_file = st.file_uploader("Wgraj plik Excel z adresami", type=["xlsx"])
if uploaded_file:
    # Wczytaj dane z Excela
    df = pd.read_excel(uploaded_file)
    st.write("Wczytane dane:")
    st.write(df)

    # Sprawdź, czy plik ma wymagane kolumny
    required_columns = ["Ulica", "Miasto", "Kod pocztowy", "Nazwa PM", "PLA", "Opis", "Kategoria"]
    if not all(column in df.columns for column in required_columns):
        st.error("Plik Excel musi zawierać kolumny: Ulica, Miasto, Kod pocztowy, Nazwa PM, PLA, Opis, Kategoria.")
    else:
        # Geokodowanie adresów
        st.write("Geokodowanie adresów...")
        df['Adres'] = df['Ulica'] + ", " + df['Miasto'] + ", " + df['Kod pocztowy']
        df['Współrzędne'] = df['Adres'].apply(geocode_address)
        df[['Lat', 'Lon']] = pd.DataFrame(df['Współrzędne'].tolist(), index=df.index)

        # Lista adresów, których nie udało się zgeokodować
        failed_geocoding = df[df['Lat'].isna() | df['Lon'].isna()]
        if not failed_geocoding.empty:
            st.write("### Adresy, których nie udało się zgeokodować:")
            st.write(failed_geocoding[['Adres']])

        # Usuń adresy, których nie udało się zgeokodować
        df = df.dropna(subset=['Lat', 'Lon'])

        # Podziel ekran na dwie kolumny: mapa (2/3 szerokości) i panel boczny (1/3 szerokości)
        col1, col2 = st.columns([2, 1])

        # Wyświetl mapę w lewej kolumnie (większa mapa)
        with col1:
            st.write("Mapa z adresami:")
            m = folium.Map(location=[df['Lat'].mean(), df['Lon'].mean()], zoom_start=6)

            # Definiuj kolory dla kategorii
            category_colors = {
                "Serwis": "blue",
                "Montaż": "green",
                "Reklamacja": "orange",
                "Inne": "purple"
            }

            # Dodaj punkty do mapy z identyfikatorami
            for idx, row in df.iterrows():
                # Sprawdź, czy punkt jest zaznaczony
                if row['Nazwa PM'] in st.session_state.selected_points:
                    marker_color = 'red'  # Zaznaczony punkt ma kolor czerwony
                    # Pobierz numer kolejności zaznaczenia
                    index = st.session_state.selected_points.index(row['Nazwa PM']) + 1
                    tooltip_text = f"{index}. {row['Nazwa PM']}<br>{row['Opis']}"  # Dodaj numer i opis do tooltip
                else:
                    # Ustal kolor na podstawie kategorii
                    marker_color = category_colors.get(row['Kategoria'], "gray")  # Domyślny kolor: szary
                    tooltip_text = f"{row['Nazwa PM']}<br>{row['Opis']}"  # Dodaj opis do tooltip

                folium.Marker(
                    location=[row['Lat'], row['Lon']],
                    popup=f"{row['Nazwa PM']}<br>{row['Opis']}",  # Nazwa PM i opis jako popup
                    tooltip=folium.Tooltip(tooltip_text, permanent=False),  # Tooltip z numerem i opisem
                    icon=folium.Icon(color=marker_color)
                ).add_to(m)

            # Interaktywna mapa w Streamlit (większa szerokość)
            map_data = st_folium(m, width=900, height=500)  # Zwiększono szerokość mapy

        # Zaznaczanie punktów
        if map_data.get("last_object_clicked"):
            clicked_lat = map_data["last_object_clicked"]["lat"]
            clicked_lon = map_data["last_object_clicked"]["lng"]

            # Znajdź adres odpowiadający klikniętym współrzędnym
            selected_point = df[(df['Lat'] == clicked_lat) & (df['Lon'] == clicked_lon)]
            if not selected_point.empty:
                selected_pm = selected_point.iloc[0]['Nazwa PM']

                # Dodaj lub usuń zaznaczony punkt z listy
                if selected_pm not in st.session_state.selected_points:
                    st.session_state.selected_points.append(selected_pm)
                    st.write(f"Zaznaczono punkt: {selected_pm}")
                else:
                    st.session_state.selected_points.remove(selected_pm)
                    st.write(f"Odznaczono punkt: {selected_pm}")

                # Wymuś odświeżenie mapy
                st.rerun()

        # Wyświetl zaznaczone punkty i czas przejazdu w prawej kolumnie
        with col2:
            if st.session_state.selected_points:
                st.write("### Zaznaczone punkty w kolejności:")
                for i, pm in enumerate(st.session_state.selected_points, 1):
                    st.write(f"{i}. {pm}")

                # Oblicz czas przejazdu między punktami
                if len(st.session_state.selected_points) > 1:
                    st.write("### Czas przejazdu między punktami:")
                    for i in range(len(st.session_state.selected_points) - 1):
                        start_point = st.session_state.selected_points[i]
                        end_point = st.session_state.selected_points[i + 1]

                        # Pobierz współrzędne punktów
                        start_coords = df[df['Nazwa PM'] == start_point][['Lat', 'Lon']].values[0]
                        end_coords = df[df['Nazwa PM'] == end_point][['Lat', 'Lon']].values[0]

                        # Oblicz czas przejazdu
                        origin = f"{start_coords[0]},{start_coords[1]}"
                        destination = f"{end_coords[0]},{end_coords[1]}"
                        duration, distance = calculate_travel_time(origin, destination)

                        if duration and distance:
                            st.write(f"{i + 1}. {start_point} -> {end_point}: {duration} ({distance})")

                # Przygotuj dane do eksportu zgodnie z nowym szablonem
                selected_df = df[df['Nazwa PM'].isin(st.session_state.selected_points)]
                selected_df = selected_df.set_index('Nazwa PM').loc[st.session_state.selected_points].reset_index()

                # Utwórz nowy plik Excel z formułami i obrysami komórek
                wb = Workbook()
                ws = wb.active

                # Dodaj nagłówki zgodnie z nowym szablonem
                headers = ["PLA", "HIPSEX", "Ultra", "Minato", "Kod pocztowy", "Nazwa PM", "Kategoria"]
                ws.append(headers)

                # Dodaj dane
                for _, row in selected_df.iterrows():
                    ws.append([row["PLA"], "", "", "", row["Kod pocztowy"], row["Nazwa PM"], row["Kategoria"]])

                # Dodaj obrysy komórek
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                for row in ws.iter_rows(min_row=2, max_row=len(selected_df) + 1, min_col=1, max_col=7):
                    for cell in row:
                        cell.border = thin_border  # Dodaj obrysy komórek

                # Zapisz plik Excel do bufora
                output = BytesIO()
                wb.save(output)
                output.seek(0)

                # Przycisk do pobrania pliku Excel
                st.download_button(
                    label="Eksportuj do Excela",
                    data=output,
                    file_name="wybrane_punkty.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.write("Nie zaznaczono jeszcze żadnych punktów.")
