# app_cuaca_svm.py

import numpy as np
import requests
import joblib
import streamlit as st


API_URL_TEMPLATE = "https://api.bmkg.go.id/publik/prakiraan-cuaca?adm4={adm4}"


@st.cache_resource
def load_model_and_scaler():
    model = joblib.load("bmkg_svm_model.pkl")
    scaler = joblib.load("bmkg_svm_scaler.pkl")
    return model, scaler


def get_one_forecast_from_bmkg(adm4_code: str):
    """
    Ambil satu record prakiraan cuaca dari BMKG untuk adm4 tertentu.
    Di sini untuk simpel: ambil entry pertama saja.
    Return: dict dengan t, hu, ws, tcc, local_datetime, lokasi.
    """
    url = API_URL_TEMPLATE.format(adm4=adm4_code)
    resp = requests.get(url, timeout=15)
    resp.raise_for_status()
    data = resp.json()

    lokasi = data.get("lokasi", {})
    provinsi = lokasi.get("provinsi")
    kotkab = lokasi.get("kotkab")
    kecamatan = lokasi.get("kecamatan")
    desa = lokasi.get("desa")

    cuaca_all = data.get("data", [])
    if not cuaca_all:
        raise ValueError("Field 'data' kosong.")

    cuaca_harian = cuaca_all[0].get("cuaca", [])
    if not cuaca_harian:
        raise ValueError("Field 'cuaca' kosong.")

    # Ambil entry pertama dari hari pertama sebagai contoh
    first_day_list = cuaca_harian[0]
    if not first_day_list:
        raise ValueError("List prakiraan hari pertama kosong.")

    entry = first_day_list[0]

    result = {
        "t": entry.get("t"),
        "hu": entry.get("hu"),
        "ws": entry.get("ws"),
        "tcc": entry.get("tcc"),
        "local_datetime": entry.get("local_datetime"),
        "weather_desc": entry.get("weather_desc"),
        "provinsi": provinsi,
        "kotkab": kotkab,
        "kecamatan": kecamatan,
        "desa": desa,
    }
    return result


def main():
    st.title("Aplikasi Prediksi Kategori Cuaca (BMKG + SVM)")
    st.write(
        "Aplikasi ini menggunakan data prakiraan cuaca BMKG "
        "dan algoritma **Support Vector Machine (SVM)** untuk "
        "memprediksi apakah cuaca **Hujan** atau **Tidak Hujan**."
    )
    st.caption("Sumber data: BMKG (https://data.bmkg.go.id/prakiraan-cuaca/)")

    try:
        model, scaler = load_model_and_scaler()
    except Exception as e:
        st.error(
            "Gagal memuat model/scaler. Pastikan sudah menjalankan bmkg_train_svm.py. "
            f"Detail error: {e}"
        )
        return

    mode = st.radio(
        "Pilih mode input:",
        ["Input manual", "Ambil dari BMKG (kode adm4)"],
    )

    if mode == "Input manual":
        st.subheader("Input Manual Parameter Cuaca")

        t = st.slider("Suhu (°C)", 0.0, 40.0, 28.0, 0.5)
        hu = st.slider("Kelembapan (%)", 0.0, 100.0, 80.0, 1.0)
        ws = st.slider("Kecepatan Angin (km/jam)", 0.0, 50.0, 5.0, 0.5)
        tcc = st.slider("Tutupan Awan (%)", 0.0, 100.0, 80.0, 1.0)

        if st.button("Prediksi Cuaca"):
            features = np.array([[t, hu, ws, tcc]])
            features_scaled = scaler.transform(features)
            pred = model.predict(features_scaled)[0]
            proba = model.predict_proba(features_scaled)[0]

            label = "Hujan" if pred == 1 else "Tidak Hujan"
            prob = proba[1] if pred == 1 else proba[0]

            st.subheader(f"Hasil Prediksi: **{label}**")
            st.write(f"Probabilitas: **{prob:.2f}**")

    else:
        st.subheader("Ambil Data dari BMKG (adm4)")

        st.write(
            "Masukkan kode wilayah **adm4** BMKG. "
            "Contoh: `31.71.01.1001` (Gambir, Jakarta Pusat) atau `31.71.03.1001` (Kemayoran)."
            "`32.75.05.1001` #Bojong Rawalumbu, `32.75.05.1002` #Pengasinan, `32.75.05.1003` #Sepanjangjaya, `32.75.05.1004` #Bojongmenteng"
        )

        adm4_code = st.text_input("Kode adm4", value="31.71.01.1001")

        if st.button("Ambil dari BMKG & Prediksi"):
            try:
                data = get_one_forecast_from_bmkg(adm4_code)
                st.success("Data berhasil diambil dari BMKG.")

                st.write("**Lokasi:**")
                st.write(
                    f"{data.get('desa')}, {data.get('kecamatan')}, "
                    f"{data.get('kotkab')}, {data.get('provinsi')}"
                )
                st.write(f"**Waktu lokal prakiraan:** {data.get('local_datetime')}")
                st.write(f"**Deskripsi cuaca BMKG:** {data.get('weather_desc')}")

                t = data.get("t")
                hu = data.get("hu")
                ws = data.get("ws")
                tcc = data.get("tcc")

                st.write("---")
                st.write("**Parameter yang dipakai model:**")
                st.write(f"Suhu (t): {t} °C")
                st.write(f"Kelembapan (hu): {hu} %")
                st.write(f"Kecepatan Angin (ws): {ws} km/jam")
                st.write(f"Tutupan Awan (tcc): {tcc} %")

                features = np.array([[t, hu, ws, tcc]])
                features_scaled = scaler.transform(features)
                pred = model.predict(features_scaled)[0]
                proba = model.predict_proba(features_scaled)[0]

                label = "Hujan" if pred == 1 else "Tidak Hujan"
                prob = proba[1] if pred == 1 else proba[0]

                st.subheader(f"Hasil Prediksi Model: **{label}**")
                st.write(f"Probabilitas (versi model): **{prob:.2f}**")

                st.info(
                    "Catatan: BMKG sudah memberikan prakiraan resmi. "
                    "Model ini hanya contoh pembelajaran machine learning "
                    "yang meniru pola dari data prakiraan BMKG."
                )

            except Exception as e:
                st.error(f"Gagal mengambil / memproses data dari BMKG: {e}")


if __name__ == "__main__":
    main()
