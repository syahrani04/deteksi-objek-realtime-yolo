# bmkg_train_svm.py

import requests
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC
from sklearn.metrics import (
    classification_report,
    confusion_matrix,
    accuracy_score,
)
import joblib


# =========================
# KONFIGURASI
# =========================
API_URL_TEMPLATE = "https://api.bmkg.go.id/publik/prakiraan-cuaca?adm4={adm4}"

# Beberapa contoh kode wilayah adm4 (bisa kamu tambah/ubah)
# 31.71.01.1001 = Gambir (Jakarta Pusat)
# 31.71.03.1001 = Kemayoran (Jakarta Pusat)

ADM4_CODES = [
    "32.75.05.1001",  # Bojong Rawalumbu
    "32.75.05.1002",  # Pengasinan
    "32.75.05.1003",  # Sepanjangjaya
    "32.75.05.1004",  # Bojongmenteng
]


# =========================
# FUNGSI BANTUAN
# =========================
def fetch_bmkg_forecast(adm4_code: str):
    """
    Ambil data prakiraan cuaca BMKG untuk satu kode adm4.
    Return: list of dict (record per waktu prakiraan).
    """
    url = API_URL_TEMPLATE.format(adm4=adm4_code)
    print(f"Mengambil data dari BMKG untuk adm4={adm4_code} ...")

    resp = requests.get(url, timeout=15)
    resp.raise_for_status()

    data = resp.json()

    # Info lokasi
    lokasi = data.get("lokasi", {})
    provinsi = lokasi.get("provinsi")
    kotkab = lokasi.get("kotkab")
    kecamatan = lokasi.get("kecamatan")
    desa = lokasi.get("desa")

    records = []

    # Struktur JSON: data[0]["cuaca"] adalah list [hari1, hari2, hari3]
    cuaca_all = data.get("data", [])
    if not cuaca_all:
        print(f"Tidak ada field 'data' untuk adm4={adm4_code}")
        return records

    cuaca_harian = cuaca_all[0].get("cuaca", [])
    if not cuaca_harian:
        print(f"Tidak ada field 'cuaca' untuk adm4={adm4_code}")
        return records

    for day_idx, day_list in enumerate(cuaca_harian, start=1):
        if not isinstance(day_list, list):
            continue

        for entry in day_list:
            record = {
                "adm4": adm4_code,
                "provinsi": provinsi,
                "kotkab": kotkab,
                "kecamatan": kecamatan,
                "desa": desa,
                "local_datetime": entry.get("local_datetime"),
                "utc_datetime": entry.get("utc_datetime"),
                "t": entry.get("t"),
                "hu": entry.get("hu"),
                "ws": entry.get("ws"),
                "tcc": entry.get("tcc"),
                "tp": entry.get("tp", 0.0),
                "weather_desc": entry.get("weather_desc"),
            }
            records.append(record)

    print(f"  → {len(records)} record diambil untuk adm4={adm4_code}")
    return records


def label_hujan_from_desc(desc: str) -> int:
    """
    Konversi deskripsi cuaca BMKG menjadi label biner:
    1 = Hujan (apapun yang mengandung kata 'hujan')
    0 = Tidak Hujan
    """
    if desc is None:
        return 0
    desc_lower = str(desc).lower()
    if "hujan" in desc_lower:
        return 1
    return 0


# =========================
# MAIN TRAINING PIPELINE
# =========================
def main():
    # 1. Ambil dan gabungkan data dari beberapa lokasi
    all_records = []

    for code in ADM4_CODES:
        try:
            recs = fetch_bmkg_forecast(code)
            all_records.extend(recs)
        except Exception as e:
            print(f"ERROR mengambil data adm4={code}: {e}")

    if not all_records:
        print("Tidak ada data yang berhasil diambil. Cek koneksi / ADM4_CODES.")
        return

    df = pd.DataFrame(all_records)
    print("\nBeberapa baris data:")
    print(df.head())

    # 2. Buat label hujan / tidak hujan
    df["label_hujan"] = df["weather_desc"].apply(label_hujan_from_desc)

    print("\nDistribusi label_hujan (0=Tidak Hujan, 1=Hujan):")
    print(df["label_hujan"].value_counts())

    # Simpan dataset mentah untuk laporan
    df.to_csv("bmkg_cuaca_dataset.csv", index=False, encoding="utf-8")
    print("\nDataset disimpan ke bmkg_cuaca_dataset.csv")

    # 3. Pilih fitur (X) dan label (y)
    feature_cols = ["t", "hu", "ws", "tcc"]
    X = df[feature_cols]
    y = df["label_hujan"]

    # Buang baris yang ada NaN di fitur/label
    mask_valid = X.notna().all(axis=1) & y.notna()
    X = X[mask_valid]
    y = y[mask_valid]

    if len(y.unique()) < 2:
        print("Label hanya punya 1 kelas. Tambah lokasi lain / data lain dulu.")
        return

    # 4. Split train/test
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
    )

    # 5. Scaling fitur
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    X_test_scaled = scaler.transform(X_test)

    # 6. Training model SVM
    model = SVC(kernel="rbf", probability=True, random_state=42)
    model.fit(X_train_scaled, y_train)

    # 7. Evaluasi
    y_pred = model.predict(X_test_scaled)

    print("\n=== CONFUSION MATRIX ===")
    print(confusion_matrix(y_test, y_pred))
    print("\n=== CLASSIFICATION REPORT ===")
    print(
        classification_report(
            y_test,
            y_pred,
            target_names=["Tidak Hujan", "Hujan"],
        )
    )
    print("Akurasi:", accuracy_score(y_test, y_pred))

    # 8. Simpan model dan scaler
    joblib.dump(model, "bmkg_svm_model.pkl")
    joblib.dump(scaler, "bmkg_svm_scaler.pkl")
    print("\nModel dan scaler berhasil disimpan:")
    print("  bmkg_svm_model.pkl")
    print("  bmkg_svm_scaler.pkl")
    print("\nJANGAN LUPA: Cantumkan BMKG sebagai sumber data di laporan.")


if __name__ == "__main__":
    main()
