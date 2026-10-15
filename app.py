import csv
import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title='Konverter e-Faktur & Data Barang', page_icon='📊', layout='centered'
)

st.title('📊 Aplikasi Konversi e-Faktur & Data Barang')
st.write(
    'Muat naik fail CSV e-Faktur (`Data 1`) dan fail rujukan (`DATA BARANG`) di bawah untuk memproses penukaran kod, satuan, dan harga.'
)

# Bahagian muat naik fail (File Uploader)
uploaded_csv = st.file_uploader(
    '1. Muat naik Fail CSV e-Faktur (Data 1)', type=['csv']
)
uploaded_excel = st.file_uploader(
    '2. Muat naik Fail Rujukan DATA BARANG (Excel)', type=['xlsx']
)

if uploaded_csv is not None and uploaded_excel is not None:
  if st.button('Proses Data'):
    with st.spinner('Sedang memproses data... Sila tunggu sebentar.'):
      try:
        # Baca fail rujukan DATA BARANG dari muat naik
        df_barang = pd.read_excel(uploaded_excel, sheet_name=0)

        barang_dict = {}
        for _, row in df_barang.iterrows():
          nama = str(row['NAMA BARANG']).strip()
          kode = str(row['KODE BARANG'])
          karton = float(row['SATUAN PER KARTON'])
          barang_dict[nama] = {'KODE': kode, 'KARTON': karton}

        # Baca fail CSV e-Faktur dari muat naik
        # Streamlit file uploader mengembalikan bytes, jadi kita decode ke string
        stringio = io.StringIO(uploaded_csv.getvalue().decode('latin1'))
        reader = csv.reader(stringio)

        faktur_header = None
        gabungan_data = []

        for row in reader:
          if not row:
            continue

          jenis_baris = row[0]

          if jenis_baris == 'FK':
            faktur_header = row
          elif jenis_baris == 'OF':
            if faktur_header:
              kode_objek_lama = row[1]
              nama_barang = row[2].strip()
              harga_satuan_lama = float(row[3])
              jumlah_barang_lama = float(row[4])
              harga_total = float(row[5])
              diskon = float(row[6]) if row[6] else 0.0
              dpp = float(row[7]) if row[7] else 0.0
              ppn = float(row[8]) if row[8] else 0.0
              tarif_ppnbm = row[9]
              ppnbm = row[10] if len(row) > 10 and row[10] else 0.0

              # Semak padanan nama barang
              if nama_barang in barang_dict:
                kode_baru = barang_dict[nama_barang]['KODE']
                satuan_karton = barang_dict[nama_barang]['KARTON']
              else:
                kode_baru = kode_objek_lama
                satuan_karton = 1

              # Peraturan harga > 120,000
              if harga_satuan_lama > 120000 and satuan_karton > 0:
                harga_satuan_baru = harga_satuan_lama / satuan_karton
                jumlah_barang_baru = jumlah_barang_lama * satuan_karton
              else:
                harga_satuan_baru = harga_satuan_lama
                jumlah_barang_baru = jumlah_barang_lama

              row_baru_of = [
                  'OF',
                  str(kode_baru),
                  nama_barang,
                  f'{harga_satuan_baru:.2f}',
                  str(int(jumlah_barang_baru)),
                  f'{harga_total:.2f}',
                  f'{diskon:.2f}',
                  f'{dpp:.2f}',
                  f'{ppn:.2f}',
                  str(tarif_ppnbm),
                  str(ppnbm),
              ]

              gabungan_row = faktur_header + row_baru_of
              gabungan_data.append(gabungan_row)

        # Buat fail Excel dalam memori (BytesIO)
        output = io.BytesIO()
        df_output = pd.DataFrame(gabungan_data)
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
          df_output.to_excel(writer, index=False, header=False)
        processed_data = output.getvalue()

        st.success('Proses konversi berjaya dilakukan!')

        # Butang muat turun hasil
        st.download_button(
            label='📥 Muat Turun Fail Excel Hasil',
            data=processed_data,
            file_name='Hasil_Konversi_Faktur.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

      except Exception as e:
        st.error(f'Terjadi ralat semasa memproses fail: {e}')
else:
  st.info('Sila muat naik kedua-dua fail di atas untuk memulakan.')
