import pandas as pd
import numpy as np
from datetime import datetime

excel_file = 'Baseline Soal Kaji Selidik DrDPH Research (Responses).xlsx'
df_master = pd.read_excel(excel_file, sheet_name='Master List Baseline')

header_mapping = {
    "Timestamp": "Timestamp",
    "Email Address": "Email Address",
    "No Kad Pengenalan \nIdentification Number": "IC Number",
    "Plaque Score ": "Plaque Score",
    "Gingival Score ": "Gingival Score",
    "Jantina\nGender ": "Gender",
    "Bangsa \nRace": "Race",
    "Sekolah \nSchool ": "School",
    "Andakah anda mempunyai telefon pintar Android?\nDo you have an Android smartphone?": "Android User Check",
    "K1. Pengambilan makanan / minuman bergula terlalu banyak menyebabkan gigi berlubang \nTaking too much sugary foods / drinks can cause tooth decay": "K1",
    "K2. Pengambilan makanan / minuman bergula melebihi 4 kali sehari menyebabkan gigi berlubang \nTaking sugary foods  / drinks more than 4 times a day causes tooth decay": "K2",
    "K3. Seseorang boleh mengurangkan risiko gigi berlubang dengan mengurangkan pengambilan makanan/minuman bergula setiap hari.\nA person can reduce the risk of tooth decay by reducing sugary food/drinks intake every day ": "K3",
    "K4. Memberus gigi menggunakan ubat gigi berfluorida tidak berkesan untuk mencegah gigi berlubang.\nBrushing teeth with fluoridated toothpaste is not effective to prevent tooth decay ": "K4",
    "K5. Untuk keperluan fluorida yang mencukupi, kita perlu menggosok gigi sekurang-kurangnya dua kali sehari\nFor adequate fluoride supply, we need to brush our teeth at least twice a day": "K5",
    "K6. Gusi berdarah semasa memberus gigi adalah tanda penyakit gusi\nBleeding gum during tooth brushing is a sign of gum disease": "K6",
    "K7. Plak gigi menyebabkan penyakit gusi\nDental plaque can cause gum disease": "K7",
    "K8. Penyakit gusi boleh menyebabkan gigi menjadi longgar / goyang\nGum disease can cause teeth to become loose": "K8",
    "K9. Memberus gigi dengan teknik yang betul meningkatkan kesihatan gusi\nBrushing teeth with the correct technique improves the health of my gum": "K9",        
    "K10. Menggunakan flos gigi untuk membersihkan kawasan di celah-celah gigi meningkatkan kesihatan gusi\nUsing dental floss to clean the areas between teeth improves gum health": "K10",
    "K11. Gigi berlubang menyebabkan gigi sakit semasa makan\nDental cavity can cause pain during eating": "K11",
    "K12. Merokok boleh menyebabkan penyakit gusi.\nSmoking can cause gum disease": "K12",
    "K13. Posisi gigi yang bertindih, terpusing, menggigit bibir, atau jongang memerlukan rawatan pendakap gigi\nTeeth position which are overlapping, rotated, biting the lip, or jutted out will require braces treatment.": "K13",
    "K14. Pemutihan gigi adalah rawatan untuk memutihkan warna gigi\nTeeth whitening is a treatment to whiten the colour of teeth": "K14",
    "K15. Disyorkan untuk membersihkan celah-celah gigi menggunakan flos gigi sekali sehari.\nIt is recommended to clean the spaces between teeth using dental floss once a day.": "K15",
    "K16. Gigi yang tidak tersusun boleh mempengaruhi keyakinan diri.\nCrooked teeth can affect self confidence": "K16",
    "1a) ia mencegah gigi saya daripada berlubang \n1a) it prevents my teeth from decaying   ": "A1a",
    "1b) ia membuatkan nafas saya segar\n1b) it freshens my breath": "A1b",
    "1c) ia mencegah gigi saya daripada menjadi kuning\n1c) it prevents my teeth from becoming yellow": "A1c",
    "1d) ia adalah sebahagian daripada penjagaan kesihatan tubuh badan\n1d) it is part of general health care": "A1d",
    "1e) ia menjadikan gusi saya sihat\n1e) it makes my gums healthy": "A1e",
    "1f) ia membantu meningkatkan penampilan saya\n1f) it helps improve my appearance": "A1f",
    "1g) ia membuat rakan-rakan saya menyukai saya\n1g) it makes my friends to like me": "A1g",
    "2a) jika diambil terlalu kerap boleh merosakkan gigi saya\n2a) if taken too often can harm my teeth": "A2a",
    "2b) patut dielakkan\n2b) should be avoided ": "A2b",
    "2c) tidak akan membahayakan gigi saya\n2c) will not harm my teeth": "A2c",
    "3a) akan menyukarkan saya untuk bergaul dengan rakan-rakan \n3a) will make it difficult for me to mingle with my friends ": "A3a",
    "3b) membuatkan saya berasa rendah diri \n3b) makes me feel inferior": "A3b",
    "3c) membuatkan orang lain menjauhkan diri daripada saya \n3c) makes other people to avoid me": "A3c",
    "3d) boleh dicegah jika saya memberus gigi dua kali sehari\n3d) can be prevented if I brush my teeth twice a day ": "A3d",
    "4a) adalah penting bagi saya\n4a) is important to me": "A4a",
    "4b) dapat meningkatkan kesihatan gusi\n4b) can improve gum’s health": "A4b",    
    "B1. Berapa kalikah anda memberus gigi setiap hari?\nB1. How many times do you brush your teeth each day?": "B1",
    "B2. Berapa kalikah anda menggunakan ubat gigi semasa memberus gigi setiap hari?\nB2. How many times do you use toothpaste when brushing your teeth each day?": "B2",
    "B3. Adakah anda menggunakan ubat gigi berfluorida ketika memberus gigi?\nB3. Do you use fluoride toothpaste when brushing your teeth?": "B3",
    "Nyatakan jenama ubat gigi anda\nState the brand of your toothpaste": "Toothpaste Brand",
    "B4. Berapa kerap anda berkumur dengan air selepas makan setiap hari?\nB4. How often do you rinse your mouth with water after eating your food?": "B4",
    "B5. Berapa kalikah anda menggunakan flos gigi untuk membersihkan sisa makanan di celah-celah gigi setiap hari?\nB5. How many times do you use dental floss to clean the food remnants between your teeth each day?": "B5",
    "B6. Berapa kalikah anda minum minuman bergas atau minuman berkarbonat setiap hari?\nB6. How many times do you take fizzy drinks or carbonated beverages each day?": "B6",
    "B7. Adakah anda mengambil makanan / minuman manis semasa waktu makan di bawah?\nB7. Do you take sweet food and/or drinks during mealtime? [Sarapan / Breakfast ]": "B7a",
    "B7. Adakah anda mengambil makanan / minuman manis semasa waktu makan di bawah?\nB7. Do you take sweet food and/or drinks during mealtime? [Minum pagi / Morning snacks]": "B7b",
    "B7. Adakah anda mengambil makanan / minuman manis semasa waktu makan di bawah?\nB7. Do you take sweet food and/or drinks during mealtime? [Makan tengah hari / Lunch ]": "B7c",
    "B7. Adakah anda mengambil makanan / minuman manis semasa waktu makan di bawah?\nB7. Do you take sweet food and/or drinks during mealtime? [Minum petang / Tea time ]": "B7d",
    "B7. Adakah anda mengambil makanan / minuman manis semasa waktu makan di bawah?\nB7. Do you take sweet food and/or drinks during mealtime? [Makan Malam / Dinner ]": "B7e",
    "B7. Adakah anda mengambil makanan / minuman manis semasa waktu makan di bawah?\nB7. Do you take sweet food and/or drinks during mealtime? [Makan lewat malam / Supper]": "B7f",
    "B8. Adakah anda merokok\nB8. Do you smoke cigarettes?": "B8",
    "Jika jawapan anda \"Ya,\" sila nyatakan bilangan rokok yang anda hisap setiap hari\nIf your answer is \"Yes,\" please indicate the number of cigarettes you smoke per day": "B8a",
    "B9. Adakah anda menghisap vape?\nB9. Do you vape?": "B9",
    "Jika jawapan anda \"Ya,\" sila nyatakan kekerapan anda menghisap vape\nIf your answer is “Yes”, please indicate the frequency of vaping.": "B9a"
    }

df_master.rename(columns=header_mapping, inplace=True)

# Define a dictionary for keyword groups and their replacements
transformations = {
    "SMK Pengkalan": [
        "Smk pengkalan", "Smk  Pengkalan ", "SMK PENGKALAN ", "SMK PRNGKALAN",
        "SEKOLAH MENENGAH KEBANGSAAN PENGKALAN", "Sekolah kebangsaaan pengkalan"
    ],
    "SMK Kampung Pasir Puteh": [
        "SMK KG PASIR PUTEH", "SMK KAMPUNG PASIR PUTEH", "SMK KAMOUNG PASIR PUTEH",
        "SMK Kampong Pasir Puteh", "SMK KG PASIR PUTIH", "SEKOLAH MENENGAH KAMPONG PASIR PUTEH",
        "SMK Kampung Pasir Puteh ", "SMK Kg. Pasir Puteh", "Sekolah Menengah Kebangsaan Kampung Pasir Puteh",
        "Sekolah menengah kampung pasir puteh ", "SMK KG Pasir Puteh ", "SMK KAMPONG PASIR PUTEH ",
        "SMK KPP", "SMK KAMPONG PASIR PUTIH", "Sekolah Kebangsaan Kampong Pasir Puteh",
        "smk kampung pasir putih ipoh", "Sekolah kebangsaan kampung pasir puteh",
        "SMK KG PASIR PUTEH,IPOH", "SMK KAMPUNG PASIR PUTIH ", "Smkkppi",
        "SMK KAMPUNG PASIR PUTEH Ipoh", "Smk kanpong pasir putih", "SMK Kampunh Pasir Puteh",
        "SMK KAMPONG PASIR PUTEH IPOH PERAK", "SMK KG PASIR  PUTEH",
        "sekolah menengah kebangsaan kampong pasir puteh", "sekolah menegah kebangsaan kampong pasir puteh"
    ],
    "SMK Gunung Rapat": [
        "SMK Gunung Rapat", "Sekolah Menengah Kebangsaan Gunung Rapat", "SMK gunung rapat",
        "SMK GUNUNG RAPAT", "SMK Gunung Rapat ", "smk gunung rapat", "Smk Gunung Rapat",
        "Sekolah Menengah Kebangsaan Gunung Rapat ", "SMK GUNUNG RAPAT ", "SMKGUNUNG RAPAT",
        "SEKOLAH MENENGAH KEBANGSAAN GUNUNG RAPAT", "SEKOLAH MENENGAH GUNUNG RAPAT",
        "Smk gunung rapat", "SEKOLAH MENENGAH KEBANGSAAN GUNUNG RAPAT ",
        "Sekolah menengah kebangsaan gunung rapat", "SMK Gunung rapat"
    ],
    "SMK Seri Ampang": [
        "SMK SERI AMPANG", "smk seri ampang", "Smk Seri Ampang",
        "Sekolah Menengah Kebangsaan SERI AMPANG ", "Sekolah Menengah Kebangsaan Seri Ampang",
        "SMK SERI Ampang", "SMK SEI AMPANG", "SMK SERI AMPANG ", "SEKOLAH MENENGAH KEBANGSAAN SERI AMPANG",
        "SMK Seri Ampang", "smk seri amapng", "SMK SERI AMPANG,IPOH", "Smk seri ampang",
        "Sekolah Menengah Seri Ampang", "Smk seri ampang,ipoh,perak", "SMK SERI AMPANG, Ipoh ",
        "Sekolah Menengah Kebangsaan Seri Ampang ", "Smk Seri Ampang ",
        "SEKOLAH MENENGAH KEBANGSAAN SERI AMPANG ", "SMK SERI AMPANG IPOH PERAK",
        "SMK SERI AMPANG IPOH PERAK "
    ],
    "Male": ["Lelaki / Male"],
    "Female": ["Perempuan / Female", "perempuan"],
    "Amway": ["Amway"],
    "Atomy": ["atome"],
    "Colgate": ["Colgate","Colgate "," Colgate","collgate","Colgade","Golgate","Colgate…","Colget","Colgat","coldget","GOLGET","Colgate,HALAGEL ","Golged","Colcagete ","Collget ","macam ii","colgate darlie","Colgate toothpaste ","Colgate & Sensodyne","COLGALTE","COLAGATE","Colgate's","Colgate and Darlie","Colegate","Colgate Optic White","Clogate","cologate","Tak tahu"],
    "Darlie": ["darlie","Darlee","DARLI","Darlie "],
    "Dentific": ["Dentific"],
    "Fresh and White": ["FRESH AND WHITE","Fresh","Fresh white","White ","Fresh and white ","Colgate/ Fresh & white"],
    "Halagel": ["Halagel"],
    "Himalaya": ["HIMALAYA"],
    "Lion": ["Lion(Strawberry)"],
    "Lotus": ["lotus"],
    "Melaleuca": ["Melaleuca "],
    "Morning Smile": ["Mornings smile"],
    "Mukmin": ["Mukmin","Muimin","Mu'min","mui'min","MURNI "],
    "Neem": ["Neem"],
    "Oral B": ["Oral B"],
    "Pepsodent": ["Pepsodent","Pepsoden"],
    "Safi": ["Safi","Sufi"],
    "Sensodyne": ["Sensondye","Sensodyne","Sensodyn ","syndyne","sensodive"],
    "Systema": ["systema"],
    "Rural": ["SMK Pengkalan", "SMK Kampung Pasir Puteh"],
    "Urban": ["SMK Gunung Rapat", "SMK Seri Ampang"],
    "Controlled Group (CG)": ["SMK Seri Ampang", "SMK Kampung Pasir Puteh"],
    "Intervention Group (IG)": ["SMK Pengkalan", "SMK Gunung Rapat"],
    "0 stick": ["0","-",".","Colgate","no","no smoke","tak","TIADA","Tidak","nan"],
    "1 stick": ["ya","1"],
    "3 sticks": ["3 batang"],
    "Android User" : ["Ya / Yes"],
    "Non-Android User" : ["Tidak / No"],
    "Non-Flouride" : ["Atomy", "Halagel", "Melaleuca", "Mukmin"],
    "Flouride-based" : ["Amway", "Colgate", "Darlie", "Dentific", "Fresh and White", "Himalaya", "Lion", "Lotus", "Morning Smile", "Neem", "Oral B", "Pepsodent", "Safi", "Sensodyne", "Systema"],
    "Less than 2 days": ["1 kali / Once", "Apabila saya ingat / When I remember", "Beberapa kali seminggu / A few times a week"],
    "2 days and more": ["2 kali / Twice", "Lebih dari 2 kali / More than twice", "Setiap kali selepas makan / Every time after meal"],
    "No": ["Tidak / No", "Tidak tahu / Don't know"],
    "Yes": ["Ya / Yes"],
    "Yes / Sometimes": ["Kadang-kadang / Bila Perlu / Sometimes/When necessary", "Ya / Yes"],
    "Rarely / Sometimes": ["Jarang / Rarely", "Kadang-kadang/bila perlu / Sometimes / When necessary", "Tidak pernah / Never"],
    "Always": ["Selalu / Always"],
    "Using dental floss at least once": ["1 kali / Once","2 kali / Twice","Lebih dari 2 kali / More than twice"],
    "Sometimes / When in need": ["Kadang-kadang/bila perlu / Sometimes/when in need"],
    "Not using dental floss": ["Tidak menggunakan flos gigi / Not using dental floss"],
    "Up to 1 time a day": ["1 kali / Once","Kadang-kadang/Bila perlu / Sometimes/When necessary","Tidak minum minuman berkarbonat / Do not take carbonated drinks"],
    "More than 1 time a day": ["2 kali / Twice","3 kali / 3 times","4 kali / 4 times","Lebih dari 4 kali / More than 4 times"]
    }

# Generic function for transformation
def transform_column(df, column_name, transformations):
    for replacement, keywords in transformations.items():
        df[column_name] = df[column_name].apply(
            lambda x: replacement if any(keyword.lower() in str(x).lower() for keyword in keywords) else x
        )
    return df

# Apply cleasing to the DataFrame based on Dictionary values
df_master = transform_column(df_master, 'School', {
    k: v for k, v in transformations.items() if k in ["SMK Pengkalan", "SMK Kampung Pasir Puteh", "SMK Gunung Rapat", "SMK Seri Ampang"]
})
df_master = transform_column(df_master, 'Gender', {
    k: v for k, v in transformations.items() if k in ["Male", "Female"]
})
df_master = transform_column(df_master, 'Toothpaste Brand', {
    k: v for k, v in transformations.items() if k in ["Amway", "Atomy", "Colgate", "Darlie", "Dentific", "Fresh and White", "Halagel","Himalaya","Lion","Lotus","Melaleuca","Morning Smile","Mukmin","Neem", "Oral B","Pepsodent","Safi","Sensodyne","Systema"]
})
df_master = transform_column(df_master, 'Android User Check', {
    k: v for k, v in transformations.items() if k in ["Android User", "Non-Android User"]
})
df_master = transform_column(df_master, 'B1', {
    k: v for k, v in transformations.items() if k in ["Less than 2 days", "2 days and more"]
})
df_master = transform_column(df_master, 'B2', {
    k: v for k, v in transformations.items() if k in ["Less than 2 days", "2 days and more"]
})
df_master = transform_column(df_master, 'B3', {
    k: v for k, v in transformations.items() if k in ["Yes", "No"]
})
df_master = transform_column(df_master, 'B4', {
    k: v for k, v in transformations.items() if k in ["Always", "Rarely / Sometimes", "Never"]
})
df_master = transform_column(df_master, 'B5', {
    k: v for k, v in transformations.items() if k in ["Using dental floss at least once", "Sometimes / When in need", "Not using dental floss"]
})
df_master = transform_column(df_master, 'B6', {
    k: v for k, v in transformations.items() if k in ["Up to 1 time a day", "More than 1 time a day"]
})
df_master = transform_column(df_master, 'B7a', {
    k: v for k, v in transformations.items() if k in ["Yes / Sometimes", "No"]
})
df_master = transform_column(df_master, 'B7b', {
    k: v for k, v in transformations.items() if k in ["Yes / Sometimes", "No"]
})
df_master = transform_column(df_master, 'B7c', {
    k: v for k, v in transformations.items() if k in ["Yes / Sometimes", "No"]
})
df_master = transform_column(df_master, 'B7d', {
    k: v for k, v in transformations.items() if k in ["Yes / Sometimes", "No"]
})
df_master = transform_column(df_master, 'B7e', {
    k: v for k, v in transformations.items() if k in ["Yes / Sometimes", "No"]
})
df_master = transform_column(df_master, 'B7f', {
    k: v for k, v in transformations.items() if k in ["Yes / Sometimes", "No"]
})

# Define the columns you're interested in
columns = ['B7a', 'B7b', 'B7c', 'B7d', 'B7e', 'B7f']

# Count the number of "Yes / Sometimes" in each row across the specified columns
df_master['B7-Overall'] = df_master[columns].apply(lambda row: 'More than 4 times' if (row == 'Yes / Sometimes').sum() > 4 else 'Up to 4 times', axis=1)

# Duplicate columns for next transformation
df_master["Urban vs Rural"] = df_master["School"].copy()
df_master["Controlled vs Intervention"] = df_master["School"].copy()
df_master["Cigarette Count"] = df_master["B8a"].copy().astype(str)
df_master["Flouride-based"] = df_master["Toothpaste Brand"].copy()
df_master["B1 Score"] = df_master["B1"].copy()
df_master["B2 Score"] = df_master["B2"].copy()
df_master["B3 Score"] = df_master["B3"].copy()
df_master["B4 Score"] = df_master["B4"].copy()
df_master["B5 Score"] = df_master["B5"].copy()
df_master["B6 Score"] = df_master["B6"].copy()
df_master["B7a Score"] = df_master["B7a"].copy()
df_master["B7b Score"] = df_master["B7b"].copy()
df_master["B7c Score"] = df_master["B7c"].copy()
df_master["B7d Score"] = df_master["B7d"].copy()
df_master["B7e Score"] = df_master["B7e"].copy()
df_master["B7f Score"] = df_master["B7f"].copy()
df_master["B7-Overall Score"] = df_master["B7-Overall"].copy()

# Apply transformation for the copied DataFrame using the registered dictionary
df_master = transform_column(df_master, 'Urban vs Rural', {
    k: v for k, v in transformations.items() if k in ["Urban","Rural"]
})
df_master = transform_column(df_master, 'Controlled vs Intervention', {
    k: v for k, v in transformations.items() if k in ["Controlled Group (CG)","Intervention Group (IG)"]
})
df_master = transform_column(df_master, 'Cigarette Count', {
    k: v for k, v in transformations.items() if k in ["0 stick","1 stick","3 sticks"]
})
df_master = transform_column(df_master, 'Flouride-based', {
    k: v for k, v in transformations.items() if k in ["Non-Flouride", "Flouride-based"]
})

df_master['B1 Score'] = df_master['B1 Score'].apply(lambda x: 1 if x == '2 days and more' else (0 if x == 'Less than 2 days' else 9999))
df_master['B2 Score'] = df_master['B2 Score'].apply(lambda x: 1 if x == '2 days and more' else (0 if x == 'Less than 2 days' else 9999))
df_master['B3 Score'] = df_master['B3 Score'].apply(lambda x: 1 if x == 'Yes' else (0 if x == 'No' else 9999))
df_master['B4 Score'] = df_master['B4 Score'].apply(lambda x: 1 if x == 'Always' else (0 if x == 'Rarely / Sometimes' else 9999))
df_master['B5 Score'] = df_master['B5 Score'].apply(lambda x: 2 if x == 'Sometimes / When in need' else (1 if x == 'Using dental floss at least once' else (0 if x == 'Not using dental floss' else 9999)))
df_master['B6 Score'] = df_master['B6 Score'].apply(lambda x: 1 if x == 'More than 1 time a day' else (0 if x == 'Up to 1 time a day' else 9999))
df_master['B7a Score'] = df_master['B7a Score'].apply(lambda x: 1 if x == 'Yes / Sometimes' else (0 if x == 'No' else 9999))
df_master['B7b Score'] = df_master['B7b Score'].apply(lambda x: 1 if x == 'Yes / Sometimes' else (0 if x == 'No' else 9999))
df_master['B7c Score'] = df_master['B7c Score'].apply(lambda x: 1 if x == 'Yes / Sometimes' else (0 if x == 'No' else 9999))
df_master['B7d Score'] = df_master['B7d Score'].apply(lambda x: 1 if x == 'Yes / Sometimes' else (0 if x == 'No' else 9999))
df_master['B7e Score'] = df_master['B7e Score'].apply(lambda x: 1 if x == 'Yes / Sometimes' else (0 if x == 'No' else 9999))
df_master['B7f Score'] = df_master['B7f Score'].apply(lambda x: 1 if x == 'Yes / Sometimes' else (0 if x == 'No' else 9999))
df_master['B7-Overall Score'] = df_master['B7-Overall Score'].apply(lambda x: 1 if x == 'More than 4 times' else (0 if x == 'Up to 4 times' else 9999))

# Map values for K1, K2, K3 ,K4, K5, K6, K7, K8, K9, K10, K11, K12, K13, K14, K15, K16 columns
def knowledge_map_values(row):
    k1_val = 1 if row['K1'] == 'Betul / True' else 0
    k2_val = 1 if row['K2'] == 'Betul / True' else 0
    k3_val = 1 if row['K3'] == 'Betul / True' else 0  
    k4_val = 1 if row['K4'] == 'Salah / False' else 0  # For K4, "Salah / False" is mapped to 1
    k5_val = 1 if row['K5'] == 'Betul / True' else 0
    k6_val = 1 if row['K6'] == 'Betul / True' else 0
    k7_val = 1 if row['K7'] == 'Betul / True' else 0
    k8_val = 1 if row['K8'] == 'Betul / True' else 0
    k9_val = 1 if row['K9'] == 'Betul / True' else 0
    k10_val = 1 if row['K10'] == 'Betul / True' else 0
    k11_val = 1 if row['K11'] == 'Betul / True' else 0
    k12_val = 1 if row['K12'] == 'Betul / True' else 0
    k13_val = 1 if row['K13'] == 'Betul / True' else 0
    k14_val = 1 if row['K14'] == 'Betul / True' else 0
    k15_val = 1 if row['K15'] == 'Betul / True' else 0
    k16_val = 1 if row['K16'] == 'Betul / True' else 0
    return k1_val + k2_val + k3_val + k4_val + k5_val + k6_val + k7_val + k8_val + k9_val + k10_val + k11_val + k12_val + k13_val + k14_val + k15_val + k16_val
def attitude_map_values(row):
    a1a_val = 4 if row['A1a'] == 'Sangat setuju / Strongly agree' else (3 if row['A1a'] == 'Setuju / Slightly agree' else (2 if row['A1a'] == 'Tidak setuju / Slightly disagree' else 1))
    a1b_val = 4 if row['A1b'] == 'Sangat setuju / Strongly agree' else (3 if row['A1b'] == 'Setuju / Slightly agree' else (2 if row['A1b'] == 'Tidak setuju / Slightly disagree' else 1))
    a1c_val = 4 if row['A1c'] == 'Sangat setuju / Strongly agree' else (3 if row['A1c'] == 'Setuju / Slightly agree' else (2 if row['A1c'] == 'Tidak setuju / Slightly disagree' else 1))
    a1d_val = 4 if row['A1d'] == 'Sangat setuju / Strongly agree' else (3 if row['A1d'] == 'Setuju / Slightly agree' else (2 if row['A1d'] == 'Tidak setuju / Slightly disagree' else 1))
    a1e_val = 4 if row['A1e'] == 'Sangat setuju / Strongly agree' else (3 if row['A1e'] == 'Setuju / Slightly agree' else (2 if row['A1e'] == 'Tidak setuju / Slightly disagree' else 1))
    a1f_val = 4 if row['A1f'] == 'Sangat setuju / Strongly agree' else (3 if row['A1f'] == 'Setuju / Slightly agree' else (2 if row['A1f'] == 'Tidak setuju / Slightly disagree' else 1))
    a1g_val = 4 if row['A1g'] == 'Sangat setuju / Strongly agree' else (3 if row['A1g'] == 'Setuju / Slightly agree' else (2 if row['A1g'] == 'Tidak setuju / Slightly disagree' else 1))
    a2a_val = 4 if row['A2a'] == 'Sangat setuju / Strongly agree' else (3 if row['A2a'] == 'Setuju / Slightly agree' else (2 if row['A2a'] == 'Tidak setuju / Slightly disagree' else 1))
    a2b_val = 4 if row['A2b'] == 'Sangat setuju / Strongly agree' else (3 if row['A2b'] == 'Setuju / Slightly agree' else (2 if row['A2b'] == 'Tidak setuju / Slightly disagree' else 1))
    # Map A2c values with reversed scale
    a2c_val = 1 if row['A2c'] == 'Sangat setuju / Strongly agree' else (2 if row['A2c'] == 'Setuju / Slightly agree' else (3 if row['A2c'] == 'Tidak setuju / Slightly disagree' else 4))
    a3a_val = 4 if row['A3a'] == 'Sangat setuju / Strongly agree' else (3 if row['A3a'] == 'Setuju / Slightly agree' else (2 if row['A3a'] == 'Tidak setuju / Slightly disagree' else 1))
    a3b_val = 4 if row['A3b'] == 'Sangat setuju / Strongly agree' else (3 if row['A3b'] == 'Setuju / Slightly agree' else (2 if row['A3b'] == 'Tidak setuju / Slightly disagree' else 1))
    a3c_val = 4 if row['A3c'] == 'Sangat setuju / Strongly agree' else (3 if row['A3c'] == 'Setuju / Slightly agree' else (2 if row['A3c'] == 'Tidak setuju / Slightly disagree' else 1))
    a3d_val = 4 if row['A3d'] == 'Sangat setuju / Strongly agree' else (3 if row['A3d'] == 'Setuju / Slightly agree' else (2 if row['A3d'] == 'Tidak setuju / Slightly disagree' else 1))
    a4a_val = 4 if row['A4a'] == 'Sangat setuju / Strongly agree' else (3 if row['A4a'] == 'Setuju / Slightly agree' else (2 if row['A4a'] == 'Tidak setuju / Slightly disagree' else 1))
    a4b_val = 4 if row['A4b'] == 'Sangat setuju / Strongly agree' else (3 if row['A4b'] == 'Setuju / Slightly agree' else (2 if row['A4b'] == 'Tidak setuju / Slightly disagree' else 1))
    # Return the sum of A1, A2, and A3 values
    return a1a_val + a1b_val + a1c_val + a1d_val + a1e_val + a1f_val + a1g_val + a2a_val + a2b_val + a2c_val + a3a_val + a3b_val + a3c_val + a3d_val + a4a_val + a4b_val

# Apply the function to each row to create the new 'Knowledge Score' column
df_master['Knowledge Score'] = df_master.apply(knowledge_map_values, axis=1)
df_master['Attitude Score'] = df_master.apply(attitude_map_values, axis=1)

# Get current datetime in the format YYYY-MM-DD_HH-MM-SS
timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

# Construct the new filename
filename = f"output_{timestamp}.xlsx"

# Export DataFrame to Excel with the new filename
df_master.to_excel(filename, index=False)

print(f"File saved as {filename}")