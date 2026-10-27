#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import xlrd

wb = xlrd.open_workbook("01-CUBE-BK VOGER 13+3 ODA.xls")
sheet = wb.sheet_by_name("YATAKLAR")

print("="*100)
print("YATAKLAR SAYFASI - TÜM VERİLER")
print("="*100)

for row_idx in range(sheet.nrows):
    row = [sheet.cell_value(row_idx, col) for col in range(sheet.ncols)]
    # Boş satırları gösterme
    if any(str(c).strip() for c in row):
        print(f"[{row_idx:2}] {' | '.join(str(c)[:30] for c in row)}")

print("\n" + "="*100)
print("ÇİFT KİŞİLİK COMFORT YATAK ve BAZA ARANIYOR")
print("="*100)

for row_idx in range(sheet.nrows):
    row = [str(sheet.cell_value(row_idx, col)) for col in range(sheet.ncols)]
    row_text = ' '.join(row).upper()
    
    if 'ÇIFT' in row_text and 'YATAK' in row_text:
        print(f"\n[{row_idx}] BULUNDU!")
        for col_idx, cell in enumerate(row):
            print(f"  Sütun {col_idx}: {cell}")
    
    if 'BAZA' in row_text:
        print(f"\n[{row_idx}] BAZA BULUNDU!")
        for col_idx, cell in enumerate(row):
            print(f"  Sütun {col_idx}: {cell}")
