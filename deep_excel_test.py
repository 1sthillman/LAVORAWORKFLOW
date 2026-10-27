#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LAVORA EXCEL DEEP ANALYSIS
Binlerce farklı Excel formatını test etmek için süper akıllı analiz
"""

import sys
import re
from pathlib import Path

try:
    import xlrd
    print("✓ xlrd modülü yüklü")
except ImportError:
    print("❌ xlrd modülü yüklenmemiş. Yükleniyor...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlrd"])
    import xlrd

def parse_number(value):
    """Türkçe ve İngilizce sayı formatlarını parse et"""
    if isinstance(value, (int, float)):
        return float(value)
    
    s = str(value).strip()
    if not s:
        return 0.0
    
    # Virgül ve nokta analizi
    has_comma = ',' in s
    has_dot = '.' in s
    
    if has_comma and has_dot:
        # Her ikisi de var: hangisi son?
        last_comma = s.rfind(',')
        last_dot = s.rfind('.')
        if last_comma > last_dot:
            # Türkçe format: 1.234.567,89
            s = s.replace('.', '').replace(',', '.')
        else:
            # İngilizce format: 1,234,567.89
            s = s.replace(',', '')
    elif has_comma:
        # Sadece virgül var
        # 1.234,56 veya 1234,56 -> Türkçe
        # 1,234 veya 12,345 -> belirsiz
        if '.' in s:
            # 1.234,56 kesin Türkçe
            s = s.replace('.', '').replace(',', '.')
        else:
            # Virgülden sonra kaç hane?
            parts = s.split(',')
            if len(parts) == 2 and len(parts[1]) <= 2:
                # 1234,56 -> Türkçe ondalık
                s = s.replace(',', '.')
            else:
                # 1,234 -> İngilizce binlik ayracı
                s = s.replace(',', '')
    elif has_dot:
        # Sadece nokta var
        dots = s.count('.')
        if dots > 1:
            # 1.234.567 -> Türkçe binlik
            s = s.replace('.', '')
        else:
            # Tek nokta: 1234.56 ya da 1.234?
            parts = s.split('.')
            if len(parts) == 2 and len(parts[1]) == 3 and not any(c.isdigit() for c in s[s.index('.')+4:]):
                # 1.234 gibi -> muhtemelen binlik
                s = s.replace('.', '')
            # Yoksa ondalık olarak bırak
    
    # Sadece rakamları, - ve . bırak
    s = re.sub(r'[^\d.\-]', '', s)
    
    try:
        return float(s)
    except ValueError:
        return 0.0

def normalize_text(text):
    """Türkçe karakterleri normalize et"""
    text = str(text).upper()
    replacements = {
        'İ': 'I', 'Ş': 'S', 'Ğ': 'G', 'Ü': 'U', 
        'Ö': 'O', 'Ç': 'C', 'ı': 'i', 'ş': 's',
        'ğ': 'g', 'ü': 'u', 'ö': 'o', 'ç': 'c'
    }
    for tr_char, en_char in replacements.items():
        text = text.replace(tr_char, en_char)
    return text

def detect_column_types(headers):
    """Sütun tiplerini otomatik tespit et"""
    headers_norm = [normalize_text(h) for h in headers]
    
    result = {
        'name': -1,
        'quantity': -1,
        'unit_price': -1,
        'total': -1,
        'description': -1,
        'material': -1,
        'dimension': -1
    }
    
    patterns = {
        'quantity': r'ADET|MIKTAR|QTY|QUANTITY|METRAJ',
        'unit_price': r'BIRIM\s*FIYAT|FIYAT|PRICE|UCRET',
        'total': r'TOPLAM|TUTAR|TOTAL|BEDEL',
        'name': r'URUN|KALEM|MALZEME|TANIM|ACIKLAMA|NAME|ITEM',
        'description': r'ACIKLAMA|DETAY|DESCRIPTION|OZELLIK',
        'material': r'MALZEME|KAPLAMA|MATERIAL|FINISH',
        'dimension': r'OLCU|EBAT|BOYUT|DIMENSION|SIZE'
    }
    
    for i, header in enumerate(headers_norm):
        for col_type, pattern in patterns.items():
            if re.search(pattern, header):
                if result[col_type] == -1:  # İlk eşleşmeyi al
                    result[col_type] = i
                    break
    
    return result

def analyze_sheet(sheet, sheet_name):
    """Tek bir Excel sayfasını detaylı analiz et"""
    print(f"\n{'='*100}")
    print(f"📄 SAYFA: {sheet_name}")
    print(f"{'='*100}")
    
    # Özet mi detay mı?
    is_summary = bool(re.search(r'ICMAL|OZET|SUMMARY', normalize_text(sheet_name)))
    print(f"📋 Tip: {'ÖZET SAYFASI 📊' if is_summary else 'DETAY SAYFASI 📝'}")
    
    # Başlık satırını bul
    header_row_idx = -1
    headers = []
    
    for row_idx in range(min(10, sheet.nrows)):
        row = [str(sheet.cell_value(row_idx, col)).strip() for col in range(sheet.ncols)]
        # Başlık satırı: birden fazla dolu hücre var ve içinde "ADET", "FIYAT" gibi kelimeler var
        non_empty = [c for c in row if c]
        if len(non_empty) >= 2:
            row_text = ' '.join(row).upper()
            if any(kw in row_text for kw in ['ADET', 'FIYAT', 'TOPLAM', 'URUN', 'KALEM']):
                header_row_idx = row_idx
                headers = row
                break
    
    if header_row_idx == -1:
        print("⚠️ Başlık satırı bulunamadı!")
        return
    
    print(f"✓ Başlık satırı: Satır {header_row_idx}")
    print(f"📑 Başlıklar: {headers}")
    
    # Sütun tiplerini tespit et
    cols = detect_column_types(headers)
    print(f"\n🔍 Tespit Edilen Sütunlar:")
    for col_type, col_idx in cols.items():
        if col_idx >= 0:
            print(f"  ✓ {col_type:15} -> Sütun {col_idx:2} ({headers[col_idx]})")
        else:
            print(f"  ✗ {col_type:15} -> BULUNAMADI")
    
    # Veri satırlarını analiz et
    print(f"\n📊 Veri Satırları Analizi:")
    print("-" * 100)
    
    total_sum = 0.0
    row_count = 0
    issues = []
    
    for row_idx in range(header_row_idx + 1, sheet.nrows):
        row = [sheet.cell_value(row_idx, col) for col in range(sheet.ncols)]
        
        # Boş satırları atla
        if not any(str(c).strip() for c in row):
            continue
        
        # Toplam satırını atla
        first_cell = str(row[0]).strip().upper()
        if any(kw in first_cell for kw in ['TOPLAM', 'TOTAL', 'GENEL', 'IMZA', 'NOT']):
            continue
        
        name = str(row[cols['name']] if cols['name'] >= 0 else row[0]).strip()
        if not name or len(name) < 2:
            continue
        
        qty_val = row[cols['quantity']] if cols['quantity'] >= 0 else ''
        pr_val = row[cols['unit_price']] if cols['unit_price'] >= 0 else ''
        tot_val = row[cols['total']] if cols['total'] >= 0 else ''
        
        qty = parse_number(qty_val)
        unit_price = parse_number(pr_val)
        total = parse_number(tot_val)
        
        row_count += 1
        print(f"\n[{row_idx}] {name}")
        print(f"      Adet     : {qty_val:>15} → {qty:>12.2f}")
        print(f"      Birim Fyt: {pr_val:>15} → {unit_price:>12.2f} TL")
        print(f"      Toplam   : {tot_val:>15} → {total:>12.2f} TL")
        
        # Çarpım kontrolü
        if qty > 0 and unit_price > 0 and total > 0:
            expected = qty * unit_price
            diff = abs(expected - total)
            
            if diff > 1:
                msg = f"⚠️ UYARI: {qty} × {unit_price:.2f} = {expected:.2f} ama dosyada {total:.2f} (fark: {diff:.2f})"
                print(f"      {msg}")
                issues.append((row_idx, name, msg))
            else:
                print(f"      ✓ Doğru: {qty} × {unit_price:.2f} = {total:.2f}")
        
        if total > 0:
            total_sum += total
    
    print(f"\n{'='*100}")
    print(f"📈 ÖZET:")
    print(f"  • Toplam satır sayısı: {row_count}")
    print(f"  • SAYFA TOPLAMI: {total_sum:,.2f} TL")
    print(f"  • Tespit edilen sorun: {len(issues)}")
    
    if issues:
        print(f"\n⚠️ SORUNLU SATIRLAR:")
        for row_idx, name, msg in issues:
            print(f"  [{row_idx}] {name}: {msg}")
    else:
        print(f"\n✅ TÜM SATIRLAR DOĞRU!")
    
    return {
        'name': sheet_name,
        'is_summary': is_summary,
        'total': total_sum,
        'row_count': row_count,
        'issues': issues,
        'columns': cols
    }

def analyze_multiplier_logic(workbook):
    """Çarpan mantığını test et"""
    print(f"\n\n{'#'*100}")
    print(f"🧮 ÇARPAN MANTIĞI ANALİZİ")
    print(f"{'#'*100}")
    
    # Özet sayfasını bul
    summary_sheet = None
    summary_name = None
    
    for sheet_name in workbook.sheet_names():
        if re.search(r'ICMAL|OZET|SUMMARY', normalize_text(sheet_name)):
            summary_sheet = workbook.sheet_by_name(sheet_name)
            summary_name = sheet_name
            break
    
    if not summary_sheet:
        print("⚠️ Özet sayfası bulunamadı. Çarpan mantığı test edilemiyor.")
        return
    
    print(f"✓ Özet sayfası bulundu: {summary_name}")
    
    # Her detay sayfası için özet sayfasıyla karşılaştır
    for sheet_name in workbook.sheet_names():
        if sheet_name == summary_name:
            continue
        
        sheet = workbook.sheet_by_name(sheet_name)
        print(f"\n{'─'*80}")
        print(f"🔗 {sheet_name} ↔️ {summary_name}")
        
        # Detay sayfası toplamını hesapla
        detail_total = calculate_sheet_total(sheet)
        print(f"  Detay toplamı: {detail_total:,.2f} TL")
        
        # Özet sayfasında bu detay sayfasını bul
        # 1. Yöntem: İsim benzerliği
        # 2. Yöntem: Toplam tutarı karşılaştır
        
        # Basitleştirilmiş isimler
        detail_name_simple = re.sub(r'[^A-Z0-9]', '', normalize_text(sheet_name))
        
        found_match = False
        for row_idx in range(summary_sheet.nrows):
            row = [str(summary_sheet.cell_value(row_idx, col)).strip() for col in range(summary_sheet.ncols)]
            row_text = ' '.join(row)
            
            # Satırda bu sayfanın adı geçiyor mu?
            if sheet_name.lower() in row_text.lower():
                print(f"  ✓ İsim eşleşmesi bulundu: Satır {row_idx}")
                print(f"    İçerik: {row_text[:100]}")
                found_match = True
                
                # Bu satırdaki adet ve fiyatı kontrol et
                for col_idx, cell in enumerate(row):
                    val = parse_number(cell)
                    if val > 0:
                        print(f"      Sütun {col_idx}: {cell} → {val}")
                break
        
        if not found_match:
            print(f"  ⚠️ Özet sayfasında '{sheet_name}' bulunamadı!")
            
            # Toplam tutara göre ara
            print(f"  🔍 Toplam tutara göre aranıyor ({detail_total:.2f} TL)...")
            for row_idx in range(summary_sheet.nrows):
                for col_idx in range(summary_sheet.ncols):
                    val = parse_number(summary_sheet.cell_value(row_idx, col_idx))
                    if abs(val - detail_total) < 1:
                        row = [str(summary_sheet.cell_value(row_idx, col)).strip() for col in range(summary_sheet.ncols)]
                        print(f"  💡 Olası eşleşme: Satır {row_idx}, Sütun {col_idx}")
                        print(f"     İçerik: {' | '.join(row)}")
                        break

def calculate_sheet_total(sheet):
    """Bir sayfanın toplamını hesapla"""
    # Başlık satırını bul
    header_row = -1
    for row_idx in range(min(10, sheet.nrows)):
        row = [str(sheet.cell_value(row_idx, col)).strip() for col in range(sheet.ncols)]
        row_text = ' '.join(row).upper()
        if any(kw in row_text for kw in ['TOPLAM', 'TOTAL', 'TUTAR']):
            header_row = row_idx
            break
    
    if header_row == -1:
        return 0.0
    
    # Toplam sütununu bul
    headers = [str(sheet.cell_value(header_row, col)).strip() for col in range(sheet.ncols)]
    total_col = -1
    for i, h in enumerate(headers):
        if re.search(r'TOPLAM|TUTAR|TOTAL', normalize_text(h)):
            total_col = i
            break
    
    if total_col == -1:
        return 0.0
    
    # Toplamı hesapla
    total = 0.0
    for row_idx in range(header_row + 1, sheet.nrows):
        first_cell = str(sheet.cell_value(row_idx, 0)).strip().upper()
        if any(kw in first_cell for kw in ['TOPLAM', 'TOTAL', 'GENEL', 'IMZA']):
            continue
        val = parse_number(sheet.cell_value(row_idx, total_col))
        if val > 0:
            total += val
    
    return total

def main():
    excel_file = Path("01-CUBE-BK VOGER 13+3 ODA.xls")
    
    if not excel_file.exists():
        print(f"❌ Dosya bulunamadı: {excel_file}")
        return
    
    print("="*100)
    print("🚀 LAVORA EXCEL DEEP ANALYSIS")
    print("="*100)
    print(f"📁 Dosya: {excel_file}")
    
    # Excel dosyasını aç
    try:
        workbook = xlrd.open_workbook(excel_file)
    except Exception as e:
        print(f"❌ Dosya açılamadı: {e}")
        return
    
    print(f"✓ Dosya başarıyla açıldı")
    print(f"📊 Toplam sayfa sayısı: {workbook.nsheets}")
    print(f"📑 Sayfa isimleri: {', '.join(workbook.sheet_names())}")
    
    # Her sayfayı analiz et
    results = []
    for sheet_name in workbook.sheet_names():
        sheet = workbook.sheet_by_name(sheet_name)
        result = analyze_sheet(sheet, sheet_name)
        if result:
            results.append(result)
    
    # Çarpan mantığını test et
    analyze_multiplier_logic(workbook)
    
    # Genel özet
    print(f"\n\n{'='*100}")
    print("📊 GENEL ÖZET")
    print(f"{'='*100}")
    
    total_all = sum(r['total'] for r in results)
    total_rows = sum(r['row_count'] for r in results)
    total_issues = sum(len(r['issues']) for r in results)
    
    print(f"✓ Analiz edilen sayfa sayısı: {len(results)}")
    print(f"✓ Toplam satır sayısı: {total_rows}")
    print(f"✓ GENEL TOPLAM: {total_all:,.2f} TL")
    print(f"✓ Tespit edilen sorun sayısı: {total_issues}")
    
    if total_issues == 0:
        print(f"\n✅ TÜM SAYFALAR DOĞRU! HİÇBİR SORUN YOK!")
    else:
        print(f"\n⚠️ {total_issues} ADET SORUN TESPİT EDİLDİ!")
    
    print(f"\n{'='*100}")
    print("✅ ANALİZ TAMAMLANDI")
    print(f"{'='*100}")

if __name__ == "__main__":
    main()
