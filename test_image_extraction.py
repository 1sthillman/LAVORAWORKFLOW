#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LAVORA EXCEL GÖRSEL ÇIKARMA TESTİ
Excel dosyasından görselleri çıkar ve analiz et
"""

import sys
import zipfile
import io
from pathlib import Path
import xml.etree.ElementTree as ET

def extract_images_from_xlsx(xlsx_file):
    """Excel dosyasından tüm görselleri çıkar"""
    print("="*100)
    print("🖼️  EXCEL GÖRSEL ÇIKARMA TESTİ")
    print("="*100)
    print(f"📁 Dosya: {xlsx_file}\n")
    
    results = {
        'sheets': {},
        'total_images': 0,
        'image_files': []
    }
    
    try:
        with zipfile.ZipFile(xlsx_file, 'r') as zip_file:
            # 1. Workbook ilişkilerini oku
            print("📋 1. Workbook yapısı analiz ediliyor...")
            
            # workbook.xml'i oku
            try:
                workbook_xml = zip_file.read('xl/workbook.xml').decode('utf-8')
                root = ET.fromstring(workbook_xml)
                
                # Namespace
                ns = {
                    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                }
                
                # Sayfa isimlerini al
                sheets = {}
                for sheet in root.findall('.//main:sheet', ns):
                    name = sheet.get('name')
                    sheet_id = sheet.get('sheetId')
                    rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    sheets[rid] = name
                    print(f"  ✓ Sayfa bulundu: {name} (ID: {sheet_id}, rId: {rid})")
                
                print(f"\n  Toplam {len(sheets)} sayfa bulundu\n")
                
            except Exception as e:
                print(f"  ⚠️ Workbook okunamadı: {e}")
                sheets = {}
            
            # 2. Görselleri bul
            print("🔍 2. Görseller aranıyor...")
            
            # xl/media/ klasöründe görselleri ara
            media_files = [f for f in zip_file.namelist() if f.startswith('xl/media/')]
            
            if media_files:
                print(f"  ✓ xl/media/ klasöründe {len(media_files)} dosya bulundu:")
                for media_file in media_files:
                    file_info = zip_file.getinfo(media_file)
                    size_kb = file_info.file_size / 1024
                    print(f"    - {media_file} ({size_kb:.1f} KB)")
                    results['image_files'].append(media_file)
                    results['total_images'] += 1
            else:
                print("  ⚠️ xl/media/ klasöründe görsel bulunamadı")
            
            # 3. Drawing ilişkilerini analiz et
            print("\n📐 3. Drawing (çizim) dosyaları analiz ediliyor...")
            
            # xl/drawings/ klasöründeki tüm dosyaları bul
            drawing_files = [f for f in zip_file.namelist() if f.startswith('xl/drawings/') and f.endswith('.xml')]
            
            if drawing_files:
                print(f"  ✓ {len(drawing_files)} drawing dosyası bulundu:")
                
                for drawing_file in drawing_files:
                    print(f"\n  📄 {drawing_file}:")
                    
                    try:
                        drawing_xml = zip_file.read(drawing_file).decode('utf-8')
                        drawing_root = ET.fromstring(drawing_xml)
                        
                        # xdr namespace
                        xdr_ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
                        
                        # Tüm resimleri bul
                        images = drawing_root.findall('.//xdr:pic', xdr_ns)
                        
                        if images:
                            print(f"    ✓ {len(images)} resim tanımı bulundu")
                            
                            # Her resim için detay
                            for idx, pic in enumerate(images, 1):
                                # Pozisyon bilgisi
                                anchor = pic.getparent()
                                if anchor is not None:
                                    from_elem = anchor.find('.//xdr:from', xdr_ns)
                                    to_elem = anchor.find('.//xdr:to', xdr_ns)
                                    
                                    if from_elem is not None:
                                        row_from = from_elem.find('.//xdr:row', xdr_ns)
                                        col_from = from_elem.find('.//xdr:col', xdr_ns)
                                        if row_from is not None and col_from is not None:
                                            print(f"      [{idx}] Pozisyon: Satır {row_from.text}, Sütun {col_from.text}")
                                    
                                    if to_elem is not None:
                                        row_to = to_elem.find('.//xdr:row', xdr_ns)
                                        col_to = to_elem.find('.//xdr:col', xdr_ns)
                                        if row_to is not None and col_to is not None:
                                            print(f"           → Satır {row_to.text}, Sütun {col_to.text}")
                        else:
                            print("    ℹ️ Bu drawing'de resim yok")
                    
                    except Exception as e:
                        print(f"    ⚠️ Drawing okunamadı: {e}")
            else:
                print("  ⚠️ Hiç drawing dosyası bulunamadı")
            
            # 4. Sayfaların drawing ilişkilerini kontrol et
            print("\n🔗 4. Sayfa-Drawing ilişkileri kontrol ediliyor...")
            
            # xl/worksheets/_rels/ klasöründeki .rels dosyalarını oku
            worksheet_rels = [f for f in zip_file.namelist() 
                            if f.startswith('xl/worksheets/_rels/') and f.endswith('.rels')]
            
            if worksheet_rels:
                for rels_file in worksheet_rels:
                    sheet_file = rels_file.replace('xl/worksheets/_rels/', '').replace('.rels', '')
                    print(f"\n  📄 {sheet_file}:")
                    
                    try:
                        rels_xml = zip_file.read(rels_file).decode('utf-8')
                        rels_root = ET.fromstring(rels_xml)
                        
                        # Relationship namespace
                        rels_ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                        
                        # Drawing ilişkilerini bul
                        relationships = rels_root.findall('.//rel:Relationship', rels_ns)
                        
                        for rel in relationships:
                            rel_type = rel.get('Type')
                            target = rel.get('Target')
                            
                            if 'drawing' in rel_type.lower():
                                print(f"    ✓ Drawing ilişkisi bulundu: {target}")
                    
                    except Exception as e:
                        print(f"    ⚠️ İlişki dosyası okunamadı: {e}")
            else:
                print("  ⚠️ Hiç sayfa ilişkisi bulunamadı")
            
    except zipfile.BadZipFile:
        print("❌ HATA: Bu geçerli bir .xlsx dosyası değil!")
        return None
    except Exception as e:
        print(f"❌ HATA: {e}")
        import traceback
        traceback.print_exc()
        return None
    
    # Özet
    print("\n" + "="*100)
    print("📊 ÖZET")
    print("="*100)
    print(f"✓ Toplam görsel sayısı: {results['total_images']}")
    print(f"✓ Görsel dosyaları: {len(results['image_files'])}")
    
    if results['total_images'] > 0:
        print("\n📁 Bulunan görseller:")
        for img in results['image_files']:
            print(f"  - {img}")
    else:
        print("\n⚠️ HİÇBİR GÖRSEL BULUNAMADI!")
        print("\nOlası sebepler:")
        print("  1. Excel dosyasında gerçekten görsel yok")
        print("  2. Görseller farklı bir formatta (embedded objects)")
        print("  3. Eski .xls formatı (ZIP değil)")
    
    print("\n" + "="*100)
    print("✅ TEST TAMAMLANDI")
    print("="*100)
    
    return results

if __name__ == "__main__":
    excel_file = Path("01-CUBE-BK VOGER 13+3 ODA.xls")
    
    if not excel_file.exists():
        print(f"❌ Dosya bulunamadı: {excel_file}")
        sys.exit(1)
    
    # Dosya uzantısını kontrol et
    if excel_file.suffix.lower() == '.xls':
        print("\n⚠️ UYARI: Bu eski .xls formatında!")
        print("Eski Excel formatı (BIFF) ZIP tabanlı değildir.")
        print(".xlsx formatına dönüştürün veya özel bir kütüphane kullanın.\n")
        
        try:
            # xlrd ile eski format okuma (görselsiz)
            import xlrd
            wb = xlrd.open_workbook(excel_file)
            print(f"✓ xlrd ile açıldı: {wb.nsheets} sayfa")
            print("⚠️ Ancak xlrd görselleri okuyamaz!\n")
        except Exception as e:
            print(f"❌ xlrd hatası: {e}\n")
        
        sys.exit(0)
    
    results = extract_images_from_xlsx(excel_file)
