#!/usr/bin/env python3
"""
.XLS dosyasındaki görselleri bulmak için gelişmiş analiz
"""

import struct
import sys

def find_image_signatures(filepath):
    """Dosyada görsel signature'larını arar"""
    
    # Bilinen görsel formatları
    signatures = {
        b'\xFF\xD8\xFF': 'JPEG',
        b'\x89PNG': 'PNG',
        b'GIF87a': 'GIF87',
        b'GIF89a': 'GIF89',
        b'BM': 'BMP',
        b'\xD7\xCD\xC6\x9A': 'WMF (Windows Metafile)',
        b'\x01\x00\x00\x00': 'EMF (Enhanced Metafile) veya diğer',
        b'\x00\x00\x01\x00': 'ICO',
        b'RIFF': 'RIFF (WebP veya WAV)',
    }
    
    with open(filepath, 'rb') as f:
        data = f.read()
    
    print(f"📂 Dosya: {filepath}")
    print(f"📏 Boyut: {len(data):,} byte\n")
    
    # BIFF kayıtlarını parse et
    print("🔍 BIFF Kayıtları Analizi:")
    print("=" * 70)
    
    pos = 0
    record_count = 0
    imdata_records = []
    
    while pos + 4 <= len(data):
        rec_type = struct.unpack('<H', data[pos:pos+2])[0]
        rec_len = struct.unpack('<H', data[pos+2:pos+4])[0]
        
        if pos + 4 + rec_len > len(data):
            break
        
        record_count += 1
        
        # IMDATA ve ilgili kayıtlar
        if rec_type in [0x7F, 0xE0, 0xE9, 0xEB]:
            rec_data = data[pos+4:pos+4+rec_len]
            hex_preview = ' '.join(f'{b:02X}' for b in rec_data[:16])
            
            print(f"Record #{record_count}:")
            print(f"  Type: 0x{rec_type:04X}")
            print(f"  Length: {rec_len:,} byte")
            print(f"  Hex: {hex_preview}...")
            
            imdata_records.append({
                'pos': pos,
                'type': rec_type,
                'len': rec_len,
                'data': rec_data
            })
            
            print()
        
        pos += 4 + rec_len
    
    print(f"\n📊 Toplam Kayıt: {record_count:,}")
    print(f"📊 IMDATA Kayıt: {len(imdata_records)}\n")
    
    # Tüm dosyada signature ara
    print("🔍 Görsel Signature Taraması:")
    print("=" * 70)
    
    found_images = []
    
    for sig, name in signatures.items():
        offset = 0
        count = 0
        while True:
            pos = data.find(sig, offset)
            if pos == -1:
                break
            
            # Önceki 100 byte'ı kontrol et (BIFF header'ı bulmak için)
            context_start = max(0, pos - 100)
            context = data[context_start:pos]
            
            # BIFF kaydı içinde mi?
            in_record = False
            for rec in imdata_records:
                if rec['pos'] <= pos < rec['pos'] + 4 + rec['len']:
                    in_record = True
                    break
            
            print(f"✓ {name} bulundu: offset={pos:,} (0x{pos:X})")
            if in_record:
                print(f"  → BIFF kaydı içinde")
            else:
                print(f"  → Serbest alanda")
            
            count += 1
            found_images.append((name, pos))
            offset = pos + len(sig)
        
        if count > 0:
            print(f"  Toplam: {count} adet\n")
    
    print("\n" + "=" * 70)
    print(f"📊 Toplam Görsel Signature: {len(found_images)}")
    
    if len(found_images) == 0:
        print("\n⚠️  HİÇ GÖRSEL SIGNATURE BULUNAMADI!")
        print("    Görseller muhtemelen:")
        print("    • OLE compound stream içinde gömülü")
        print("    • Farklı bir encoding ile saklanmış")
        print("    • Harici dosya olarak linked")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Kullanım: python find_images_in_xls.py <dosya.xls>")
        sys.exit(1)
    
    find_image_signatures(sys.argv[1])
