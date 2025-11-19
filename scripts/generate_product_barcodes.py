"""
סקריפט חד-פעמי לייצור מק"טים לכל המוצרים הקיימים במסד הנתונים.
"""

import json
import os
from datetime import datetime


def calculate_ean13_checksum(base_12_digits: str) -> str:
    """
    Calculate EAN-13 checksum digit.
    
    Args:
        base_12_digits: String of 12 digits
        
    Returns:
        Single digit checksum (0-9)
    """
    if len(base_12_digits) != 12:
        raise ValueError("Base must be exactly 12 digits")
    
    # Sum odd positions (1st, 3rd, 5th, etc. - index 0, 2, 4...)
    odd_sum = sum(int(base_12_digits[i]) for i in range(0, 12, 2))
    
    # Sum even positions (2nd, 4th, 6th, etc. - index 1, 3, 5...) and multiply by 3
    even_sum = sum(int(base_12_digits[i]) for i in range(1, 12, 2)) * 3
    
    # Total sum
    total = odd_sum + even_sum
    
    # Checksum is (10 - (total mod 10)) mod 10
    checksum = (10 - (total % 10)) % 10
    
    return str(checksum)


def generate_next_barcode(current_barcode: str) -> str:
    """
    Generate the next barcode in sequence with proper EAN-13 checksum.
    
    Args:
        current_barcode: Current 13-digit EAN-13 barcode
        
    Returns:
        Next 13-digit EAN-13 barcode
    """
    if len(current_barcode) != 13:
        raise ValueError("Barcode must be exactly 13 digits")
    
    # Remove checksum digit (last digit)
    base_12 = current_barcode[:12]
    
    # Increment the base
    base_number = int(base_12) + 1
    
    # Pad back to 12 digits
    new_base_12 = str(base_number).zfill(12)
    
    # Calculate new checksum
    checksum = calculate_ean13_checksum(new_base_12)
    
    # Return full 13-digit barcode
    return new_base_12 + checksum


def main():
    """Main function to generate barcodes for all existing products."""
    print("=" * 60)
    print("ייצור מק\"טים למוצרים קיימים")
    print("=" * 60)
    print()
    
    # Load products catalog
    products_file = 'products_catalog.json'
    if not os.path.exists(products_file):
        print(f"❌ שגיאה: קובץ {products_file} לא נמצא")
        return
    
    with open(products_file, 'r', encoding='utf-8') as f:
        products = json.load(f)
    
    print(f"📦 נטענו {len(products)} מוצרים")
    
    # Load last barcode
    barcodes_file = 'barcodes_data.json'
    if os.path.exists(barcodes_file):
        with open(barcodes_file, 'r', encoding='utf-8') as f:
            barcodes_data = json.load(f)
        last_barcode = barcodes_data.get('last_barcode', '7297555019592')
    else:
        last_barcode = '7297555019592'
    
    print(f"🔢 הברקוד האחרון: {last_barcode}")
    print()
    
    # Count products without barcode
    products_without_barcode = [p for p in products if not p.get('barcode')]
    print(f"📊 מוצרים ללא מק\"ט: {len(products_without_barcode)}")
    print(f"📊 מוצרים עם מק\"ט: {len(products) - len(products_without_barcode)}")
    print()
    
    if not products_without_barcode:
        print("✅ כל המוצרים כבר יש להם מק\"ט!")
        return
    
    # Ask for confirmation
    response = input(f"❓ האם לייצר {len(products_without_barcode)} מק\"טים חדשים? (y/n): ")
    if response.lower() not in ['y', 'yes', 'כן']:
        print("❌ בוטל על ידי המשתמש")
        return
    
    print()
    print("⏳ מייצר מק\"טים...")
    print()
    
    # Generate barcodes for products without one
    current_barcode = last_barcode
    generated_count = 0
    
    for product in products:
        if not product.get('barcode'):
            # Generate next barcode
            current_barcode = generate_next_barcode(current_barcode)
            product['barcode'] = current_barcode
            generated_count += 1
            
            # Print progress every 10 products
            if generated_count % 10 == 0:
                print(f"  ✓ נוצרו {generated_count} מק\"טים...")
    
    print(f"  ✓ נוצרו {generated_count} מק\"טים!")
    print()
    
    # Save updated products catalog
    print("💾 שומר קטלוג מוצרים...")
    with open(products_file, 'w', encoding='utf-8') as f:
        json.dump(products, f, indent=2, ensure_ascii=False)
    print("  ✓ הקטלוג נשמר בהצלחה")
    
    # Update last barcode
    print("💾 מעדכן ברקוד אחרון...")
    barcodes_data = {
        'last_barcode': current_barcode,
        'last_updated': datetime.now().isoformat()
    }
    with open(barcodes_file, 'w', encoding='utf-8') as f:
        json.dump(barcodes_data, f, indent=2, ensure_ascii=False)
    print(f"  ✓ הברקוד האחרון עודכן ל: {current_barcode}")
    
    print()
    print("=" * 60)
    print("✅ הסקריפט הסתיים בהצלחה!")
    print(f"📊 סה\"כ נוצרו {generated_count} מק\"טים")
    print(f"🔢 הברקוד האחרון במערכת: {current_barcode}")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print()
        print("=" * 60)
        print(f"❌ שגיאה: {str(e)}")
        print("=" * 60)
        import traceback
        traceback.print_exc()

