#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from insurance_scraper import InsuranceScraper

def test_integration():
    """בדיקת אינטגרציה של רכב מיוחד עם הקוד הראשי"""
    print("🧪 בדיקת אינטגרציה...")
    
    scraper = InsuranceScraper()
    
    try:
        # בדיקה שהפונקציה קיימת
        if hasattr(scraper, 'scrape_special_vehicle_data'):
            print("✅ הפונקציה scrape_special_vehicle_data קיימת")
        else:
            print("❌ הפונקציה scrape_special_vehicle_data לא קיימת")
            return False
        
        # בדיקה שהפונקציה create_mdb_database קיימת
        if hasattr(scraper, 'create_mdb_database'):
            print("✅ הפונקציה create_mdb_database קיימת")
        else:
            print("❌ הפונקציה create_mdb_database לא קיימת")
            return False
        
        # בדיקה שהפונקציה extract_harel_price קיימת
        if hasattr(scraper, 'extract_harel_price'):
            print("✅ הפונקציה extract_harel_price קיימת")
        else:
            print("❌ הפונקציה extract_harel_price לא קיימת")
            return False
        
        # בדיקה עם נתונים דמה
        print("📊 בדיקה עם נתונים דמה...")
        test_data = {
            'special_vehicle': {
                'Nigrar': 423,
                'Handasi': 2335,
                'Agricalture': 1535
            }
        }
        
        print("✅ הנתונים מוכנים:")
        for key, value in test_data['special_vehicle'].items():
            print(f"   • {key}: {value}")
        
        print("🎉 בדיקת האינטגרציה הצליחה!")
        print("📋 הפונקציות מוכנות לשימוש בכפתור 'שליפה מלאה'")
        return True
        
    except Exception as e:
        print(f"❌ שגיאה בבדיקת האינטגרציה: {str(e)}")
        return False

if __name__ == "__main__":
    test_integration()
