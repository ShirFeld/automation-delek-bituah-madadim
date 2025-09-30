#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'BituahRechev'))
from insurance_scraper import InsuranceScraper

def test_special_only():
    """בדיקה של רכב מיוחד בלבד"""
    scraper = InsuranceScraper()
    try:
        print("🚗 מתחיל בדיקת רכב מיוחד בלבד...")
        
        # שליפת נתוני רכב מיוחד בלבד
        special_data = scraper.scrape_special_vehicle_data()
        
        if special_data:
            print(f"\n✅ הצלחה! נתוני רכב מיוחד:")
            for key, value in special_data.items():
                print(f"  {key}: {value} ₪")
        else:
            print("❌ נכשל בשליפת נתוני רכב מיוחד")
            
    except Exception as e:
        print(f"❌ שגיאה בבדיקת רכב מיוחד: {str(e)}")
        
    finally:
        scraper.close_driver()

if __name__ == "__main__":
    test_special_only()

