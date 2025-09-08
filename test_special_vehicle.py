#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os

# הוספת הנתיב לתיקיית BituahRechev
sys.path.append(os.path.join(os.path.dirname(__file__), 'BituahRechev'))

from insurance_scraper import InsuranceScraper

def test_special_vehicle():
    """בדיקה של רכב מיוחד בלבד"""
    print("🚗 מתחיל בדיקת רכב מיוחד בלבד!")
    print("="*60)
    
    scraper = InsuranceScraper()
    
    try:
        if scraper.setup_driver(visible=True):
            print("✅ דפדפן מוכן")
            
            # שליפת נתונים לרכב מיוחד
            print("\n🚗 מתחיל רכב מיוחד...")
            special_results = scraper.scrape_special_vehicle_data()
            
            if special_results:
                print(f"\n🎉 תוצאות רכב מיוחד:")
                for category, price in special_results.items():
                    if price:
                        print(f"✅ {category}: {price:,.0f} ₪")
                    else:
                        print(f"❌ {category}: לא נמצא מחיר")
            else:
                print("❌ לא התקבלו תוצאות לרכב מיוחד")
                
        else:
            print("❌ שגיאה בדפדפן")
            
    except Exception as e:
        print(f"❌ שגיאה: {str(e)}")
        
    finally:
        scraper.cleanup()
        print("\n🏁 בדיקה הושלמה")

if __name__ == "__main__":
    test_special_vehicle()



