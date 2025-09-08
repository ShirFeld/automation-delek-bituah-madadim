#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from insurance_scraper import InsuranceScraper

def test_commercial_only():
    """בדיקה של רכב מסחרי בלבד"""
    scraper = InsuranceScraper()
    try:
        print("🚛 מתחיל בדיקת רכב מסחרי בלבד...")
        
        # שליפת נתוני רכב מסחרי בלבד
        commercial_data = scraper.scrape_commercial_car_data()
        
        if commercial_data:
            print(f"\n✅ הצלחה! נתוני רכב מסחרי:")
            for age_group, weights in commercial_data.items():
                print(f"  {age_group}:")
                for weight, price in weights.items():
                    print(f"    {weight}: {price} ₪")
        else:
            print("❌ נכשל בשליפת נתוני רכב מסחרי")
            
    except Exception as e:
        print(f"❌ שגיאה: {str(e)}")
    finally:
        scraper.close_driver()

if __name__ == "__main__":
    test_commercial_only()
