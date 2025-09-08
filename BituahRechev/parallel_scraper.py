#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from insurance_scraper import InsuranceScraper
import queue

class ParallelInsuranceScraper:
    def __init__(self, max_browsers=4):
        self.max_browsers = max_browsers
        self.results = {}
        self.scrapers = []
        self.lock = threading.Lock()
        
    def create_scraper_instance(self):
        """יוצר instance חדש של scraper"""
        scraper = InsuranceScraper()
        if scraper.setup_driver(visible=False):
            return scraper
        return None
    
    def scrape_single_scenario(self, scenario_data):
        """שליפת תרחיש יחיד"""
        scenario_id = scenario_data['id']
        car_type = scenario_data['car_type']
        age = scenario_data['age']
        engine_volume = scenario_data['engine_volume']
        license_years = scenario_data['license_years']
        
        try:
            # יצירת scraper חדש לכל תרחיש
            scraper = self.create_scraper_instance()
            if not scraper:
                return scenario_id, None, "Failed to create scraper"
            
            print(f"🔄 מתחיל תרחיש {scenario_id}: {car_type}, גיל {age}, נפח {engine_volume}")
            
            # שליפת המחיר
            price = scraper.scrape_single_combination(car_type, age, engine_volume, license_years)
            
            # ניקוי
            scraper.cleanup()
            
            if price:
                print(f"✅ תרחיש {scenario_id}: {price} ₪")
                return scenario_id, price, "Success"
            else:
                print(f"❌ תרחיש {scenario_id}: נכשל")
                return scenario_id, None, "No price found"
                
        except Exception as e:
            print(f"❌ שגיאה בתרחיש {scenario_id}: {str(e)}")
            return scenario_id, None, str(e)
    
    def prepare_all_scenarios(self):
        """הכנת כל התרחישים לעיבוד מקבילי"""
        scenarios = []
        
        # תרחישי רכב פרטי (24 תרחישים)
        private_scenarios = [
            # גיל 17-20
            {"id": "P1", "car_type": "private", "age": 19, "engine_volume": 900, "license_years": 2},
            {"id": "P2", "car_type": "private", "age": 19, "engine_volume": 1200, "license_years": 2},
            {"id": "P3", "car_type": "private", "age": 19, "engine_volume": 1800, "license_years": 2},
            {"id": "P4", "car_type": "private", "age": 19, "engine_volume": 2200, "license_years": 2},
            
            # גיל 21-23
            {"id": "P5", "car_type": "private", "age": 22, "engine_volume": 900, "license_years": 4},
            {"id": "P6", "car_type": "private", "age": 22, "engine_volume": 1200, "license_years": 4},
            {"id": "P7", "car_type": "private", "age": 22, "engine_volume": 1800, "license_years": 4},
            {"id": "P8", "car_type": "private", "age": 22, "engine_volume": 2200, "license_years": 4},
            
            # גיל 24-29
            {"id": "P9", "car_type": "private", "age": 25, "engine_volume": 900, "license_years": 7},
            {"id": "P10", "car_type": "private", "age": 25, "engine_volume": 1200, "license_years": 7},
            {"id": "P11", "car_type": "private", "age": 25, "engine_volume": 1800, "license_years": 7},
            {"id": "P12", "car_type": "private", "age": 25, "engine_volume": 2200, "license_years": 7},
            
            # גיל 30-39
            {"id": "P13", "car_type": "private", "age": 31, "engine_volume": 900, "license_years": 13},
            {"id": "P14", "car_type": "private", "age": 31, "engine_volume": 1200, "license_years": 13},
            {"id": "P15", "car_type": "private", "age": 31, "engine_volume": 1800, "license_years": 13},
            {"id": "P16", "car_type": "private", "age": 31, "engine_volume": 2200, "license_years": 13},
            
            # גיל 40-49
            {"id": "P17", "car_type": "private", "age": 41, "engine_volume": 900, "license_years": 23},
            {"id": "P18", "car_type": "private", "age": 41, "engine_volume": 1200, "license_years": 23},
            {"id": "P19", "car_type": "private", "age": 41, "engine_volume": 1800, "license_years": 23},
            {"id": "P20", "car_type": "private", "age": 41, "engine_volume": 2200, "license_years": 23},
            
            # גיל 50+
            {"id": "P21", "car_type": "private", "age": 51, "engine_volume": 900, "license_years": 33},
            {"id": "P22", "car_type": "private", "age": 51, "engine_volume": 1200, "license_years": 33},
            {"id": "P23", "car_type": "private", "age": 51, "engine_volume": 1800, "license_years": 33},
            {"id": "P24", "car_type": "private", "age": 51, "engine_volume": 2200, "license_years": 33},
        ]
        
        # תרחישי רכב מסחרי (10 תרחישים)
        commercial_scenarios = [
            # גיל 17-20
            {"id": "C1", "car_type": "commercial", "age": 19, "engine_volume": 4, "license_years": 1},
            {"id": "C2", "car_type": "commercial", "age": 19, "engine_volume": 4.5, "license_years": 1},
            
            # גיל 21-23
            {"id": "C3", "car_type": "commercial", "age": 23, "engine_volume": 4, "license_years": 5},
            {"id": "C4", "car_type": "commercial", "age": 23, "engine_volume": 4.5, "license_years": 5},
            
            # גיל 24-39
            {"id": "C5", "car_type": "commercial", "age": 25, "engine_volume": 4, "license_years": 7},
            {"id": "C6", "car_type": "commercial", "age": 25, "engine_volume": 4.5, "license_years": 7},
            
            # גיל 40-49
            {"id": "C7", "car_type": "commercial", "age": 43, "engine_volume": 4, "license_years": 17},
            {"id": "C8", "car_type": "commercial", "age": 43, "engine_volume": 4.5, "license_years": 17},
            
            # גיל 50+
            {"id": "C9", "car_type": "commercial", "age": 52, "engine_volume": 4, "license_years": 26},
            {"id": "C10", "car_type": "commercial", "age": 52, "engine_volume": 4.5, "license_years": 26},
        ]
        
        scenarios = private_scenarios + commercial_scenarios
        return scenarios
    
    def run_parallel_scraping(self, progress_callback=None):
        """הפעלת שליפה מקבילית"""
        scenarios = self.prepare_all_scenarios()
        total_scenarios = len(scenarios)
        completed = 0
        
        print(f"🚀 מתחיל שליפה מקבילית עם {self.max_browsers} דפדפנים")
        print(f"🎯 סך הכל: {total_scenarios} תרחישים")
        
        start_time = time.time()
        
        # עיבוד מקבילי
        with ThreadPoolExecutor(max_workers=self.max_browsers) as executor:
            # שליחת כל התרחישים לעיבוד
            future_to_scenario = {
                executor.submit(self.scrape_single_scenario, scenario): scenario 
                for scenario in scenarios
            }
            
            # איסוף התוצאות
            for future in as_completed(future_to_scenario):
                scenario = future_to_scenario[future]
                try:
                    scenario_id, price, status = future.result()
                    
                    with self.lock:
                        self.results[scenario_id] = {
                            'price': price,
                            'status': status,
                            'scenario': scenario
                        }
                        completed += 1
                    
                    if progress_callback:
                        progress_callback(completed, total_scenarios, scenario_id, price, status)
                    
                    print(f"📊 התקדמות: {completed}/{total_scenarios} ({(completed/total_scenarios)*100:.1f}%)")
                    
                except Exception as e:
                    print(f"❌ שגיאה בעיבוד תרחיש {scenario['id']}: {str(e)}")
        
        end_time = time.time()
        total_time = end_time - start_time
        
        # סיכום תוצאות
        successful = len([r for r in self.results.values() if r['price'] is not None])
        
        print(f"\n🏆 סיכום שליפה מקבילית:")
        print(f"✅ הצליח: {successful}/{total_scenarios} תרחישים")
        print(f"⏱️ זמן כולל: {total_time:.1f} שניות")
        print(f"⚡ ממוצע לתרחיש: {total_time/total_scenarios:.1f} שניות")
        
        return self.results
    
    def organize_results_for_tables(self):
        """ארגון התוצאות לטבלאות"""
        organized_data = {
            'private_car': {},
            'commercial_car': {}
        }
        
        # מיפוי מזהי תרחישים לטבלאות
        scenario_mapping = {
            # רכב פרטי
            'P1': ('private_car', '17-20', 'עד 1050'),
            'P2': ('private_car', '17-20', 'מ-1051 עד 1550'),
            'P3': ('private_car', '17-20', 'מ-1551 עד 2050'),
            'P4': ('private_car', '17-20', 'מ-2051 ומעלה'),
            
            'P5': ('private_car', '21-23', 'עד 1050'),
            'P6': ('private_car', '21-23', 'מ-1051 עד 1550'),
            'P7': ('private_car', '21-23', 'מ-1551 עד 2050'),
            'P8': ('private_car', '21-23', 'מ-2051 ומעלה'),
            
            'P9': ('private_car', '24-29', 'עד 1050'),
            'P10': ('private_car', '24-29', 'מ-1051 עד 1550'),
            'P11': ('private_car', '24-29', 'מ-1551 עד 2050'),
            'P12': ('private_car', '24-29', 'מ-2051 ומעלה'),
            
            'P13': ('private_car', '30-39', 'עד 1050'),
            'P14': ('private_car', '30-39', 'מ-1051 עד 1550'),
            'P15': ('private_car', '30-39', 'מ-1551 עד 2050'),
            'P16': ('private_car', '30-39', 'מ-2051 ומעלה'),
            
            'P17': ('private_car', '40-49', 'עד 1050'),
            'P18': ('private_car', '40-49', 'מ-1051 עד 1550'),
            'P19': ('private_car', '40-49', 'מ-1551 עד 2050'),
            'P20': ('private_car', '40-49', 'מ-2051 ומעלה'),
            
            'P21': ('private_car', '50- ומעלה', 'עד 1050'),
            'P22': ('private_car', '50- ומעלה', 'מ-1051 עד 1550'),
            'P23': ('private_car', '50- ומעלה', 'מ-1551 עד 2050'),
            'P24': ('private_car', '50- ומעלה', 'מ-2051 ומעלה'),
            
            # רכב מסחרי
            'C1': ('commercial_car', '17-20', 'עד 4000 (כולל)'),
            'C2': ('commercial_car', '17-20', 'מעל 4000'),
            'C3': ('commercial_car', '21-23', 'עד 4000 (כולל)'),
            'C4': ('commercial_car', '21-23', 'מעל 4000'),
            'C5': ('commercial_car', '24-39', 'עד 4000 (כולל)'),
            'C6': ('commercial_car', '24-39', 'מעל 4000'),
            'C7': ('commercial_car', '40-49', 'עד 4000 (כולל)'),
            'C8': ('commercial_car', '40-49', 'מעל 4000'),
            'C9': ('commercial_car', '50- ומעלה', 'עד 4000 (כולל)'),
            'C10': ('commercial_car', '50- ומעלה', 'מעל 4000'),
        }
        
        # ארגון הנתונים
        for scenario_id, result in self.results.items():
            if scenario_id in scenario_mapping and result['price']:
                car_type, age_group, engine_group = scenario_mapping[scenario_id]
                
                if age_group not in organized_data[car_type]:
                    organized_data[car_type][age_group] = {}
                
                organized_data[car_type][age_group][engine_group] = result['price']
        
        return organized_data
    
    def cleanup_all(self):
        """ניקוי כל ה-scrapers"""
        for scraper in self.scrapers:
            try:
                scraper.cleanup()
            except:
                pass

def main():
    """פונקציה לבדיקה"""
    parallel_scraper = ParallelInsuranceScraper(max_browsers=4)
    
    def progress_update(completed, total, scenario_id, price, status):
        print(f"📈 {scenario_id}: {price if price else 'נכשל'} | {completed}/{total}")
    
    try:
        results = parallel_scraper.run_parallel_scraping(progress_callback=progress_update)
        organized_data = parallel_scraper.organize_results_for_tables()
        
        # יצירת טבלאות
        scraper = InsuranceScraper()
        scraper.insurance_data = organized_data
        image_path = scraper.save_tables_as_image()
        
        print(f"\n📊 טבלאות נשמרו ב: {image_path}")
        
    finally:
        parallel_scraper.cleanup_all()

if __name__ == "__main__":
    main()

