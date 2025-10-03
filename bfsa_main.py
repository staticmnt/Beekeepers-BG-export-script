import requests
import pandas as pd
from datetime import datetime, timedelta
import time
import os

class BFSAExcelExtractor:
    def __init__(self, auth_token):
        self.session = requests.Session()
        self.base_url = "https://epord.bfsa.bg"
        self.headers = {
            'accept': 'application/json, text/plain, */*',
            'authorization': f'Bearer {auth_token}',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'referer': 'https://epord.bfsa.bg/'
        }
    
    def date_to_timestamp(self, date_str):
        """Конвертира YYYY-MM-DD в Unix timestamp"""
        dt = datetime.strptime(date_str, '%Y-%m-%d')
        return int(dt.timestamp())
    
    def timestamp_to_bulgarian_date(self, timestamp):
        """Конвертира Unix timestamp в български формат DD.MM.YYYY"""
        try:
            if isinstance(timestamp, str):
                timestamp = int(timestamp)
            return datetime.fromtimestamp(timestamp).strftime('%d.%m.%Y')
        except (ValueError, TypeError):
            return "Невалидна дата"
    
    def get_date_range_string(self, start_timestamp, end_timestamp):
        """Създава низ с диапазон от дати във формат 'DD.MM.YYYY - DD.MM.YYYY'"""
        start_date = self.timestamp_to_bulgarian_date(start_timestamp)
        end_date = self.timestamp_to_bulgarian_date(end_timestamp)
        
        if start_date != "Невалидна дата" and end_date != "Невалидна дата":
            return f"{start_date} - {end_date}"
        else:
            return ""
    
    def get_events_for_period(self, start_date, end_date):
        """Взема всички събития за даден период"""
        url = f"{self.base_url}/api/events"
        params = {
            'from': self.date_to_timestamp(start_date),
            'to': self.date_to_timestamp(end_date),
            'all': 1,
            'page': 0,
            'size': 1000
        }
        
        print(f"📅 Извличане на събития: {start_date} до {end_date}")
        
        try:
            response = self.session.get(url, params=params, headers=self.headers, timeout=60)
            
            if response.status_code == 200:
                data = response.json()
                
                if data.get('status') == 'success' and 'data' in data:
                    events_data = data['data']
                    items = events_data.get('items', [])
                    
                    print(f"✅ Намерени {len(items)} събития")
                    return items
            
            print(f"❌ Грешка от API: {response.status_code}")
            return []
            
        except Exception as e:
            print(f"❌ Грешка при заявка: {e}")
            return []
    
    def extract_data_to_dataframe(self, start_date, end_date):
        """Извлича данни и ги конвертира в DataFrame с български колони"""
        print("=" * 70)
        print("📊 ИЗВЛИЧАНЕ НА ДАННИ ЗА EXCEL")
        print("=" * 70)
        
        all_events_data = []
        
        # Обработване на по-малки партиди за да се избегне timeout
        current_start = datetime.strptime(start_date, '%Y-%m-%d')
        end_dt = datetime.strptime(end_date, '%Y-%m-%d')
        
        batch_count = 0
        
        while current_start < end_dt:
            batch_count += 1
            batch_end = current_start + timedelta(days=60)
            if batch_end > end_dt:
                batch_end = end_dt
            
            batch_start_str = current_start.strftime('%Y-%m-%d')
            batch_end_str = batch_end.strftime('%Y-%m-%d')
            
            print(f"🔄 Партида {batch_count}: {batch_start_str} до {batch_end_str}")
            
            events = self.get_events_for_period(batch_start_str, batch_end_str)
            
            for event in events:
                # Извличане на информация за датата
                start_timestamp = event.get('start_date')
                end_timestamp = event.get('end_date')
                date_range = self.get_date_range_string(start_timestamp, end_timestamp)
                
                # Извличане на основна информация
                block_name = event.get('block', {}).get('name', '')
                area = event.get('area', '')
                crop = event.get('crop', '')
                status = event.get('status', '')
                
                # Извличане на GPS координати
                centroid = event.get('block', {}).get('centroid', {})
                coordinates = centroid.get('coordinates', [])
                gps = ""
                if len(coordinates) >= 2:
                    gps = f"{coordinates[0]}, {coordinates[1]}"
                
                # Извличане на информация за продуктите
                products = event.get('products', [])
                
                if products:
                    for product in products:
                        product_name = product.get('name', '')
                        active_content = product.get('active_content', '')
                        dose = product.get('dose', '')
                        
                        event_data = {
                            'Дата': date_range,
                            'Блок': block_name,
                            'Площ': area,
                            'Култура': crop,
                            'Статус': status,
                            'Препарат': product_name,
                            'Активно вещество': active_content,
                            'Доза': dose,
                            'GPS': gps
                        }
                        all_events_data.append(event_data)
                else:
                    event_data = {
                        'Дата': date_range,
                        'Блок': block_name,
                        'Площ': area,
                        'Култура': crop,
                        'Статус': status,
                        'Препарат': '',
                        'Активно вещество': '',
                        'Доза': '',
                        'GPS': gps
                    }
                    all_events_data.append(event_data)
            
            print(f"✅ Партида {batch_count} завършена: {len(events)} събития обработени")
            
            current_start = batch_end + timedelta(days=1)
            time.sleep(1)
        
        # Създаване на DataFrame
        df = pd.DataFrame(all_events_data)
        
        print(f"\n🎯 Общо редове в Excel: {len(df)}")
        return df
    
    def save_to_excel(self, df, filename):
        """Запазва DataFrame във Excel с форматиране"""
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Събития', index=False)
                
                workbook = writer.book
                worksheet = writer.sheets['Събития']
                
                # Автоматично настройване на ширината на колоните
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                
                worksheet.freeze_panes = 'A2'
            
            print(f"📁 Excel файл запазен: {filename}")
            return True
            
        except Exception as e:
            print(f"❌ Грешка при запазване на Excel: {e}")
            return False
    
    def create_excel_report(self, start_date, end_date):
        """Основна функция за създаване на Excel отчет"""
        print("🚀 Започва извличането на данни...")
        print(f"📅 Период: {start_date} до {end_date}")
        print("⏰ Това може да отнеме няколко минути...")
        print()
        
        # Извличане на данни в DataFrame
        df = self.extract_data_to_dataframe(start_date, end_date)
        
        if df.empty:
            print("❌ Не са намерени данни за този период")
            return
        
        # Създаване на име на файл с timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        excel_filename = f"BFSA_Събития_{timestamp}.xlsx"
        
        # Запазване в Excel
        print(f"\n💾 Запазване във Excel файл...")
        success = self.save_to_excel(df, excel_filename)
        
        if success:
            print(f"\n🎉 УСПЕШНО ЗАВЪРШВАНЕ!")
            print(f"📊 Файл: {excel_filename}")
            print(f"📈 Общо редове: {len(df)}")
            print(f"📅 Период: {start_date} до {end_date}")
            
            # Показване на статистика
            print(f"\n📋 Статистика:")
            print(f"   • Уникални блокове: {df['Блок'].nunique()}")
            print(f"   • Уникални култури: {df['Култура'].nunique()}")
            print(f"   • Уникални препарати: {df['Препарат'].nunique()}")
            print(f"   • Събития с GPS: {df[df['GPS'] != ''].shape[0]}")
            
            if not df.empty:
                print(f"   • Пример за дата: {df['Дата'].iloc[0]}")
        else:
            print("❌ Неуспешно запазване на Excel файл")