import sys
import subprocess
import os

def install_packages():
    """Инсталира необходимите пакети, ако не са инсталирани"""
    packages = ['pandas', 'openpyxl', 'requests']
    
    for package in packages:
        try:
            __import__(package)
            print(f"✅ {package} е вече инсталиран")
        except ImportError:
            print(f"📦 Инсталиране на {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def main():
    print("=" * 70)
    print("BFSA Excel Екстрактор - Настройване на средата...")
    print("=" * 70)
    
    # Инсталиране на необходимите пакети
    install_packages()
    
    print("\n" + "=" * 70)
    print("Стартиране на BFSA Excel Екстрактор...")
    print("=" * 70)
    
    # Импортиране и стартиране на основния скрипт
    from bfsa_main import BFSAExcelExtractor
    
    # Вашият токен за удостоверяване
    print("\n🔐 Моля, въведете вашия токен за удостоверяване:")
    print("1. Отидете на https://epord.bfsa.bg/")
    print("2. Влезте в акаунта си")
    print("3. Натиснете F12 за да отворите Developer Tools")
    print("4. Отидете на раздела Network")
    print("5. Използвайте филтъра за дати на сайта")
    print("6. Намерете някой API заявка и копирайте 'Authorization' хедъра")
    print("7. Поставете вашия токен тук (започва с 'eyJ...'):")
    
    auth_token = input("Токен: ").strip()
    
    if not auth_token:
        print("❌ Не е предоставен токен. Излизане.")
        return
    
    extractor = BFSAExcelExtractor(auth_token)
    
    # Въвеждане на период от потребителя
    print("\n📅 Моля, въведете период за извличане на данни:")
    print("   Формат: YYYY-MM-DD (например: 2024-01-01)")
    
    start_date = input("Начална дата (YYYY-MM-DD): ").strip()
    end_date = input("Крайна дата (YYYY-MM-DD): ").strip()
    
    # Проверка на датите
    try:
        from datetime import datetime
        datetime.strptime(start_date, '%Y-%m-%d')
        datetime.strptime(end_date, '%Y-%m-%d')
    except ValueError:
        print("❌ Невалиден формат на датата! Използвайте YYYY-MM-DD")
        return
    
    if start_date >= end_date:
        print("❌ Началната дата трябва да е преди крайната дата!")
        return
    
    print(f"\n📅 Период: {start_date} до {end_date}")
    print("⏰ Това може да отнеме няколко минути...")
    print()
    
    input("Натиснете Enter за да започнете извличането...")
    
    try:
        extractor.create_excel_report(start_date, end_date)
    except Exception as e:
        print(f"\n💥 Грешка: {e}")
        import traceback
        traceback.print_exc()
    
    print("\nНатиснете Enter за изход...")
    input()

if __name__ == "__main__":
    main()