import pandas as pd

def generate_base_file():
    data = {
        'Номер': ['А123АА', 'В456ВВ', 'С789СС', 'М001ММ', 'Х999ХХ'],
        'Марка': ['КАМАЗ', 'ЗИЛ', 'SCANIA', 'VOLVO', 'ГАЗЕЛЬ'],
        'Контрагент': [
            'ООО Ромашка', 
            'ОАО Строитель', 
            'ИП Иванов', 
            'ООО ТрансЛогистик',
            'АО МегаСтрой'
        ]
    }
    
    df = pd.DataFrame(data)
    
    # Сохраняем в Excel файл
    filename = 'base.xlsx'
    df.to_excel(filename, index=False)
    print(f"Файл '{filename}' успешно сгенерирован.")

if __name__ == '__main__':
    generate_base_file()
