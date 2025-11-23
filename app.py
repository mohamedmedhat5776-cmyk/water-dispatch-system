from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from openpyxl import load_workbook
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

class ExcelHandler:
    def __init__(self):
        self.local_file = "Dispatch order.xlsx"
        
    def update_dispatch_data(self, location, quantity, day_of_month):
        """تحديث بيانات التوزيع في Excel"""
        try:
            # فتح ملف Excel الموجود
            wb = load_workbook(self.local_file)
            sheet = wb[" Daily Dispatch"]
            
            print(f"📍 Updating Excel: '{location}', Qty: {quantity}, Day: {day_of_month}")
            
            # البحث عن الموقع في العمود B (العمود 2)
            row_num = None
            for row in range(4, 80):  # من الصف 4 إلى 79
                cell_value = sheet.cell(row=row, column=2).value
                if cell_value and str(cell_value).strip() == location:
                    row_num = row
                    break
            
            if row_num:
                # تحديد العمود بناءً على اليوم (G=7 هو اليوم 1)
                column_num = 6 + int(day_of_month)
                
                # تحديث الخلية
                sheet.cell(row=row_num, column=column_num).value = float(quantity)
                
                # حفظ الملف
                wb.save(self.local_file)
                print("✅ Excel file updated successfully!")
                return True
            else:
                print(f"❌ Location '{location}' not found in Excel")
                return False
                
        except Exception as e:
            print(f"❌ Error updating Excel: {e}")
            return False
    
    def update_water_data(self, ship_number, meter1_final, meter2_final, meter1_previous, date):
        """تحديث بيانات المياه في Excel"""
        try:
            # فتح ملف Excel
            wb = load_workbook(self.local_file)
            sheet = wb["Water Quantity"]
            
            print(f"🚢 Updating water data - Ship: {ship_number}")
            
            # تحديث بيانات السفينة (الصفوف 7-10)
            row_num = 6 + int(ship_number)  # السفينة 1 -> الصف 7، السفينة 2 -> الصف 8، إلخ
            
            # العمود E (5) - القراءة النهائية
            sheet.cell(row=row_num, column=5).value = float(meter1_final)
            
            # العمود D (4) - القراءة السابقة
            sheet.cell(row=row_num, column=4).value = float(meter1_previous)
            
            # حساب الحجم تلقائياً - العمود F (6)
            volume = float(meter1_final) - float(meter1_previous)
            sheet.cell(row=row_num, column=6).value = volume
            
            # حفظ الملف
            wb.save(self.local_file)
            print("✅ Water data updated in Excel!")
            return True
            
        except Exception as e:
            print(f"❌ Error updating water data: {e}")
            return False

# إنشاء كائن ExcelHandler
excel_handler = ExcelHandler()

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/save_data', methods=['POST'])
def save_data():
    try:
        data = request.json
        print(f"📨 Received data: {data}")
        
        if data['type'] == 'dispatch':
            success = excel_handler.update_dispatch_data(
                data['location'],
                data['quantity'],
                data['dayOfMonth']
            )
        elif data['type'] == 'meter':
            success = excel_handler.update_water_data(
                data['shipNumber'],
                data['meter1Final'],
                data['meter2Final'],
                data['meter1Previous'],
                data['date']
            )
        else:
            success = False
            
        return jsonify({'success': success, 'message': 'تم الحفظ في Excel بنجاح' if success else 'فشل في الحفظ'})
        
    except Exception as e:
        print(f"🔥 Error in save_data: {e}")
        return jsonify({'success': False, 'message': f'خطأ: {str(e)}'})

@app.route('/view_data')
def view_data():
    """عرض رسالة تأكيد أن التطبيق شغال"""
    return jsonify({
        'status': 'running',
        'message': 'التطبيق شغال بنجاح!',
        'excel_file': 'Dispatch order.xlsx'
    })

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    print(f"🚀 Starting Water Dispatch Application on port {port}...")
    print(f"📊 Using Excel file: Dispatch order.xlsx")
    app.run(host='0.0.0.0', port=port, debug=False)
