from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import json
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

class DataHandler:
    def __init__(self):
        self.data_file = "data.json"
        
    def update_dispatch_data(self, location, quantity, day_of_month):
        """تحديث بيانات التوزيع"""
        try:
            print(f"📍 Updating dispatch: '{location}', Qty: {quantity}, Day: {day_of_month}")
            
            # حفظ البيانات في JSON
            data = self._load_data()
            
            if 'dispatch' not in data:
                data['dispatch'] = {}
            
            data['dispatch'][f"{location}_{day_of_month}"] = {
                'quantity': float(quantity),
                'date': datetime.now().isoformat(),
                'location': location,
                'day_of_month': day_of_month
            }
            
            self._save_data(data)
            print("✅ Dispatch data saved successfully!")
            return True
            
        except Exception as e:
            print(f"❌ Error updating dispatch data: {e}")
            return False
    
    def update_water_data(self, ship_number, meter1_final, meter2_final, meter1_previous, date):
        """تحديث بيانات المياه"""
        try:
            print(f"🚢 Updating water data - Ship: {ship_number}")
            
            data = self._load_data()
            
            if 'water' not in data:
                data['water'] = {}
            
            data['water'][f"ship_{ship_number}_{date}"] = {
                'ship_number': ship_number,
                'meter1_final': float(meter1_final),
                'meter2_final': float(meter2_final),
                'meter1_previous': float(meter1_previous),
                'date': date,
                'volume': float(meter1_final) - float(meter1_previous)
            }
            
            self._save_data(data)
            print("✅ Water data saved successfully!")
            return True
            
        except Exception as e:
            print(f"❌ Error updating water data: {e}")
            return False
    
    def _load_data(self):
        """تحميل البيانات من ملف JSON"""
        try:
            with open(self.data_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}
    
    def _save_data(self, data):
        """حفظ البيانات لملف JSON"""
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

# إنشاء كائن DataHandler
data_handler = DataHandler()

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/save_data', methods=['POST'])
def save_data():
    try:
        data = request.json
        print(f"📨 Received data: {data}")
        
        if data['type'] == 'dispatch':
            success = data_handler.update_dispatch_data(
                data['location'],
                data['quantity'],
                data['dayOfMonth']
            )
        elif data['type'] == 'meter':
            success = data_handler.update_water_data(
                data['shipNumber'],
                data['meter1Final'],
                data['meter2Final'],
                data['meter1Previous'],
                data['date']
            )
        else:
            success = False
            
        return jsonify({'success': success, 'message': 'تم الحفظ بنجاح' if success else 'فشل في الحفظ'})
        
    except Exception as e:
        print(f"🔥 Error in save_data: {e}")
        return jsonify({'success': False, 'message': f'خطأ: {str(e)}'})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    print(f"🚀 Starting Water Dispatch Application on port {port}...")
    app.run(host='0.0.0.0', port=port, debug=False)
