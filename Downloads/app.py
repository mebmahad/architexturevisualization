from flask import Flask, request, render_template
from openpyxl import Workbook

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        date = request.form['date']
        morning_tasks = request.form['morning']
        afternoon_tasks = request.form['afternoon']
        evening_tasks = request.form['evening']

        # Write data to Excel file
        write_to_excel(date, morning_tasks, afternoon_tasks, evening_tasks)

        return 'Schedule saved successfully!'
    return render_template('schedule_form.html')

def write_to_excel(date, morning_tasks, afternoon_tasks, evening_tasks):
    # Create or load workbook
    wb = Workbook()
    ws = wb.active

    # Set headers if new workbook
    if not ws['A1'].value:
        ws.append(['Date', 'Morning Tasks', 'Afternoon Tasks', 'Evening Tasks'])

    # Append data
    ws.append([date, morning_tasks, afternoon_tasks, evening_tasks])

    # Save workbook
    wb.save('daily_schedule.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
