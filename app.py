from flask import Flask, render_template, request, send_from_directory, redirect, url_for
import os
import pandas as pd
import re
from decimal import Decimal, getcontext

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['EXCEL_FILE'] = os.path.join(app.config['UPLOAD_FOLDER'], 'customer_loans.xlsx')

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Set decimal precision
getcontext().prec = 10

def load_data():
    if os.path.exists(app.config['EXCEL_FILE']):
        return pd.read_excel(app.config['EXCEL_FILE']).to_dict('records')
    return []

def save_data(data):
    df = pd.DataFrame(data)
    df.to_excel(app.config['EXCEL_FILE'], index=False)

def format_reference_number(ref_num):
    """Format reference number according to business rules"""
    cleaned = re.sub(r'\s+', '', ref_num)
    if ' ' in ref_num:
        return '   '.join(cleaned)
    return cleaned

def format_name(name):
    """Format name according to business rules"""
    name = name.upper()
    parts = re.split(r'\.|\s+', name)
    parts = [p for p in parts if p]
    if len(parts) >= 3:
        return f"{parts[0]}.{parts[1]}  {parts[2]}  {'  '.join(parts[3:])}"
    return '  '.join(parts)

def format_city_state(city_state):
    """Format city and state according to business rules"""
    city_state = city_state.upper()
    if 'DC' in city_state:
        return 'NA'
    if ',' in city_state:
        city, state = city_state.split(',')
        return f"{city.strip()} , {state.strip()}"
    return city_state

def words_to_number(words):
    """Convert written numbers to numeric values"""
    word_to_num = {
        'zero': 0, 'one': 1, 'two': 2, 'three': 3, 'four': 4,
        'five': 5, 'six': 6, 'seven': 7, 'eight': 8, 'nine': 9,
        'ten': 10, 'eleven': 11, 'twelve': 12, 'thirteen': 13, 'fourteen': 14,
        'fifteen': 15, 'sixteen': 16, 'seventeen': 17, 'eighteen': 18, 'nineteen': 19,
        'twenty': 20, 'thirty': 30, 'forty': 40, 'fifty': 50,
        'sixty': 60, 'seventy': 70, 'eighty': 80, 'ninety': 90,
        'hundred': 100, 'thousand': 1000, 'million': 1000000,
        'billion': 1000000000, 'trillion': 1000000000000
    }
    
    words = words.lower().replace('-', ' ').replace(' and ', ' ').split()
    number = 0
    current = 0
    
    for word in words:
        if word in word_to_num:
            num = word_to_num[word]
            if num == 100:
                current *= num
            elif num >= 1000:
                current *= num
                number += current
                current = 0
            else:
                current += num
        elif word == 'cents':
            break
    
    number += current
    return number

def format_currency(value):
    """Format currency with commas and spaces as per business rules"""
    if pd.isna(value) or value == 0:
        return "NA"
    
    value = Decimal(str(value))
    int_part = int(value)
    dec_part = value - int_part
    
    str_value = f"{int_part:,}"
    parts = str_value.split(',')
    
    formatted_parts = []
    for i, part in enumerate(parts):
        if len(parts) - i == 4:
            formatted_parts.append(f"  {part}  ,")
        elif len(parts) - i == 3:
            formatted_parts.append(f"  {part}  ,")
        elif len(parts) - i == 2:
            formatted_parts.append(f"  {part}  ,")
        else:
            formatted_parts.append(part)
    
    formatted_int = ''.join(formatted_parts).strip(',')
    formatted_dec = f"{dec_part:.2f}"[1:] if dec_part else ".00"
    
    return f"$  {formatted_int}{formatted_dec}"

def calculate_loan_details(data):
    """Perform all loan calculations based on input data"""
    try:
        purchase_value_words = data['Purchase Value (in words)']
        purchase_value = Decimal(words_to_number(purchase_value_words))
        
        purchase_reduction = Decimal(data['Purchase Value Reduction (%)']) / 100
        down_payment = Decimal(data['Down Payment (%)']) / 100
        loan_period = Decimal(data['Loan Period (Years)'])
        annual_interest = Decimal(data['Annual Interest Rate (%)']) / 100
        monthly_principal_reduction = Decimal(data['Monthly Principal Reduction (%)']) / 100
        total_interest_reduction = Decimal(data['Total Interest Reduction (%)']) / 100
        
        reduced_purchase_value = purchase_value * purchase_reduction
        reduced_purchase_value = reduced_purchase_value.quantize(Decimal('0.01'))
        purchase_value_to_enter = purchase_value - reduced_purchase_value
        
        downpayment_value = purchase_value_to_enter * down_payment
        downpayment_value = downpayment_value.quantize(Decimal('0.01'))
        loan_amount = purchase_value_to_enter - downpayment_value
        
        annual_principal = loan_amount / loan_period
        annual_principal = annual_principal.quantize(Decimal('0.01'))
        monthly_principal = annual_principal / 12
        monthly_principal = monthly_principal.quantize(Decimal('0.01'))
        principal_to_enter = monthly_principal * monthly_principal_reduction
        principal_to_enter = principal_to_enter.quantize(Decimal('0.01'))
        
        interest_per_annum = loan_amount * annual_interest
        interest_per_annum = interest_per_annum.quantize(Decimal('0.01'))
        total_interest = interest_per_annum * loan_period
        total_interest = total_interest.quantize(Decimal('0.01'))
        total_interest_to_enter = total_interest * total_interest_reduction
        total_interest_to_enter = total_interest_to_enter.quantize(Decimal('0.01'))
        
        loan_percent = (1 - down_payment) * 100
        if loan_percent <= 84.99:
            insurance_rate = Decimal('0.0032')
        elif loan_percent <= 85:
            insurance_rate = Decimal('0.0021')
        elif loan_percent <= 90:
            insurance_rate = Decimal('0.0041')
        elif loan_percent <= 95:
            insurance_rate = Decimal('0.0067')
        else:
            insurance_rate = Decimal('0.0085')
        
        property_insurance_annum = loan_amount * insurance_rate
        property_insurance_annum = property_insurance_annum.quantize(Decimal('0.01'))
        property_insurance_month = property_insurance_annum / 12
        property_insurance_month = property_insurance_month.quantize(Decimal('0.01'))
        
        if 80.01 <= loan_percent <= 85 and loan_period <= 20:
            pmi_rate = Decimal('0.0019')
        elif 80.01 <= loan_percent <= 85 and loan_period > 20:
            pmi_rate = Decimal('0.0032')
        elif 85.01 <= loan_percent <= 90 and loan_period <= 20:
            pmi_rate = Decimal('0.0023')
        elif 85.01 <= loan_percent <= 90 and loan_period > 20:
            pmi_rate = Decimal('0.0052')
        elif 90.01 <= loan_percent <= 95 and loan_period <= 20:
            pmi_rate = Decimal('0.0026')
        elif 90.01 <= loan_percent <= 95 and loan_period > 20:
            pmi_rate = Decimal('0.0078')
        else:
            pmi_rate = None
        
        if pmi_rate is not None:
            pmi_annum = loan_amount * pmi_rate
            pmi_annum = pmi_annum.quantize(Decimal('0.01'))
        else:
            pmi_annum = "NA"
        
        formatted_results = {
            'Customer Reference Number': format_reference_number(data['Customer Reference Number']),
            'Customer Name': format_name(data['Customer Name']),
            'City, State': format_city_state(data['City, State']),
            'Purchase Value and Down Payment': f"{format_currency(purchase_value_to_enter)} AND {int(down_payment*100)}%",
            'Loan Period and Annual Interest': f"{int(loan_period)} YEARS AND {annual_interest*100:.2f}%",
            'Guarantor Name': format_name(data['Guarantor Name']),
            'Guarantor Reference Number': format_reference_number(data['Guarantor Reference Number']),
            'Loan amount and principal': f"{format_currency(loan_amount)} AND {format_currency(principal_to_enter)}",
            'Total Interest for Loan Period and Property tax for Loan Period': f"{format_currency(total_interest_to_enter)} AND NA",
            'Property Insurance per month and PMI per annum': f"{format_currency(property_insurance_month)} AND {pmi_annum if isinstance(pmi_annum, str) else format_currency(pmi_annum)}",
            # Store original values for editing
            'Purchase Value (in words)': data['Purchase Value (in words)'],
            'Purchase Value Reduction (%)': data['Purchase Value Reduction (%)'],
            'Down Payment (%)': data['Down Payment (%)'],
            'Loan Period (Years)': data['Loan Period (Years)'],
            'Annual Interest Rate (%)': data['Annual Interest Rate (%)'],
            'Monthly Principal Reduction (%)': data['Monthly Principal Reduction (%)'],
            'Total Interest Reduction (%)': data['Total Interest Reduction (%)'],
        }
        
        return formatted_results
    
    except Exception as e:
        print(f"Error in calculations: {str(e)}")
        return None

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['GET', 'POST'])
def submit():
    if request.method == 'POST':
        try:
            raw_data = {
                'Customer Reference Number': request.form['customerRef'],
                'Customer Name': request.form['customerName'],
                'City, State': request.form['cityState'],
                'Purchase Value (in words)': request.form['purchaseValue'],
                'Purchase Value Reduction (%)': request.form['purchaseReduction'],
                'Down Payment (%)': request.form['downPayment'],
                'Loan Period (Years)': request.form['loanPeriod'],
                'Annual Interest Rate (%)': request.form['annualInterest'],
                'Monthly Principal Reduction (%)': request.form['monthlyPrincipalReduction'],
                'Total Interest Reduction (%)': request.form['totalInterestReduction'],
                'Guarantor Name': request.form['guarantorName'],
                'Guarantor Reference Number': request.form['guarantorRef'],
            }
            
            formatted_data = calculate_loan_details(raw_data)
            
            if not formatted_data:
                raise ValueError("Failed to process loan data")
            
            entries = load_data()
            entries.append(formatted_data)
            save_data(entries)
            
            return render_template('submit.html', message="Data submitted successfully!", entries=entries)
        
        except Exception as e:
            return render_template('submit.html', message=f"Error processing data: {str(e)}", entries=load_data())
    
    return render_template('submit.html', entries=load_data())

@app.route('/edit/<int:index>', methods=['GET'])
def edit(index):
    entries = load_data()
    if 0 <= index < len(entries):
        return render_template('edit.html', entry=entries[index], index=index)
    return redirect(url_for('submit'))

@app.route('/update/<int:index>', methods=['POST'])
def update(index):
    try:
        raw_data = {
            'Customer Reference Number': request.form['customerRef'],
            'Customer Name': request.form['customerName'],
            'City, State': request.form['cityState'],
            'Purchase Value (in words)': request.form['purchaseValue'],
            'Purchase Value Reduction (%)': request.form['purchaseReduction'],
            'Down Payment (%)': request.form['downPayment'],
            'Loan Period (Years)': request.form['loanPeriod'],
            'Annual Interest Rate (%)': request.form['annualInterest'],
            'Monthly Principal Reduction (%)': request.form['monthlyPrincipalReduction'],
            'Total Interest Reduction (%)': request.form['totalInterestReduction'],
            'Guarantor Name': request.form['guarantorName'],
            'Guarantor Reference Number': request.form['guarantorRef'],
        }
        
        formatted_data = calculate_loan_details(raw_data)
        
        if not formatted_data:
            raise ValueError("Failed to process loan data")
        
        entries = load_data()
        if 0 <= index < len(entries):
            entries[index] = formatted_data
            save_data(entries)
        
        return render_template('submit.html', message="Entry updated successfully!", entries=entries)
    
    except Exception as e:
        return render_template('submit.html', message=f"Error updating data: {str(e)}", entries=load_data())

@app.route('/delete/<int:index>', methods=['GET'])
def delete(index):
    try:
        entries = load_data()
        if 0 <= index < len(entries):
            entries.pop(index)
            save_data(entries)
            return render_template('submit.html', message="Entry deleted successfully!", entries=entries)
        return render_template('submit.html', message="Invalid entry index", entries=entries)
    except Exception as e:
        return render_template('submit.html', message=f"Error deleting data: {str(e)}", entries=load_data())

@app.route('/download')
def download():
    return send_from_directory(app.config['UPLOAD_FOLDER'], 'customer_loans.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)