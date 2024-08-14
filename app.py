
from flask import Flask, request, jsonify
import pandas as pd
import os

app = Flask(__name__)

@app.route('/add_person', methods=['POST'])
def add_person():
    data = request.form
    
    # Define the new row with all necessary attributes
    new_row = {
        'Family Name': data.get('family_name', ''),
        'First Name': data.get('first_name', ''),
        'Actual job title / position': data.get('job_title', ''),
        'Home Plant (LOKID)': data.get('home_plant', ''),
        'Contact information': data.get('contact_info', ''),
        'Qualified System Lead Auditor': 'x' if 'qualified_system_lead_auditor' in data else '',
        'Qualified System Co-Auditor': 'x' if 'qualified_system_co_auditor' in data else '',
        'Approved for following customers (System)': data.get('approved_customers_system', ''),
        'Qualified Process Lead Auditor': 'x' if 'qualified_process_lead_auditor' in data else '',
        'Qualified Process Co-Auditor': 'x' if 'qualified_process_co_auditor' in data else '',
        'Approved for following customers (Process)': data.get('approved_customers_process', ''),
        'Good knowledge of the LEON Instructions and other standards': 'x' if 'good_knowledge_leon' in data else '',
        'Has knowledge/experience in moderation, communication, leadership, project management': 'x' if 'moderation_communication' in data else '',
        'Has knowledge in conducting, planning, reporting and close out': 'x' if 'conducting_planning_reporting' in data else '',
        'Understands the automotive process approach, including risk-based thinking': 'x' if 'automotive_process_approach' in data else '',
        'Knowledge of CSR (Customer Specific Requirements)': 'x' if 'knowledge_csr' in data else '',
        'SPC Knowledge': 'x' if 'spc_knowledge' in data else '',
        'MSA Knowledge': 'x' if 'msa_knowledge' in data else '',
        'APQP Knowledge': 'x' if 'apqp_knowledge' in data else '',
        'PPAP Knowledge': 'x' if 'ppap_knowledge' in data else '',
        'FMEA Knowledge': 'x' if 'fmea_knowledge' in data else ''
    }
    
    # Create a DataFrame
    df = pd.DataFrame([new_row])
    
    # Define the output Excel file path
    output_file = 'new_qualified_person.xlsx'
    
    # Save the DataFrame to an Excel file
    try:
        df.to_excel(output_file, index=False)
        return jsonify({"message": "Person added successfully!", "file": output_file})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/plan_audit', methods=['POST'])
def plan_audit():
    data = request.form
    
    # Define the new row with all necessary attributes
    new_row = {
        'Site / BU / Department': data.get('site_bu_department', ''),
        'Processes': data.get('processes', ''),
        'Auditor': data.get('auditor', ''),
        'Co Auditor': data.get('co_auditor', ''),
        'Feb': data.get('feb', ''),
        'March': data.get('march', ''),
        'Apr': data.get('apr', ''),
        'May': data.get('may', ''),
        'June': data.get('june', ''),
        'July': data.get('july', ''),
        'Aug': data.get('aug', ''),
        'Sept': data.get('sept', ''),
        'Oct': data.get('oct', ''),
        'Nov': data.get('nov', ''),
        'Dec': data.get('dec', ''),
        'Jan': data.get('jan', ''),
        'Audit ID': data.get('audit_id', ''),
        'Result': data.get('result', ''),
        'Remark': data.get('remark', ''),
        'Status': data.get('status', '')
    }
    
    # Define the output Excel file path
    output_file = 'audit_plan.xlsx'
    
    # Check if the file exists
    if os.path.exists(output_file):
        # If the file exists, append the new row
        existing_df = pd.read_excel(output_file)
        new_df = pd.DataFrame([new_row])
        df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        # If the file does not exist, create a new DataFrame
        df = pd.DataFrame([new_row])
    
    # Save the DataFrame to an Excel file
    try:
        df.to_excel(output_file, index=False)
        return jsonify({"message": "Audit plan added successfully!", "file": output_file})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/update_audit', methods=['POST'])
def update_audit():
    data = request.form
    audit_id = data.get('audit_id', '')

    # Load the existing audit plan file
    try:
        df = pd.read_excel('audit_plan.xlsx')
    except Exception as e:
        return jsonify({"error": "Failed to read audit plan file: " + str(e)})

    # Find the row with the matching Audit ID
    if audit_id in df['Audit ID'].values:
        # Update the row with new values
        df.loc[df['Audit ID'] == audit_id, 'Site / BU / Department'] = data.get('site_bu_department', df.loc[df['Audit ID'] == audit_id, 'Site / BU / Department'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Processes'] = data.get('processes', df.loc[df['Audit ID'] == audit_id, 'Processes'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Auditor'] = data.get('auditor', df.loc[df['Audit ID'] == audit_id, 'Auditor'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Co Auditor'] = data.get('co_auditor', df.loc[df['Audit ID'] == audit_id, 'Co Auditor'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Feb'] = data.get('feb', df.loc[df['Audit ID'] == audit_id, 'Feb'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'March'] = data.get('march', df.loc[df['Audit ID'] == audit_id, 'March'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Apr'] = data.get('apr', df.loc[df['Audit ID'] == audit_id, 'Apr'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'May'] = data.get('may', df.loc[df['Audit ID'] == audit_id, 'May'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'June'] = data.get('june', df.loc[df['Audit ID'] == audit_id, 'June'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'July'] = data.get('july', df.loc[df['Audit ID'] == audit_id, 'July'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Aug'] = data.get('aug', df.loc[df['Audit ID'] == audit_id, 'Aug'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Sept'] = data.get('sept', df.loc[df['Audit ID'] == audit_id, 'Sept'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Oct'] = data.get('oct', df.loc[df['Audit ID'] == audit_id, 'Oct'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Nov'] = data.get('nov', df.loc[df['Audit ID'] == audit_id, 'Nov'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Dec'] = data.get('dec', df.loc[df['Audit ID'] == audit_id, 'Dec'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Jan'] = data.get('jan', df.loc[df['Audit ID'] == audit_id, 'Jan'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Result'] = data.get('result', df.loc[df['Audit ID'] == audit_id, 'Result'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Remark'] = data.get('remark', df.loc[df['Audit ID'] == audit_id, 'Remark'].values[0])
        df.loc[df['Audit ID'] == audit_id, 'Status'] = data.get('status', df.loc[df['Audit ID'] == audit_id, 'Status'].values[0])
    else:
        return jsonify({"error": "Audit ID not found."})
    
    # Save the updated DataFrame back to the Excel file
    try:
        df.to_excel('audit_plan.xlsx', index=False)
        return jsonify({"message": "Audit plan updated successfully!"})
    except Exception as e:
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    app.run(debug=True)
