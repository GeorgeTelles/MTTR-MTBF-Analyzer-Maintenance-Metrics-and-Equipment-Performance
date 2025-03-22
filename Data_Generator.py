import pandas as pd
import random
import datetime

def generate_equipment_data(num_equipments):
    equipment_types = ['High Pressure Pump', 'Gas Compressor', 'Control Valve', 'Gas Turbine', 'Electric Generator']
    models = ['X100', 'V200', 'M300', 'T400', 'B500', 'T600']
    locations = ['Distillation Section', 'Cracking Unit', 'Hydrogenation Unit', 'Solvent Recovery Section', 'Power Generation Area']
    
    equipments = []
    for i in range(1, num_equipments + 1):
        equipment = {
            'Equipment_ID': i,
            'Equipment_Name': f'{equipment_types[random.randint(0, len(equipment_types) - 1)]} {chr(65 + i % 26)}',
            'Model': random.choice(models),
            'Type': random.choice(equipment_types),
            'Location': random.choice(locations),
            'Acquisition_Date': str(datetime.date(random.randint(2015, 2022), random.randint(1, 12), random.randint(1, 28)))
        }
        equipments.append(equipment)
    
    return pd.DataFrame(equipments)

def generate_failure_data(num_failures, num_equipments):
    failure_causes = [
        'Pump system failure', 'Gas compressor failure', 'Control valve corrosion', 
        'Compressor overheating', 'Electric generator failure', 'Hydraulic system failure', 
        'Cooling system failure', 'Electrical failure', 'Fluid leakage', 'Vacuum system failure'
    ]
    maintenance_types = ['Corrective', 'Preventive', 'Predictive']
    failures = []
    
    for i in range(1, num_failures + 1):
        equipment_id = random.randint(1, num_equipments)
        failure_date = datetime.datetime(2025, random.randint(1, 12), random.randint(1, 28), random.randint(0, 23), random.randint(0, 59), random.randint(0, 59))
        downtime = random.randint(2, 10)  # hours of downtime
        repair_time = random.randint(2, 6)  # hours of repair
        resolution_date = failure_date + datetime.timedelta(hours=downtime + repair_time)
        
        failure = {
            'Failure_ID': i,
            'Equipment_ID': equipment_id,
            'Failure_Date': failure_date,
            'Resolution_Date': resolution_date,
            'Downtime': f'{downtime} hours',
            'Failure_Cause': random.choice(failure_causes),
            'Maintenance_Type': random.choice(maintenance_types),
            'Repair_Time': f'{repair_time} hours',
            'Maintenance_Cost': f'$ {random.randint(200, 5000)}.00'
        }
        failures.append(failure)
    
    return pd.DataFrame(failures)

def generate_preventive_maintenance_data(num_maintenance, num_equipments):
    preventive_maintenance = []
    descriptions = ['Oil change', 'Seal replacement', 'Safety inspection', 'Filter cleaning', 'Replacement of damaged parts']
    
    for i in range(1, num_maintenance + 1):
        equipment_id = random.randint(1, num_equipments)
        maintenance_date = datetime.datetime(2025, random.randint(1, 12), random.randint(1, 28), random.randint(0, 23), random.randint(0, 59))
        maintenance = {
            'Maintenance_ID': i,
            'Equipment_ID': equipment_id,
            'Maintenance_Date': maintenance_date,
            'Maintenance_Type': 'Preventive',
            'Description': random.choice(descriptions),
            'Maintenance_Cost': f'$ {random.randint(100, 2000)}.00'
        }
        preventive_maintenance.append(maintenance)
    
    return pd.DataFrame(preventive_maintenance)

num_equipments = 100
num_failures = 400
num_maintenance = 200  # Example of preventive maintenance

equipments_df = generate_equipment_data(num_equipments)
failures_df = generate_failure_data(num_failures, num_equipments)
preventive_maintenance_df = generate_preventive_maintenance_data(num_maintenance, num_equipments)

with pd.ExcelWriter('Equipment_and_Failures_Refinery.xlsx') as writer:
    equipments_df.to_excel(writer, sheet_name='Equipments', index=False)
    failures_df.to_excel(writer, sheet_name='Equipment_Failures', index=False)
    preventive_maintenance_df.to_excel(writer, sheet_name='Preventive_Maintenance', index=False)

print("Spreadsheet generated successfully!")