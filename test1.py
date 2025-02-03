import pandas as pd
import xlsxwriter
from RFEM.initModel import Model, Client
from RFEM.Results.resultTables import ResultTables
from RFEM.enums import CaseObjectType, ObjectTypes
from RFEM.Tools.GetObjectNumbersByType import GetObjectNumbersByType

client = Client('http://localhost:8081/wsdl')
model_name = client.service.get_model_list().name
model = Model(False, model_name[0])
excel_data = []
surf_nums = GetObjectNumbersByType(ObjectTypes.E_OBJECT_TYPE_SURFACE)
surface_groups = {}
for j in surf_nums:
    x = round(Model.clientModel.service.get_surface(j).center_of_gravity_x, 2)
    y = round(Model.clientModel.service.get_surface(j).center_of_gravity_y, 2)
    pos = "XY" in Model.clientModel.service.get_surface(j).position_short
    if pos:
        continue

    key = (x, y)
    
    if key not in surface_groups:
        surface_groups[key] = []
    
    surface_groups[key].append(j)

for key, surfaces in surface_groups.items():
    for i in surfaces:
        data = ResultTables.SurfacesBasicInternalForces(CaseObjectType.E_OBJECT_TYPE_DESIGN_SITUATION, 1, i)
        z_values = sorted(set(d.get('grid_point_coordinate_z', None) for d in data if d.get('grid_point_coordinate_z') is not None))
        if z_values:
            median_z = z_values[len(z_values) // 2]
            filtered_data = [d for d in data if d.get('grid_point_coordinate_z') == median_z]
            forces = [item['axial_force_ny'] for item in filtered_data]
            forces.sort()
            N_min_d = round(forces[0] / 1000, 1)
            N_max_d = round(forces[-1] / 1000, 1)

            excel_data.append([i, N_min_d, N_max_d])

df = pd.DataFrame(excel_data, columns=['Wall number', 'N_min,d', 'N_max,d'])

df.to_excel('surface_forces.xlsx', index=False, engine='xlsxwriter')

print("Done!")
