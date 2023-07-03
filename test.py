from pyautocad import Autocad

import win32com.client
import csv

acad = win32com.client.Dispatch("AutoCAD.Application")

circuit = []

for entity in acad.ActiveDocument.ModelSpace:
    name = entity.Entityname
    
    # matching default autocad block reference
    if name == "AcDbBlockReference":
        HasAttributes = entity.HasAttributes

        if HasAttributes:
            for attrib in list(entity.GetAttributes()):
                circuit.append(attrib.TextString)


circuits = list(set(circuit))
sorted_circuits = sorted(circuits, reverse=False)


with open("f6.csv", "w") as f:
    writer = csv.writer(f,delimiter='\n')
    writer.writerows(sorted_circuits)
