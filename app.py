
import win32com.client
import csv


# using the IDispatch interface of win32com to access the autocad com object
acad = win32com.client.Dispatch("AutoCAD.Application") 

circuits = []

for entity in acad.ActiveDocument.ModelSpace:
    name = entity.Entityname
    
    # matching default autocad block reference
    if name == "AcDbBlockReference":
        HasAttributes = entity.HasAttributes

        if HasAttributes:
            for attrib in list(entity.GetAttributes()):
                circuits.append(attrib.TextString)


circuits = list(set(circuits))
sorted_circuits = sorted(circuits, reverse=False)

writer_path = "f6.csv" 

with open(writer_path, "w") as f:
    writer = csv.writer(f,delimiter='\n')
    writer.writerow(sorted_circuits)
