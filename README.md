# autocad-attributes
A python script i used to get drawing attributes from an autocad file

## How it works
This script uses IDispatch interface from wincom32 in creating an Autocad COM object.
[Introduction to COM concept - Read Chapter 5](https://drive.google.com/file/d/1yMonHygGIsx1JUVYl-GysBZ95kL50ncy/view?usp=sharing)

After the object has been created we access the active document using autocad's VBA (ActiveDocument).

We then get the `ActiveModelSpace` [Autocad Document Object](https://help.autodesk.com/view/ACD/2023/ENU/?guid=GUID-9216BFCD-D358-4FC6-B631-B52E6693D242)

Then we access the access the `EntityName` and `GetAttributes()`


## Usage
1. Open up the Autocad file you wish to get attributes from
2. run `app.py`
3. Check the csv file for `writer_path` to see attributes listed


### Other Recources
1. [AutoCAD Document object in pyautocad](https://www.supplychaindataanalytics.com/autocad-document-object-in-pyautocad/)





