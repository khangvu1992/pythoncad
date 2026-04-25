import win32com.client

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

print("\n👉 Hãy chọn 1 object để lấy layer")

try:
    obj = doc.Utility.GetEntity()[0]
except:
    print("❌ Hủy chọn")
    exit()

layer_name = obj.Layer
print("Layer được chọn:", layer_name)
