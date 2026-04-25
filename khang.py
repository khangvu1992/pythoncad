import win32com.client
import time

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

print("\n👉 Hãy chọn 1 object để lấy layer")

try:
    obj_pick = doc.Utility.GetEntity()[0]
except:
    print("❌ Hủy chọn")
    exit()

layer_name = obj_pick.Layer
print("Layer được chọn:", layer_name)

# ===== tạo selection set =====
name = "SS_" + str(int(time.time()))

# xóa nếu trùng
for s in doc.SelectionSets:
    if s.Name == name:
        s.Delete()

ss = doc.SelectionSets.Add(name)

print("\n👉 Quét đối tượng...")
ss.SelectOnScreen()

print(f"\nTổng object đã chọn: {ss.Count}")

# ===== lọc theo layer =====
filtered = []

for obj in ss:
    try:
        if obj.Layer == layer_name:
            filtered.append(obj)
    except:
        pass

# ===== kết quả =====
print(f"\n✅ Có {len(filtered)} object thuộc layer {layer_name}")

# ===== in thử =====
for obj in filtered[:10]:
    print(obj.ObjectName, "| Layer:", obj.Layer)

# cleanup
ss.Delete()
