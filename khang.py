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


print("\nTYPE:", obj_pick.ObjectName)

for attr in dir(obj_pick):
    if attr.startswith("_"):
        continue

    try:
        value = getattr(obj_pick, attr)

        if callable(value):
            continue

        print(f"{attr:<30} = {value} ({type(value).__name__})")

    except Exception as e:
        print(f"{attr:<30} = <ERROR>")


print("Layer được chọn:", obj_pick)
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



# ===== hàm lấy tọa độ =====
def get_coordinates(obj):
    try:
        # LINE
        if obj.ObjectName == "AcDbLine":
            return {
                "StartPoint": obj.StartPoint,
                "EndPoint": obj.EndPoint
            }

        # CIRCLE
        elif obj.ObjectName == "AcDbCircle":
            return {
                "Center": obj.Center,
                "Radius": obj.Radius
            }

        # ARC
        elif obj.ObjectName == "AcDbArc":
            return {
                "Center": obj.Center,
                "StartAngle": obj.StartAngle,
                "EndAngle": obj.EndAngle
            }

        # TEXT / MTEXT
        elif obj.ObjectName in ["AcDbText", "AcDbMText"]:
            return {
                "InsertionPoint": obj.InsertionPoint
            }

        # POLYLINE
        elif obj.ObjectName in ["AcDbPolyline", "AcDb2dPolyline"]:
            return {
                "Coordinates": obj.Coordinates
            }

        # BLOCK
        elif obj.ObjectName == "AcDbBlockReference":
            return {
                "InsertionPoint": obj.InsertionPoint,
                "Name": obj.Name
            }

        else:
            # fallback nếu có Coordinates
            if hasattr(obj, "Coordinates"):
                return {"Coordinates": obj.Coordinates}

    except:
        pass

    return {"Info": "Không lấy được tọa độ"}

print("\n TỌA ĐỘ OBJECT:")

for i, obj in enumerate(filtered[:1000]):  # in 1000 cái đầu
    coords = get_coordinates(obj)
    print(f"\n--- Object {i+1} ---")
    print("Type:", obj.ObjectName)
    print("Layer:", obj.Layer)

    for k, v in coords.items():
        print(f"{k}: {v}")


# cleanup
ss.Delete()
