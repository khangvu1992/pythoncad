import win32com.client
import time

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

name = "SS_" + str(int(time.time()))
ss = doc.SelectionSets.Add(name)
ss.SelectOnScreen()

def dump_object(obj):
    print("\n" + "#" * 80)
    print("TYPE:", obj.ObjectName)

    props = dir(obj)

    for p in props:
        if p.startswith("_"):
            continue

        try:
            value = getattr(obj, p)

            # bỏ method
            if callable(value):
                continue

            print(f"{p} = {value}")

        except Exception:
            pass


for obj in ss:
    dump_object(obj)
