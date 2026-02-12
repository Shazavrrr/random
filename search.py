import win32com.client
import re


class FrameAnalyzer:
    def __init__(self):
        self.acad = win32com.client.GetActiveObject("AutoCAD.Application")
        self.doc = self.acad.ActiveDocument
        self.model_space = self.doc.ModelSpace
        self.all_items = [self.model_space.Item(i) for i in range(self.model_space.Count)]

    def analyze(self):
        frames = []
        sheet_index = 1

        for obj in self.all_items:
            if obj.EntityName not in ("AcDbBlockReference", "AcDbExternalReference"):
                continue

            try:
                min_pt, max_pt = obj.GetBoundingBox()
            except Exception:
                continue

            w = round(abs(max_pt[0] - min_pt[0]), 0)
            h = round(abs(max_pt[1] - min_pt[1]), 0)
            short = min(w, h)
            long = max(w, h)

            fmt_name = self._detect_gost_format(short, long)

            if not fmt_name:
                continue

            if w < h and "x" not in fmt_name:
                fmt_name += " верт."

            sheet_number = self._find_sheet_number(min_pt, max_pt)

            frames.append({
                "sheet": sheet_index,
                "sheet_number": sheet_number,
                "format": fmt_name,
                "w": w,
                "h": h,
                "min": [min_pt[0], min_pt[1], min_pt[2]],
                "max": [max_pt[0], max_pt[1], max_pt[2]]
            })

            sheet_index += 1

        frames.sort(key=lambda x: (x["min"][0], -x["min"][1]))

        for i, frame in enumerate(frames):
            frame["sheet"] = i + 1

        return frames

    def _detect_gost_format(self, short, long):
        if 28000 <= short <= 31500:
            m = round(long / 21000)
            return "А4" if m == 1 else ("А3" if m == 2 else f"А4x{m}")

        elif 41000 <= short <= 43500:
            m = round(long / 29700)
            return "А3" if m == 1 else ("А2" if m == 2 else f"А3x{m}")

        elif 58000 <= short <= 61000:
            m = round(long / 42000)
            return "А2" if m == 1 else ("А1" if m == 2 else f"А2x{m}")

        elif 83000 <= short <= 86000:
            m = round(long / 59400)
            return "А1" if m == 1 else ("А0" if m == 2 else f"А1x{m}")

        elif 118000 <= short <= 121000:
            m = round(long / 84100)
            return "А0" if m == 1 else f"А0x{m}"

        elif 20000 <= short <= 22000:
            return "А4"

        return None

    def _find_sheet_number(self, min_pt, max_pt):
        corner = (max_pt[0], min_pt[1])  # правый нижний угол
        found_sheet = "???"

        for ent in self.all_items:
            if ent.EntityName not in ("AcDbText", "AcDbMText"):
                continue

            try:
                ins = ent.InsertionPoint
            except Exception:
                continue

            if min_pt[0] <= ins[0] <= max_pt[0] and min_pt[1] <= ins[1] <= max_pt[1]:
                dist = ((ins[0] - corner[0])**2 + (ins[1] - corner[1])**2) ** 0.5
                if dist < 16000:
                    val = re.sub(
                        r'\\P|\\f|\\H|\\A|\\W|{|}|\A1;|\.0000',
                        '',
                        ent.TextString
                    ).split(';')[-1].strip()

                    if re.match(r'^\d+(\.\d+)?$', val):
                        found_sheet = val
                        break

        return found_sheet


def analyze_to_json(json_path):
    import json

    try:
        analyzer = FrameAnalyzer()
        frames = analyzer.analyze()

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(frames, f, ensure_ascii=False, indent=4)

        return len(frames)

    except Exception as e:
        return f"Ошибка при анализе: {e}"
