import win32com.client
import os
import time


class PlotManager:
    def __init__(self, printer_name, output_dir):
        self.printer_name = printer_name
        self.output_dir = output_dir

        self.acad = win32com.client.GetActiveObject("AutoCAD.Application")
        self.doc = self.acad.ActiveDocument
        self.layout = self.doc.ActiveLayout

        self._configure_layout()

    def _configure_layout(self):
        self.layout.ConfigName = self.printer_name
        self.layout.UseStandardScale = True
        self.layout.StandardScale = 0  # Scale to fit
        self.layout.CenterPlot = True
        self.layout.PlotType = 4       # acWindow
        self.layout.StyleSheet = "monochrome.ctb"

    def plot_frames(self, frames, progress_callback=None, log_callback=None):
        count = 0
        total = len(frames)

        for index, frame in enumerate(frames, start=1):
            try:
                self._plot_single_frame(frame)

                count += 1

                if progress_callback:
                    progress_callback(index)

                if log_callback:
                    log_callback(
                        f"[{index}/{total}] Успешно: Лист {frame['sheet']}"
                    )

            except Exception as e:
                if log_callback:
                    log_callback(
                        f"[{index}/{total}] Ошибка листа {frame['sheet']}: {e}"
                    )

        return count

    def _plot_single_frame(self, frame):
        target_w = frame["w"]
        target_h = frame["h"]

        canonical_name = self._find_best_media(target_w, target_h)

        if not canonical_name:
            raise Exception(
                f"Формат {target_w}x{target_h} не найден в плоттере"
            )

        self.layout.CanonicalMediaName = canonical_name
        self.layout.PlotRotation = 1 if target_w > target_h else 0

        p1, p2 = frame["min"], frame["max"]

        point1 = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_ARRAY |
            win32com.client.pythoncom.VT_R8,
            p1[:2]
        )
        point2 = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_ARRAY |
            win32com.client.pythoncom.VT_R8,
            p2[:2]
        )

        self.layout.SetWindowToPlot(point1, point2)

        pdf_name = f"Лист_{frame['sheet']}_{target_w}x{target_h}.pdf"
        pdf_path = os.path.join(self.output_dir, pdf_name)

        self.doc.Plot.PlotToFile(pdf_path, self.printer_name)

        self._wait_for_file(pdf_path)

    def _find_best_media(self, width, height):
        best_match = None
        tolerance = 10

        for media in self.layout.GetCanonicalMediaNames():
            try:
                pw, ph = self.layout.GetPaperSize(media, 0, 0)

                direct = abs(pw - width) < tolerance and abs(ph - height) < tolerance
                swapped = abs(ph - width) < tolerance and abs(pw - height) < tolerance

                if direct or swapped:
                    best_match = media
                    break
            except Exception:
                continue

        return best_match

    def _wait_for_file(self, path, timeout=30):
        start_time = time.time()
        while not os.path.exists(path):
            if time.time() - start_time > timeout:
                raise Exception("Таймаут ожидания PDF")
            time.sleep(0.2)


def start_plot_process(printer_name, output_dir, frames,
                       progress_callback=None, log_callback=None):
    try:
        manager = PlotManager(printer_name, output_dir)
        return manager.plot_frames(
            frames,
            progress_callback=progress_callback,
            log_callback=log_callback
        )
    except Exception as e:
        return f"Критическая ошибка: {e}"
