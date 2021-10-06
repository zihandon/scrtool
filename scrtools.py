#!/usr/bin/env python3

import io
import os
import sys

import pyperclip
import win32clipboard
from io import BytesIO

from PIL import Image
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QSystemTrayIcon,qApp,QMenu,QAction

from system_hotkey import SystemHotkey
from PyQt5.QtCore import QObject,pyqtSignal
import system_hotkey

#联系
from aboutme import MyQrWidget as Qr 

#切换baidu paddleocr
# import numpy as np
# from paddleocr import PaddleOCR

#GLOBAL_PADDLE_OCR = None

class Snipper(QtWidgets.QWidget):
    #定义一个热键信号
    sig_keyhot = pyqtSignal(str)
    flag_captable = False

    def __init__(self, parent=None, flags=Qt.WindowFlags()):
        super().__init__(parent=parent, flags=flags)

        self.hk_capture, self.hk_exit, self.hk_table, self.hk_pic = SystemHotkey(),SystemHotkey(),SystemHotkey(),SystemHotkey(),
        self.hk_capture.register(('alt','c'),callback=lambda x:self.send_key_event("capture"))
        self.hk_table.register(('alt','t'),callback=lambda x:self.send_key_event("table"))
        self.hk_exit.register(('alt','q'),callback=lambda x:self.send_key_event("exit"))
        self.hk_pic.register(('alt','x'),callback=lambda x:self.send_key_event("pic"))
        self.sig_keyhot.connect(self.hotkey_process)

    #热键信号发送函数(将外部信号，转化成qt信号)
    def send_key_event(self,i_str):
        self.sig_keyhot.emit(i_str)

    def hotkey_process(self, i_str):
        self.flag_captable = False
        self.flag_capic = False
        if i_str == "capture":
            self.capture()
        if i_str == "table":
            self.flag_captable = True
            self.capture()
        elif i_str == "exit":
            self.quit()
        elif i_str == "aboutme":
            self.aboutme()
        elif i_str =="pic":
            self.flag_capic = True
            self.capture()
        else:
            pass
    
    def aboutme(self):
        self.aboutmeWin = Qr()
        self.aboutmeWin.show()
        global tp
        tp.show()

    def quit(self):
        print(f"INFO: inteltools exit!")
        QtWidgets.QApplication.quit()

    def capture(self):
        print(f"INFO: start capture!")

        self.setWindowFlags(
            Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Dialog
        )
        self.setWindowState(self.windowState() | Qt.WindowFullScreen)
        self.show()
        
        self.start, self.end = QtCore.QPoint(), QtCore.QPoint()
        self.screen = QtWidgets.QApplication.screenAt(QtGui.QCursor.pos()).grabWindow(0)
        palette = QtGui.QPalette()
        palette.setBrush(self.backgroundRole(), QtGui.QBrush(self.screen))
        self.setPalette(palette)

        QtWidgets.QApplication.setOverrideCursor(QtGui.QCursor(QtCore.Qt.CrossCursor))
        
        pass

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            QtWidgets.QApplication.quit()

        return super().keyPressEvent(event)

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QtGui.QColor(0, 0, 0, 100))
        painter.drawRect(0, 0, self.width(), self.height())

        if self.start == self.end:
            return super().paintEvent(event)

        painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255), 3))
        painter.setBrush(painter.background())
        painter.drawRect(QtCore.QRect(self.start, self.end))
        return super().paintEvent(event)

    def mousePressEvent(self, event):
        self.start = self.end = event.pos()
        self.update()
        return super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        self.end = event.pos()
        self.update()
        return super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if self.start == self.end:
            return super().mouseReleaseEvent(event)

        self.hide()
        QtWidgets.QApplication.processEvents()
        shot = self.screen.copy(QtCore.QRect(self.start, self.end))

        if self.flag_captable:
            processImage2excel_pdocr(shot)
        elif self.flag_capic:
            #将shot放入剪贴板
            shot.save('shot.png','PNG')
            image = Image.open('shot.png')
            output = BytesIO()
            image.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            notify(f'已完成截图，并存入剪切板\n')
            
        else:
            processImage_pdocr(shot)
        #QtWidgets.QApplication.quit()

        QtWidgets.QApplication.setOverrideCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))

def processImage_pdocr(img):
    buffer = QtCore.QBuffer()
    buffer.open(QtCore.QBuffer.ReadWrite)
    img.save(buffer, "PNG")
    pil_img = Image.open(io.BytesIO(buffer.data()))
    buffer.close()

    try:
        import numpy as np
        from paddleocr import PaddleOCR

        model_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ocr')
        paddleocr_engine = PaddleOCR(use_gpu=False,det_model_dir= os.path.join(model_path,'det'), \
            cls_model_dir=os.path.join(model_path,'cls'), \
            rec_model_dir=os.path.join(model_path,'rec'), \
            rec_char_dict_path=os.path.join(model_path,'ppocr_keys_v1.txt'))

        paddleocr_engine = PaddleOCR(show_log=False, use_gpu=False)

        ocr_result = paddleocr_engine.ocr(np.array(pil_img))
        ocr_result = [line[1][0] for line in ocr_result]
        result = '\n'.join(ocr_result)

    except RuntimeError as error:
        print(f"ERROR: An paddleocr error occurred when trying to process the image: {error}")
        notify(f"An error paddleocr occurred when trying to process the image: {error}")
        return

    if result:
        #result = ''.join(result.split(' '))
        pyperclip.copy(result)
        print(f'INFO: Copied result to the clipboard')
        #notify(f'Copied "{result}" to the clipboard')
        notify(f'识别成功！文字识别结果已保存到剪贴板\n')
    else:
        print(f"INFO: Unable to read text from image, did not copy")
        notify(f"Unable to read text from image, did not copy")

def processImage2excel_pdocr(img):
    buffer = QtCore.QBuffer()
    buffer.open(QtCore.QBuffer.ReadWrite)
    img.save(buffer, "PNG")
    pil_img = Image.open(io.BytesIO(buffer.data()))
    buffer.close()

    try:
        import numpy as np
        from paddleocr import PPStructure,draw_structure_result,save_structure_res
        table_engine = PPStructure(show_log=False, use_gpu=False)

        model_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ocr')
        table_engine = PPStructure(show_log=False, use_gpu=False, \
                det_model_dir= os.path.join(model_path,'det'), \
                e2e_char_dict_path = os.path.join(model_path,'table/dict/ic15_dict.txt'), \
                cls_model_dir=os.path.join(model_path,'cls'), \
                rec_model_dir=os.path.join(model_path,'rec'), \
                table_model_dir = os.path.join(model_path,'table'), \
                table_char_dict_path = os.path.join(model_path,'table/dict/table_structure_dict.txt'), \
                rec_char_dict_path=os.path.join(model_path,'ppocr_keys_v1.txt'), \
                vis_font_path = os.path.join(model_path,'table/fonts/simfang.ttf'))

        result = table_engine(np.array(pil_img.convert("RGB")))
        save_structure_res(result, os.path.dirname(os.path.abspath(__file__)), 'table')

        
        print(f'INFO: get table is ok!')  
        notify(f'识别成功！表格识别结果已保存到table文件夹\n')

    except RuntimeError as error:
        print(f"ERROR: An paddleocr error occurred when trying to process the image: {error}")
        notify(f"An error paddleocr occurred when trying to process the image: {error}")
        return

def notify(msg):
    global tp
    tp.show()
    tp.showMessage("智能截取工具", msg, QtWidgets.QSystemTrayIcon.NoIcon)
    #tp.hide()
    
if __name__ == "__main__":
 
    print(f"INFO: inteltools is running...")

    #QtWidgets.QApplication.setQuitOnLastWindowClosed(False) 
    QtCore.QCoreApplication.setAttribute(Qt.AA_DisableHighDpiScaling)
    app = QtWidgets.QApplication(sys.argv)
    window = QtWidgets.QMainWindow()
    snipper = Snipper(window)
    snipper.show()
        
    # 在系统托盘处显示图标
    tp = QSystemTrayIcon(window)
    tp.setIcon(QIcon('tools.ico'))
    tp.setToolTip(u'Alt+X截图，Alt+C文字识别，Alt+T表格，Alt+Q退出')

    picAct = QAction('截图(&Screen)',triggered = lambda x:snipper.send_key_event('pic'))
    capAct = QAction('文字识别(&OCR)',triggered = lambda x:snipper.send_key_event('capture'))
    tblAct = QAction('表单(&Table)',triggered = lambda x:snipper.send_key_event('table'))
    extAct = QAction('退出(&Quit)',triggered = lambda x: tp.setVisible(False) or snipper.send_key_event('exit'))
    donAct = QAction('关于作者(&about me)',triggered = lambda x: tp.setVisible(False) or snipper.send_key_event('aboutme'))

    tpMenu = QMenu()
    tpMenu.addAction(picAct)
    tpMenu.addAction(capAct)
    tpMenu.addAction(tblAct)
    tpMenu.addAction(extAct)
    tpMenu.addAction(donAct)
    tp.setContextMenu(tpMenu)
    tp.show()

    notify(u"\nAlt+X：截图，\nAlt+C：识别文字，\nAlt+T：识别表格，\nAlt+Q：退出\n")


    sys.exit(app.exec_())

