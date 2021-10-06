import io
import base64
from PIL import Image,ImageQt
from PyQt5 import QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog,QLabel,QHBoxLayout

#from donate_qr import img as qr_png

class MyQrWidget(QDialog):
    def __init__(self):
        super().__init__()

        self.setWindowFlags(
            Qt.WindowCloseButtonHint
        )
        self.setWindowTitle(u"关于作者")
        self.hbox = QHBoxLayout (self)

        # pil_img = Image.open(io.BytesIO(base64.b64decode(qr_png)))
        # pp = ImageQt.toqpixmap(pil_img)
   
        # label = QLabel()
        # # label.setPixmap(QtGui.QPixmap(pp))
        # # label.setScaledContents(True)
        # label.setText('作者：唐光溥')
        # self.hbox.addWidget(label)

        # label1=QLabel()
        # label1.setText('wechat:MercedesTang')
        # self.hbox.addWidget(label1)

        label2=QLabel()
        label2.setPixmap(ImageQt.QPixmap('1.png'))
        self.hbox.addWidget(label2)




    def closeEvent(self, event):
        event.ignore()  # 忽略关闭事件
        self.hide()  # 隐藏窗体