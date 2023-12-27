from PyQt6 import QtWidgets,uic
from PyQt6.QtCore import QDate
import sys
from PyQt6.QtWidgets import *
import pyodbc
from openpyxl import Workbook

conn_str = (
    "Driver={SQL Server};"
    "Server=DESKTOP-IUVTIEA\SQLEXPRESS;"
    "Database=Final;"
    "Trusted_Connection=yes"
)
db = pyodbc.connect(conn_str)

#Giao diện chọn CÁ NHÂN hoặc NHÓM
class OpenW(QMainWindow):
    def __init__(self):
        super(OpenW,self).__init__()
        uic.loadUi("OpenW.ui",self)
        self.btCaNhan.clicked.connect(self.CaNhan)
        self.btNhom.clicked.connect(self.Nhom)

    def Nhom(self):
        widget.setFixedSize(425,320)
        widget.setCurrentIndex(2)
        # widget.show()
        # app.exec()

    def CaNhan(self):
        widget.setFixedSize(650,600)
        widget.setCurrentIndex(3)
#Giao diện làm việc chính cho CÁ NHÂN
class CaNhanW(QMainWindow):
    def __init__(self):
        super(CaNhanW,self).__init__()
        uic.loadUi("CaNhanMain.ui",self)
        self.setFixedSize(650, 600)
        self.ngay.setDate(QDate.currentDate())
        self.ngay_2.setDate(QDate.currentDate())
        self.ngay_2.setVisible(False)
        self.Load_data()
        self.tbCaNhan.itemSelectionChanged.connect(self.on_item_selection_changed)
        self.cbBox.currentIndexChanged.connect(self.on_combobox_changed)
        self.btThem.clicked.connect(self.Them_data)
        self.btSua.clicked.connect(self.Sua_data)
        self.btXoa.clicked.connect(self.Xoa_data)
        self.btHT.clicked.connect(self.HoanT)
        self.btTim.clicked.connect(self.Tim_data)
        self.btEx.clicked.connect(self.export_to_excel)
        self.btBack.clicked.connect(self.home)

    def home(self):
        widget.setFixedSize(310,180)
        widget.setCurrentIndex(0)

    def Load_data(self):
        self.tbCaNhan.clear()

        query=db.cursor()
        query.execute("select * from CaNhan")
        kq=query.fetchall()
        self.tbCaNhan.setColumnHidden(0, True)
        self.tbCaNhan.setColumnCount(5)
        self.tbCaNhan.setColumnWidth(0, 30)
        self.tbCaNhan.setColumnWidth(1, 90) #10
        self.tbCaNhan.setColumnWidth(2, 130) #20
        self.tbCaNhan.setColumnWidth(3, 150) #50
        self.tbCaNhan.setColumnWidth(4, 100) #20
        self.tbCaNhan.setColumnWidth(6, 100)
        header_labels = ['Stt', 'Ngày giao', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
        self.tbCaNhan.setHeaderLabels(header_labels)

        for row in kq:
            item = QTreeWidgetItem(self.tbCaNhan)
            for i, value in enumerate(row):
                item.setText(i, str(value))

    def Them_data(self):
        nd=self.txtNoiDung.text()
        ngay=self.ngay.date().toString("yyyy-MM-dd")
        mota=self.txtMoTa.text()
        trangthai='Chua Hoan Thanh'
        if not nd:
            QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần Nội Dung!')
            return
        if not mota:
            QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần Mô Tả')
            return
        query=db.cursor()
        query.execute(f"INSERT INTO CaNhan (tg, noidung, mota, trangthai) VALUES ( '{ngay}', '{nd}', '{mota}', '{trangthai}')")
        db.commit()

        self.Load_data()
        self.txtNoiDung.clear()
        self.txtMoTa.clear()

    def Sua_data(self):
        selected_items = self.tbCaNhan.selectedItems()
        if selected_items:
            selected_item = selected_items[0].text(0)
            nd=self.txtNoiDung.text()
            ngay=self.ngay.date().toString("yyyy-MM-dd")
            mota=self.txtMoTa.text()
            #trangthai='Chua Hoan Thanh'
            if not nd:
                QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần Nội Dung!')
                return
            if not mota:
                QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần Mô Tả')
                return
            query=db.cursor()
            query.execute("UPDATE CaNhan SET tg=?, noidung=?, mota=? WHERE stt=?", ngay, nd, mota, selected_item)
            db.commit()

            self.Load_data()
            self.txtNoiDung.clear()
            self.txtMoTa.clear()
        else:
            QMessageBox.information(self,'Thông Báo','Bạn chưa chọn dòng muốn sửa')

    def Xoa_data(self):
        selected_items = self.tbCaNhan.selectedItems()
        if selected_items: 
            selected_item = selected_items[0].text(0)

            query=db.cursor()
            query.execute("DELETE FROM CaNhan WHERE stt=?", selected_item)
            db.commit()

            self.Load_data()
        else:
            QMessageBox.information(self,'Thông Báo','Bạn chưa chọn dòng muốn xóa')  

    def HoanT(self):
        selected_items = self.tbCaNhan.selectedItems()
        if selected_items:
            selected_item = selected_items[0].text(0)
            ngay=self.ngay_2.date().toString("yyyy-MM-dd")
            time=None
            if selected_items[0].text(4)=="Hoan Thanh":
                query=db.cursor()
                query.execute("UPDATE CaNhan SET trangthai='Chua Hoan Thanh',tht=? WHERE stt=?",time, selected_item)
                db.commit()
            if selected_items[0].text(4)=="Chua Hoan Thanh":
                query=db.cursor()
                query.execute("UPDATE CaNhan SET trangthai='Hoan Thanh',tht=? WHERE stt=?",ngay, selected_item)
                db.commit()
            self.Load_data()
        else:
            QMessageBox.information(self,'Thông Báo','Bạn chưa chọn dòng')
    
    def on_item_selection_changed(self):
        selected_items = self.tbCaNhan.selectedItems()
        if selected_items:
            if selected_items[0].text(4)=="Hoan Thanh":
                self.btHT.setText("Làm Lại")
            else:
                self.btHT.setText("Hoàn thành")
            self.txtNoiDung.setText(selected_items[0].text(2))
            self.txtMoTa.setText(selected_items[0].text(3))
            self.ngay.setDate(QDate.fromString(selected_items[0].text(1),"yyyy-MM-dd"))

    def on_combobox_changed(self):
        selected_text = self.cbBox.currentText()

        if selected_text == "Tat Ca":
            self.Load_data()
        elif selected_text == "Hoan Thanh":
            self.tbCaNhan.clear()

            query=db.cursor()
            query.execute("select * from CaNhan Where trangthai='Hoan Thanh'")
            kq=query.fetchall()
            self.tbCaNhan.setColumnHidden(0, True)
            self.tbCaNhan.setColumnCount(5)
            self.tbCaNhan.setColumnWidth(0, 30)
            self.tbCaNhan.setColumnWidth(1, 90) #10
            self.tbCaNhan.setColumnWidth(2, 130) #20
            self.tbCaNhan.setColumnWidth(3, 150) #50
            self.tbCaNhan.setColumnWidth(4, 100) #20
            self.tbCaNhan.setColumnWidth(6, 100)
            header_labels = ['Stt', 'Ngày giao', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
            self.tbCaNhan.setHeaderLabels(header_labels)

            for row in kq:
                item = QTreeWidgetItem(self.tbCaNhan)
                for i, value in enumerate(row):
                    item.setText(i, str(value))
        elif selected_text == "Chua Hoan Thanh":
            self.tbCaNhan.clear()

            query=db.cursor()
            query.execute("select * from CaNhan Where trangthai='Chua Hoan Thanh'")
            kq=query.fetchall()
            self.tbCaNhan.setColumnHidden(0, True)
            self.tbCaNhan.setColumnCount(5)
            self.tbCaNhan.setColumnWidth(0, 30)
            self.tbCaNhan.setColumnWidth(1, 90) #10
            self.tbCaNhan.setColumnWidth(2, 130) #20
            self.tbCaNhan.setColumnWidth(3, 150) #50
            self.tbCaNhan.setColumnWidth(4, 100) #20
            self.tbCaNhan.setColumnWidth(6, 100)
            header_labels = ['Stt', 'Ngày giao', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
            self.tbCaNhan.setHeaderLabels(header_labels)

            for row in kq:
                item = QTreeWidgetItem(self.tbCaNhan)
                for i, value in enumerate(row):
                    item.setText(i, str(value))

    def Tim_data(self):
        search_text = self.txtTim.text()
        if search_text:
            self.tbCaNhan.clear()

            query=db.cursor()
            sql_command = "SELECT * FROM CaNhan WHERE noidung LIKE ?"
            query.execute(sql_command, ('%' + search_text + '%',))
            kq=query.fetchall()
            self.tbCaNhan.setColumnHidden(0, True)
            self.tbCaNhan.setColumnCount(5)
            self.tbCaNhan.setColumnWidth(0, 30)
            self.tbCaNhan.setColumnWidth(1, 90) #10
            self.tbCaNhan.setColumnWidth(2, 130) #20
            self.tbCaNhan.setColumnWidth(3, 150) #50
            self.tbCaNhan.setColumnWidth(4, 100) #20
            self.tbCaNhan.setColumnWidth(6, 100)
            header_labels = ['Stt', 'Ngày giao', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
            self.tbCaNhan.setHeaderLabels(header_labels)

            for row in kq:
                item = QTreeWidgetItem(self.tbCaNhan)
                for i, value in enumerate(row):
                    item.setText(i, str(value))
        else:
            self.Load_data()
        
        self.txtTim.clear()

    def export_to_excel(self):
        # Tạo một Workbook
        workbook = Workbook()
        # Tạo một WorkSheet
        worksheet = workbook.active
        # Đặt tiêu đề cột
        columns = ["Ngày giao", "Nội dung", "Loại","Trạng thái", "Ngày HT"]
        worksheet.append(columns)
        # Lặp qua dữ liệu QTreeWidget và thêm vào WorkSheet
        for row in range(self.tbCaNhan.topLevelItemCount()):
            item = self.tbCaNhan.topLevelItem(row)
            row_data = [item.text(col) for col in range(1,item.columnCount())]
            worksheet.append(row_data)
        file=self.txtEx.text()
        if not file:
            QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần tên cho File Excel')
            return
        # Lưu Workbook thành tệp Excel
        workbook.save(file+".xlsx")
        QMessageBox.information(self,'Thông Báo','Đã lưu ra File Excel')

#Giao diện đăng ký cho NHÓM
class DKApp(QMainWindow):
    def __init__(self):
        super(DKApp,self).__init__()
        uic.loadUi("DangKy.ui",self)
        self.btDk_dk.clicked.connect(self.DK)
        self.btThoat.clicked.connect(self.Thoat)

    def Thoat(self):
        widget.setFixedSize(310,180)
        widget.setCurrentIndex(0)

    def DK(self):
        un=self.txtTk_dk.text()
        psw=self.txtMk_dk.text()
        hoten=self.txthoten.text()
        sdt=self.txtsdt.text()
        if not un:
            QMessageBox.information(self, 'Thông báo', 'Vui lòng nhập Tài Khoản')
            return
        if not psw:
            QMessageBox.information(self, 'Thông báo', 'Vui lòng nhập Mật khẩu')
            return
        if not hoten:
            QMessageBox.information(self, 'Thông báo', 'Vui lòng nhập Họ tên')
            return
        if not sdt:
            QMessageBox.information(self, 'Thông báo', 'Vui lòng nhập Số điện thoại')
            return
        query=db.cursor()
        query.execute("select * from TaiKhoan where tk='"+un+"' ")
        kt=query.fetchone()
        if kt:
            QMessageBox.information(self,"Register Output","Đăng ký thất bại")
        else:
            query.execute("insert into TaiKhoan values ('"+un+"','"+psw+"','customer','"+hoten+"','"+sdt+"')")
            db.commit()
            QMessageBox.information(self,"Register Output","Đăng ký thành công")
            widget.setFixedSize(425,320)
            widget.setCurrentIndex(2)

#Giao diện làm việc chính cho ADMIN_NHÓM
class AdminApp(QMainWindow):
    
    def __init__(self):
        super(AdminApp,self).__init__()
        uic.loadUi("adminW.ui",self)
        self.txtte.setVisible(False)
        self.tabWidget.setCurrentIndex(2)
        self.tabWidget.tabBar().setTabText(0, "")
        self.tabWidget.tabBar().setTabText(1, "")
        self.tabWidget.setTabEnabled(0, False)
        self.tabWidget.setTabEnabled(1, False)
        self.btDn.clicked.connect(self.DN)
        self.btDk.clicked.connect(self.DangKy)
        self.tbNhom.itemSelectionChanged.connect(self.on_item_selection_changed)
        self.cbBox.currentIndexChanged.connect(self.on_combobox_changed)
        self.btHT.clicked.connect(self.HoanT)
        self.btTim.clicked.connect(self.Tim_data)
        self.btEx.clicked.connect(self.export_to_excel)
        self.btBack.clicked.connect(self.Logout)
        self.btThoat.clicked.connect(self.Thoat)
        #Phần Admin
        self.ngay_2.setDate(QDate.currentDate())
        self.ngay_2.setVisible(False)
        self.tbKeHoach.itemSelectionChanged.connect(self.on_item_selection_changed2)
        self.cbBox_2.currentIndexChanged.connect(self.on_combobox_changed_KH)
        self.ngay.setDate(QDate.currentDate())
        self.Load_data2()
        self.Load_data3()
        self.btKeHoach.clicked.connect(self.HideWgKH)
        self.btNhiemVu.clicked.connect(self.HideWgNV)
        self.btThem.clicked.connect(self.Them_data)
        self.btSua.clicked.connect(self.Sua_data)
        self.btXoa.clicked.connect(self.Xoa_data)
        self.add_Item()
        self.cbTK.currentIndexChanged.connect(self.update_hoten)
        self.cbMaKH.currentIndexChanged.connect(self.update_hoten)
        self.btThem_2.clicked.connect(self.Them_data2)
        self.btXoa_2.clicked.connect(self.Xoa_data2)
        self.btSua_2.clicked.connect(self.Sua_data2)
        self.cbBox_2.setVisible(False)
        self.load_cbbox_makh_2()
        self.cbMaKH_2.currentIndexChanged.connect(self.Load_data_makh2)
        self.btTimND.clicked.connect(self.Tim_data_NV)

    def Tim_data_NV(self):
        search_text = self.txtNDTim.text()
        if search_text:
            self.tbNhiemVu.clear()

            query=db.cursor()
            query.execute("SELECT NVnKH.idNVKH,TaiKhoan.tk,TaiKhoan.hoten,KeHoach.idkh,KeHoach.noidung FROM TaiKhoan,KeHoach,NVnKH WHERE TaiKhoan.id=NVnKH.id AND KeHoach.idkh=NVnKH.idKH AND KeHoach.noidung=?",search_text)
            kq=query.fetchall()
            self.tbNhiemVu.setColumnHidden(0, True)
            self.tbNhiemVu.setColumnCount(5)
            self.tbNhiemVu.setColumnWidth(0, 30)
            self.tbNhiemVu.setColumnWidth(1, 100)
            self.tbNhiemVu.setColumnWidth(2, 150)
            self.tbNhiemVu.setColumnWidth(3, 150)
            self.tbNhiemVu.setColumnWidth(4, 200)
            header_labels = ['idkh', 'Tài Khoản', 'Họ Tên','Mã Kế Hoạch', 'Nội Dung']
            self.tbNhiemVu.setHeaderLabels(header_labels)

            for row in kq:
                item = QTreeWidgetItem(self.tbNhiemVu)
                for i, value in enumerate(row):
                    item.setText(i, str(value))
            self.update_hoten()
        else:
            self.Load_data3()
        
        self.txtTim.clear()

    def Load_data_makh2(self):
        selected_idkh = self.cbMaKH_2.currentText()

        self.tbNhiemVu.clear()

        query=db.cursor()
        query.execute("SELECT NVnKH.idNVKH,TaiKhoan.tk,TaiKhoan.hoten,KeHoach.idkh,KeHoach.noidung FROM TaiKhoan,KeHoach,NVnKH WHERE TaiKhoan.id=NVnKH.id AND KeHoach.idkh=NVnKH.idKH AND KeHoach.idkh=? ",selected_idkh)
        kq=query.fetchall()
        self.tbNhiemVu.setColumnHidden(0, True)
        self.tbNhiemVu.setColumnCount(5)
        self.tbNhiemVu.setColumnWidth(0, 30)
        self.tbNhiemVu.setColumnWidth(1, 100)
        self.tbNhiemVu.setColumnWidth(2, 150)
        self.tbNhiemVu.setColumnWidth(3, 150)
        self.tbNhiemVu.setColumnWidth(4, 200)
        header_labels = ['idkh', 'Tài Khoản', 'Họ Tên','Mã Kế Hoạch', 'Nội Dung']
        self.tbNhiemVu.setHeaderLabels(header_labels)

        for row in kq:
            item = QTreeWidgetItem(self.tbNhiemVu)
            for i, value in enumerate(row):
                item.setText(i, str(value))
        self.update_hoten()

    def load_cbbox_makh_2(self):
        query=db.cursor()
        query.execute("SELECT idkh FROM KeHoach")
        kq=query.fetchall()
        for row in kq:
            self.cbMaKH_2.addItem(str(row.idkh))

    def on_combobox_changed_KH(self):
        selected_text = self.cbBox_2.currentText()

        if selected_text == "Tat Ca":
            self.Load_data2()
        elif selected_text == "Hoan Thanh":
            self.tbKeHoach.clear()

            query=db.cursor()
            query.execute("select * from KeHoach WHERE trangthai=?",selected_text)
            kq=query.fetchall()
            self.tbKeHoach.setColumnHidden(0, True)
            self.tbKeHoach.setColumnCount(5)
            self.tbKeHoach.setColumnWidth(0, 30)
            self.tbKeHoach.setColumnWidth(1, 90) #10
            self.tbKeHoach.setColumnWidth(2, 130) #20
            self.tbKeHoach.setColumnWidth(3, 150) #50
            self.tbKeHoach.setColumnWidth(4, 130) #20
            self.tbKeHoach.setColumnWidth(5, 100)
            header_labels = ['Stt', 'Ngày giao', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
            self.tbKeHoach.setHeaderLabels(header_labels)

            for row in kq:
                item = QTreeWidgetItem(self.tbKeHoach)
                for i, value in enumerate(row):
                    item.setText(i, str(value))
        elif selected_text == "Chua Hoan Thanh":
            self.tbKeHoach.clear()

            query=db.cursor()
            query.execute("select * from KeHoach WHERE trangthai=?",selected_text)
            kq=query.fetchall()
            self.tbKeHoach.setColumnHidden(0, True)
            self.tbKeHoach.setColumnCount(5)
            self.tbKeHoach.setColumnWidth(0, 30)
            self.tbKeHoach.setColumnWidth(1, 90) #10
            self.tbKeHoach.setColumnWidth(2, 130) #20
            self.tbKeHoach.setColumnWidth(3, 150) #50
            self.tbKeHoach.setColumnWidth(4, 130) #20
            self.tbKeHoach.setColumnWidth(5, 100)
            header_labels = ['Stt', 'Ngày giao', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
            self.tbKeHoach.setHeaderLabels(header_labels)

            for row in kq:
                item = QTreeWidgetItem(self.tbKeHoach)
                for i, value in enumerate(row):
                    item.setText(i, str(value))

    def Sua_data2(self):
        selected_items = self.tbNhiemVu.selectedItems()
        if selected_items:
            selected_item = selected_items[0].text(0)
            nd=self.txtND.text()
            ten=self.txtTenNV.text()
            selected_tk = self.cbTK.currentText()
            selected_idkh = self.cbMaKH.currentText()
            if not nd:
                QMessageBox.warning(self, 'Lỗi', 'Vui lòng chọn lại Mã Kế Hoạch')
                return
            if not ten:
                QMessageBox.warning(self, 'Lỗi', 'Vui lòng chọn lại Tài Khoản')
                return
            # query=db.cursor()
            # query.execute("UPDATE NVnKH SET idKH=?,id=(SELECT id FROM TaiKhoan WHERE tk = ?) WHERE idNVKH=?", selected_idkh,selected_tk,selected_item)
            # db.commit()
            query_admin=db.cursor()
            query_admin.execute("select * from NVnKH where idKH=? and id=(SELECT id FROM TaiKhoan WHERE tk = ?)",selected_idkh,selected_tk)
            kt=query_admin.fetchone()
            if kt:
                QMessageBox.information(self,"Thông báo","Nhân viên "+ten+" đã được giao kế hoạch "+nd+" rồi!")
            else:
                query=db.cursor()
                query.execute("UPDATE NVnKH SET idKH=?,id=(SELECT id FROM TaiKhoan WHERE tk = ?) WHERE idNVKH=?", selected_idkh,selected_tk,selected_item)
                db.commit()

            self.Load_data3()
            self.Load_data()
        else:
            QMessageBox.information(self,'Thông Báo','Bạn chưa chọn dòng muốn sửa')

    def Xoa_data2(self):
        selected_items = self.tbNhiemVu.selectedItems()
        if selected_items: 
            selected_item = selected_items[0].text(0)

            query=db.cursor()
            query.execute("DELETE FROM NVnKH WHERE idNVKH=?", selected_item)
            db.commit()

            self.Load_data3()
            self.Load_data()
        else:
            QMessageBox.information(self,'Thông Báo','Bạn chưa chọn dòng muốn xóa')

    def Them_data2(self):
        nd=self.txtND.text()
        ten=self.txtTenNV.text()
        a=self.txtte.text()
        selected_tk = self.cbTK.currentText()
        selected_idkh = self.cbMaKH.currentText()
        if not nd:
            QMessageBox.warning(self, 'Lỗi', 'Vui lòng chọn lại Mã Kế Hoạch')
            return
        if not ten:
            QMessageBox.warning(self, 'Lỗi', 'Vui lòng chọn lại Tài Khoản')
            return
        query_admin=db.cursor()
        query_admin.execute("select * from NVnKH where idKH=? and id=(SELECT id FROM TaiKhoan WHERE tk = ?)",selected_idkh,selected_tk)
        kt=query_admin.fetchone()
        if kt:
            QMessageBox.information(self,"Thông báo","Nhân viên "+ten+" đã được giao kế hoạch "+nd+" rồi!")
        else:
            query=db.cursor()
            query.execute("INSERT INTO NVnKH (idKH,id,tkgiao) VALUES (?,(SELECT id FROM TaiKhoan WHERE tk = ?),?)",selected_idkh,selected_tk,a)
            db.commit()
            
        self.Load_data3()
        self.Load_data2()
        self.Load_data()
        
    def Load_data3(self):
        self.tbNhiemVu.clear()
        a=self.txtte.text()
        query=db.cursor()
        query.execute("SELECT NVnKH.idNVKH,TaiKhoan.tk,TaiKhoan.hoten,KeHoach.idkh,KeHoach.noidung FROM TaiKhoan,KeHoach,NVnKH WHERE TaiKhoan.id=NVnKH.id AND KeHoach.idkh=NVnKH.idKH AND NVnKH.tkgiao=?",a)
        kq=query.fetchall()
        self.tbNhiemVu.setColumnHidden(0, True)
        self.tbNhiemVu.setColumnCount(5)
        self.tbNhiemVu.setColumnWidth(0, 30)
        self.tbNhiemVu.setColumnWidth(1, 100)
        self.tbNhiemVu.setColumnWidth(2, 150)
        self.tbNhiemVu.setColumnWidth(3, 150)
        self.tbNhiemVu.setColumnWidth(4, 200)
        header_labels = ['idkh', 'Tài Khoản', 'Họ Tên','Mã Kế Hoạch', 'Nội Dung']
        self.tbNhiemVu.setHeaderLabels(header_labels)

        for row in kq:
            item = QTreeWidgetItem(self.tbNhiemVu)
            for i, value in enumerate(row):
                item.setText(i, str(value))
        self.update_hoten()

    def add_Item(self):
        self.cbTK.clear()
        self.cbMaKH.clear()
        query=db.cursor()
        query.execute("SELECT tk FROM TaiKhoan")
        kq=query.fetchall()
        for row in kq:
            self.cbTK.addItem(row.tk)
        query=db.cursor()
        query.execute("SELECT idkh FROM KeHoach")
        kq=query.fetchall()
        for row in kq:
            self.cbMaKH.addItem(str(row.idkh))

    def update_hoten(self):
        # Lấy giá trị tk được chọn từ QComboBox
        selected_tk = self.cbTK.currentText()
        selected_idkh = self.cbMaKH.currentText()

        # Thực hiện truy vấn để lấy hoten tương ứng từ cơ sở dữ liệu
        query=db.cursor()
        query.execute("SELECT hoten FROM TaiKhoan WHERE tk=?", (selected_tk,))
        hoten_result = query.fetchone()
        # Hiển thị hoten tương ứng trong self.txthoten
        if hoten_result:
            self.txtTenNV.setText(hoten_result[0])
        else:
            self.txtTenNV.setText("")

        query=db.cursor()
        query.execute("SELECT noidung FROM KeHoach WHERE idkh=?", (selected_idkh,))
        noidung_result = query.fetchone()
        if noidung_result:
            self.txtND.setText(noidung_result[0])
        else:
            self.txtND.setText("")

        query=db.cursor()
        query.execute("SELECT sdt FROM TaiKhoan WHERE tk=?", (selected_tk,))
        sdt_result = query.fetchone()
        if sdt_result:
            self.txtSdt.setText(sdt_result[0])
        else:
            self.txtSdt.setText("")

    def Xoa_data(self):
        selected_items = self.tbKeHoach.selectedItems()
        if selected_items: 
            selected_item = selected_items[0].text(0)

            query=db.cursor()
            query.execute("DELETE FROM NVnKH WHERE idkh=?", selected_item)
            db.commit()
            query.execute("DELETE FROM KeHoach WHERE idkh=?", selected_item)
            db.commit()

            self.Load_data2()
            self.Load_data()
        else:
            QMessageBox.information(self,'Thông Báo','Bạn chưa chọn dòng muốn xóa')
        
        self.add_Item()
        self.Load_data3()

    def Sua_data(self):
        selected_items = self.tbKeHoach.selectedItems()
        if selected_items:
            selected_item = selected_items[0].text(0)
            nd=self.txtNoiDung.text()
            ngay=self.ngay.date().toString("yyyy-MM-dd")
            mota=self.txtMoTa.text()
            trangthai='Chua Hoan Thanh'
            if not nd:
                QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần Nội Dung!')
                return
            if not mota:
                QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần Mô Tả')
                return
            query=db.cursor()
            query.execute("UPDATE KeHoach SET ngay=?, noidung=?, mota=?, trangthai=? WHERE idkh=?", ngay, nd, mota, trangthai, selected_item)
            db.commit()

            self.Load_data2()
            self.Load_data3()
            self.Load_data()
            self.txtNoiDung.clear()
            self.txtMoTa.clear()
        else:
            QMessageBox.information(self,'Thông Báo','Bạn chưa chọn dòng muốn sửa')

    def on_item_selection_changed2(self):
        selected_items = self.tbKeHoach.selectedItems()
        if selected_items:
            self.txtNoiDung.setText(selected_items[0].text(2))
            self.txtMoTa.setText(selected_items[0].text(3))
            self.ngay.setDate(QDate.fromString(selected_items[0].text(1),"yyyy-MM-dd"))           

    def Them_data(self):
        nd=self.txtNoiDung.text()
        ngay=self.ngay.date().toString("yyyy-MM-dd")
        mota=self.txtMoTa.text()
        trangthai='Chua Hoan Thanh'
        if not nd:
            QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần Nội Dung!')
            return
        if not mota:
            QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần Mô Tả')
            return
        query=db.cursor()
        query.execute(f"INSERT INTO KeHoach (ngay, noidung, mota, trangthai) VALUES ( '{ngay}', '{nd}', '{mota}', '{trangthai}')")
        db.commit()
        a=self.txtte.text() #Lấy ra tên tài khoản
        query.execute("INSERT INTO NVnKH (idkh, id) VALUES ((SELECT MAX(idkh) FROM KeHoach),(SELECT id FROM TaiKhoan WHERE tk = ?))",a)
        db.commit()
        self.Load_data2()
        self.Load_data()
        self.txtNoiDung.clear()
        self.txtMoTa.clear()
        self.add_Item()

    def Load_data2(self):
        self.tbKeHoach.clear()

        a=self.txtte.text()
        query=db.cursor()
        query.execute("SELECT KeHoach.idkh , KeHoach.ngay, KeHoach.noidung, KeHoach.mota, KeHoach.trangthai,KeHoach.ngayht FROM KeHoach JOIN NVnKH ON KeHoach.idkh = NVnKH.idkh WHERE NVnKH.id IN (SELECT id FROM TaiKhoan WHERE tk =?)",a)
        kq=query.fetchall()
        self.tbKeHoach.setColumnCount(5)
        self.tbKeHoach.setColumnHidden(0, True)
        self.tbKeHoach.setColumnWidth(0, 30)
        self.tbKeHoach.setColumnWidth(1, 90)
        self.tbKeHoach.setColumnWidth(2, 130)
        self.tbKeHoach.setColumnWidth(3, 150)
        self.tbKeHoach.setColumnWidth(4, 130)
        self.tbKeHoach.setColumnWidth(5, 100)
        header_labels = [ 'idKH','Ngày giao', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
        self.tbKeHoach.setHeaderLabels(header_labels)

        for row in kq:
            item = QTreeWidgetItem(self.tbKeHoach)
            for i, value in enumerate(row):
                item.setText(i, str(value))

    def HideWgKH(self):
        self.wgKeHoach.setVisible(True)
        self.tbKeHoach.setVisible(True)
        self.wgNhiemVu.setVisible(False)
        self.tbNhiemVu.setVisible(False)
        self.cbBox_2.setVisible(True)
        self.cbMaKH_2.setVisible(False)
        self.txtNDTim.setVisible(False)
        self.btTimND.setVisible(False)
        self.label_12.setText("BẢNG KẾ HOẠCH ĐÃ GIAO")

    def HideWgNV(self):
        self.wgKeHoach.setVisible(False)
        self.tbKeHoach.setVisible(False)
        self.wgNhiemVu.setVisible(True)
        self.tbNhiemVu.setVisible(True)
        self.cbBox_2.setVisible(False)
        self.cbMaKH_2.setVisible(True)
        self.txtNDTim.setVisible(True)
        self.btTimND.setVisible(True)
        self.label_12.setText("BẢNG NHIỆM VỤ ĐÃ GIAO")

    def Thoat(self):
        widget.setFixedSize(310, 180)
        widget.setCurrentIndex(0)

    def DangKy(self):
        widget.setFixedSize(421, 386)
        widget.setCurrentIndex(1)

    def Logout(self):
        self.tabWidget.tabBar().setTabText(2, "Đăng Nhập")
        self.tabWidget.setTabEnabled(2, True)
        widget.setFixedSize(425,320)
        self.tabWidget.setCurrentIndex(2)
        self.tabWidget.tabBar().setTabText(0, "")
        self.tabWidget.tabBar().setTabText(1, "")
        self.tabWidget.setTabEnabled(0, False)
        self.tabWidget.setTabEnabled(1, False)

    def DN(self):
        un=self.txtTk.text()
        self.txtte.setText(un)
        print(self.txtte.text())
        psw=self.txtMk.text()
        if not un:
            QMessageBox.information(self, 'Thông báo', 'Vui lòng nhập Tài Khoản')
            return
        if not psw:
            QMessageBox.information(self, 'Thông báo', 'Vui lòng nhập Mật khẩu')
            return
        query_admin=db.cursor()
        query_admin.execute("select * from TaiKhoan where tk='"+un+"' and mk='"+psw+"' and role = 'admin' ")
        kt_admin=query_admin.fetchone()
        if kt_admin:
            QMessageBox.information(self,"Login Output","Đăng nhập thành công với quyền quản trị")
            self.tabWidget.tabBar().setTabText(0, "Thông Tin Kế Hoạch")
            self.tabWidget.tabBar().setTabText(1, "Quản Lý")
            self.tabWidget.tabBar().setTabText(2, "")
            self.tabWidget.setTabEnabled(0, True)
            self.tabWidget.setTabEnabled(1, True)
            self.tabWidget.setTabEnabled(2, False)
            widget.setFixedSize(691, 631)
            self.tabWidget.setCurrentIndex(1)
        query=db.cursor() 
        query.execute("select * from TaiKhoan where tk='"+un+"' and mk='"+psw+"' and role = 'customer' ")          
        kt=query.fetchone()       
        if kt:
            QMessageBox.information(self,"Login Output","Đăng nhập thành công")
            self.tabWidget.tabBar().setTabText(0, "Thông Tin Kế Hoạch")
            self.tabWidget.tabBar().setTabText(1, "")
            self.tabWidget.tabBar().setTabText(2, "")
            self.tabWidget.setTabEnabled(0, True)
            self.tabWidget.setTabEnabled(1, False)
            self.tabWidget.setTabEnabled(2, False)
            widget.setFixedSize(691, 631)
            self.tabWidget.setCurrentIndex(0)
        if not kt_admin and not kt:
            QMessageBox.information(self,"Thông báo","Đăng nhập thất bại")
        self.Load_data()
        self.Load_data2()
        self.Load_data3()
        self.cbTK.clear()
        self.cbMaKH.clear()
        self.add_Item()

    def Load_data(self):
        self.tbNhom.clear()

        a=self.txtte.text()
        query=db.cursor()
        query.execute("SELECT KeHoach.idkh , KeHoach.ngay, KeHoach.noidung, KeHoach.mota, KeHoach.trangthai,KeHoach.ngayht FROM KeHoach JOIN NVnKH ON KeHoach.idkh = NVnKH.idkh WHERE NVnKH.id IN (SELECT id FROM TaiKhoan WHERE tk =?)",a)
        kq=query.fetchall()
        self.tbNhom.setColumnCount(5)
        self.tbNhom.setColumnHidden(0, True)
        self.tbNhom.setColumnWidth(0, 30)
        self.tbNhom.setColumnWidth(1, 90)
        self.tbNhom.setColumnWidth(2, 130)
        self.tbNhom.setColumnWidth(3, 150)
        self.tbNhom.setColumnWidth(4, 130)
        self.tbNhom.setColumnWidth(5, 100)
        header_labels = [ 'idKH','Ngày giao', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
        self.tbNhom.setHeaderLabels(header_labels)

        for row in kq:
            item = QTreeWidgetItem(self.tbNhom)
            for i, value in enumerate(row):
                item.setText(i, str(value))

    def HoanT(self):
        selected_items = self.tbNhom.selectedItems()
        if selected_items:
            selected_item = selected_items[0].text(0)
            ngay=self.ngay_2.date().toString("yyyy-MM-dd")
            time=None
            if selected_items[0].text(4)=="Hoan Thanh":
                query=db.cursor()
                query.execute("UPDATE KeHoach SET trangthai='Chua Hoan Thanh',ngayht=? WHERE idkh=?",time, selected_item)
                db.commit()
            if selected_items[0].text(4)=="Chua Hoan Thanh":
                query=db.cursor()
                query.execute("UPDATE KeHoach SET trangthai='Hoan Thanh',ngayht=? WHERE idkh=?",ngay, selected_item)
                db.commit()
            self.Load_data()
        else:
            QMessageBox.information(self,'Thông Báo','Bạn chưa chọn dòng')
        
        self.Load_data2()

    def on_item_selection_changed(self):
        selected_items = self.tbNhom.selectedItems()
        if selected_items:
            if selected_items[0].text(4)=="Hoan Thanh":
                self.btHT.setText("Làm Lại")
            else:
                self.btHT.setText("Hoàn thành")

    def on_combobox_changed(self):
        selected_text = self.cbBox.currentText()

        if selected_text == "Tat Ca":
            self.Load_data()
        elif selected_text == "Hoan Thanh":
            self.tbNhom.clear()

            a=self.txtte.text()
            query=db.cursor()
            query.execute("SELECT KeHoach.idkh , KeHoach.ngay, KeHoach.noidung, KeHoach.mota, KeHoach.trangthai,KeHoach.ngayht FROM KeHoach JOIN NVnKH ON KeHoach.idkh = NVnKH.idkh WHERE NVnKH.id IN (SELECT id FROM TaiKhoan WHERE tk =?) AND KeHoach.trangthai='Hoan Thanh'",a)
            kq=query.fetchall()
            self.tbNhom.setColumnCount(5)
            self.tbNhom.setColumnHidden(0, True)
            self.tbNhom.setColumnWidth(0, 30)
            self.tbNhom.setColumnWidth(1, 90)
            self.tbNhom.setColumnWidth(2, 130)
            self.tbNhom.setColumnWidth(3, 150)
            self.tbNhom.setColumnWidth(4, 130)
            self.tbNhom.setColumnWidth(5, 100)
            header_labels = [ 'idKH','Ngày', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
            self.tbNhom.setHeaderLabels(header_labels)

            for row in kq:
                item = QTreeWidgetItem(self.tbNhom)
                for i, value in enumerate(row):
                    item.setText(i, str(value))
        elif selected_text == "Chua Hoan Thanh":
            self.tbNhom.clear()

            a=self.txtte.text()
            query=db.cursor()
            query.execute("SELECT KeHoach.idkh , KeHoach.ngay, KeHoach.noidung, KeHoach.mota, KeHoach.trangthai,KeHoach.ngayht FROM KeHoach JOIN NVnKH ON KeHoach.idkh = NVnKH.idkh WHERE NVnKH.id IN (SELECT id FROM TaiKhoan WHERE tk =?) AND KeHoach.trangthai='Chua Hoan Thanh'",a)
            kq=query.fetchall()
            self.tbNhom.setColumnCount(5)
            self.tbNhom.setColumnHidden(0, True)
            self.tbNhom.setColumnWidth(0, 30)
            self.tbNhom.setColumnWidth(1, 90)
            self.tbNhom.setColumnWidth(2, 130)
            self.tbNhom.setColumnWidth(3, 150)
            self.tbNhom.setColumnWidth(4, 130)
            self.tbNhom.setColumnWidth(5, 100)
            header_labels = [ 'idKH','Ngày', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
            self.tbNhom.setHeaderLabels(header_labels)

            for row in kq:
                item = QTreeWidgetItem(self.tbNhom)
                for i, value in enumerate(row):
                    item.setText(i, str(value))

    def Tim_data(self):
        search_text = self.txtTim.text()
        if search_text:
            self.tbNhom.clear()

            a=self.txtte.text()
            query=db.cursor()
            query.execute("SELECT KeHoach.idkh , KeHoach.ngay, KeHoach.noidung, KeHoach.mota, KeHoach.trangthai, KeHoach.ngayht FROM KeHoach JOIN NVnKH ON KeHoach.idkh = NVnKH.idkh WHERE NVnKH.id IN (SELECT id FROM TaiKhoan WHERE tk =?) AND KeHoach.noidung=?",a,search_text)
            kq=query.fetchall()
            self.tbNhom.setColumnCount(5)
            self.tbNhom.setColumnHidden(0, True)
            self.tbNhom.setColumnWidth(0, 30)
            self.tbNhom.setColumnWidth(1, 90)
            self.tbNhom.setColumnWidth(2, 130)
            self.tbNhom.setColumnWidth(3, 150)
            self.tbNhom.setColumnWidth(4, 130)
            self.tbNhom.setColumnWidth(5, 100)
            header_labels = [ 'idKH','Ngày', 'Nội dung', 'Loại', 'Trạng thái', 'Ngày HT']
            self.tbNhom.setHeaderLabels(header_labels)

            for row in kq:
                item = QTreeWidgetItem(self.tbNhom)
                for i, value in enumerate(row):
                    item.setText(i, str(value))
        else:
            self.Load_data()
        
        self.txtTim.clear()

    def export_to_excel(self):
        # Tạo một Workbook
        workbook = Workbook()
        # Tạo một WorkSheet
        worksheet = workbook.active
        # Đặt tiêu đề cột
        columns = ["Ngày", "Nội dung", "Loại","Trạng thái", "Ngày HT"]
        worksheet.append(columns)
        # Lặp qua dữ liệu QTreeWidget và thêm vào WorkSheet
        for row in range(self.tbNhom.topLevelItemCount()):
            item = self.tbNhom.topLevelItem(row)
            row_data = [item.text(col) for col in range(1,item.columnCount())]
            worksheet.append(row_data)
        file=self.txtEx.text()
        if not file:
            QMessageBox.warning(self, 'Lỗi', 'Vui lòng nhập phần tên cho File Excel')
            return
        # Lưu Workbook thành tệp Excel
        workbook.save(file+".xlsx")
        QMessageBox.information(self,'Thông Báo','Đã lưu ra File Excel')


app=QApplication(sys.argv)
widget=QtWidgets.QStackedWidget()
#Khai báo các widgets
MoDau=OpenW()
DangKy_NHOM=DKApp()
Admin=AdminApp()
CaNhan=CaNhanW()
#Thêm các widget 
widget.addWidget(MoDau) #0
widget.addWidget(DangKy_NHOM) #1
widget.addWidget(Admin) #2
widget.addWidget(CaNhan) #3
#Hiển thị widget
widget.setFixedSize(310, 180)
widget.setCurrentIndex(0)
widget.show()
app.exec()