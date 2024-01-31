import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QLabel, QVBoxLayout
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.signal import savgol_filter
import numpy as np
from scipy.integrate import simpson
from numpy import trapz
import openpyxl as xl
import shutil
import os
from time import sleep
from PyQt5.QtCore import QObject, QThread, pyqtSignal
from PyQt5.QtGui     import *
from PyQt5.QtCore    import *
from PyQt5.QtWidgets import *
from PyQt5 import QtGui
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QVBoxLayout, QPushButton, QWidget, QFileDialog
import openpyxl
from datetime import date
import sys
import matplotlib
matplotlib.use('Qt5Agg')

from PyQt5 import QtCore, QtGui, QtWidgets

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg, NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget


# -------------------------------------------------- class to crate dynamic graphics

class MplCanvas(FigureCanvasQTAgg):

    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = None
        # self.axes = self.fig.add_subplot(111)
        super(MplCanvas, self).__init__(self.fig)


class Create_Graph_Window(QtWidgets.QMainWindow):

    def __init__(self,data_to_build, *args, **kwargs):
        super(Create_Graph_Window, self).__init__(*args, **kwargs)

        sc = MplCanvas(self, width=5, height=4, dpi=100)
        sc.axes = sc.fig.add_subplot(111, title=data_to_build[2], xlabel=data_to_build[3], ylabel=data_to_build[4])
        sc.axes.plot(data_to_build[0], data_to_build[1])




        # Create toolbar, passing canvas as first parament, parent (self, the MainWindow) as second.
        toolbar = NavigationToolbar(sc, self)

        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(toolbar)
        layout.addWidget(sc)

        # Create a placeholder widget to hold our toolbar and canvas.
        widget = QtWidgets.QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

        self.show()
# -----------------------------------------------------------------------






def sum_massives(x, y):
    tmp = []
    for i in range(len(x)):
        tmp.append(x[i] + y[i])
    return tmp

def data_preparation(df, flag_recalc, geom_params, k_press=3):

    if flag_recalc:
        square_of_example = geom_params[-1] * geom_params[-2]


    df_head_stuff = df.iloc[:3]
    df = df.iloc[3:]
    pos = df[1].tolist()
    force = df[2].tolist()
    time = df[3].tolist()
    exten = df[4].tolist()
    strain = df[5].tolist()
    if flag_recalc:
        stress = [i / square_of_example for i in force]
    else:
        stress = df[6].tolist()

    new_data = [] # данные в которых мы усреднили все значения где время одинаковое
    sum_pos = pos[0]
    sum_force = force[0]
    sum_exten = exten[0]
    sum_strain = strain[0]
    sum_stress = stress[0]
    cnt = 1
    for i in range(1, len(pos)):
        if time[i] == time[i - 1]:
            cnt += 1
            sum_pos += pos[i]
            sum_force += force[i]
            sum_exten += exten[i]
            sum_strain += strain[i]
            sum_stress += stress[i]
        else:
            tmp = [sum_pos / cnt, sum_force / cnt, time[i - 1], sum_exten / cnt, sum_strain / cnt, sum_stress / cnt]
            cnt = 1
            sum_pos = pos[i]
            sum_force = force[i]
            sum_exten = exten[i]
            sum_strain = strain[i]
            sum_stress = stress[i]
            new_data.append(tmp)
    tmp = [sum_pos / cnt, sum_force / cnt, time[i - 1], sum_exten / cnt, sum_strain / cnt, sum_stress / cnt]
    new_data.append(tmp)


    x = np.array([i[2] for i in new_data])
    y = np.array([i[-1] for i in new_data])
    x_new_strain = np.array([i[4] for i in new_data])
    y_smooth = savgol_filter(y, 50, 3)
    y_smooth = y_smooth.tolist()

    local_maxis = []
    local_maxis_new_strain = []
    for i in range(len(y_smooth)):
        tmp = []
        for j in y_smooth[max(0, i - int(len(y_smooth) * 0.1)): min(len(y_smooth), i + int(len(y_smooth) * 0.1))]:
            tmp.append(j)
        if max(tmp) == y_smooth[i]:
            local_maxis.append([y_smooth[i], x[i]])
            local_maxis_new_strain.append([y_smooth[i], x_new_strain[i]])

    local_maxis.sort()
    local_maxis = local_maxis[::-1]

    local_maxis_new_strain.sort()
    local_maxis_new_strain = local_maxis_new_strain[::-1]


    local_mins = []
    local_mins_new_strain = []
    for i in range(len(y_smooth)):
        tmp = []
        for j in y_smooth[max(0, i - int(len(y_smooth) * 0.05)): min(len(y_smooth), i + int(len(y_smooth) * 0.05))]:
            tmp.append(j)
        if min(tmp) == y_smooth[i]:
            local_mins.append([y_smooth[i], x[i]])
            local_mins_new_strain.append([y_smooth[i], x_new_strain[i]])

    local_mins.sort()
    local_mins = local_mins[::-1]

    local_mins_new_strain.sort()
    local_mins_new_strain = local_mins_new_strain[::-1]

    maxi_points = [[i[-1], i[0]] for i in local_maxis]
    mini_points = [[i[-1], i[0]] for i in local_mins]
    maxi_points.sort()
    mini_points.sort()

    del mini_points[-1]
    del mini_points[0]


    maxi_points_new_strain = [[i[-1], i[0]] for i in local_maxis_new_strain]
    mini_points_new_strain = [[i[-1], i[0]] for i in local_mins_new_strain]
    maxi_points_new_strain.sort()
    mini_points_new_strain.sort()

    del mini_points_new_strain[-1]
    del mini_points_new_strain[0]

    final_data = []
    ind = 0
    while new_data[ind][2] < mini_points[-1][0]:
        final_data.append(new_data[ind])
        ind+=1
    while ind < len(new_data):
        tmp = [0] * len(new_data[ind])
        for i in range(k_press):
            if ind + i >= len(new_data):
                break
            tmp = sum_massives(tmp, new_data[ind + i])
        final_data.append([i / k_press for i in tmp])
        ind += k_press
    for i in range(len(final_data)):
        final_data[i] = [i + 1] + final_data[i]
    tmp_X = np.array([i[3] for i in final_data])
    tmp_Y = np.array([i[-1] for i in final_data])
    tmp_X_new_strain = np.array([i[5] for i in final_data])
    final_data = np.array(final_data)
    final_data = pd.DataFrame(final_data)
    # print(type(final_data), type(df_head_stuff))
    final_data = pd.concat([df_head_stuff, final_data])
    return [[final_data, tmp_X, tmp_Y, maxi_points, mini_points], [tmp_X_new_strain, tmp_Y, maxi_points_new_strain, mini_points_new_strain]]


def create_plot(x, y, maxi_points, mini_points, name, path_to_write, flag_strain, naming_stuff, main_class_prop, ind_of_window):
    data_for_graph_window = [x, y]
    plt.clf()
    if name in naming_stuff.keys():
        plt.title(naming_stuff[name])
        data_for_graph_window.append(naming_stuff[name])
    else:
        plt.title(name)
        data_for_graph_window.append(name)
    if flag_strain:
        plt.xlabel("Strain, %")
        data_for_graph_window.append("Strain, %")
    else:
        plt.xlabel("Time, s")
        data_for_graph_window.append("Time, s")
    plt.ylabel("Stress, N/mm^2")
    data_for_graph_window.append("Stress, N/mm^2")
    plt.plot(x, y)
    for i in maxi_points:
        plt.plot(i[0], i[1], 'o', color="red")
        # plt.annotate("(" + str(i[0]) + ";" + str(i[1]) + ")", (i[0], i[1]))
    for i in mini_points:
        plt.plot(i[0], i[1], 'o', color="green")
        # plt.annotate("(" + str(i[0]) + ";" + str(i[1]) + ")", (i[0], i[1]))
    tmp_path = ""

    if flag_strain:
        tmp_path = os.path.join(path_to_write, name + 'strain.png')
    else:
        tmp_path = os.path.join(path_to_write, name+'.png')
    plt.savefig(tmp_path)
    main_class_prop.additional_windows[ind_of_window] = Create_Graph_Window(data_for_graph_window)
    # plt.show()


def create_plots(graphiks_stuff, path_to_write, flag_strain, naming_stuff, main_class_prop, ind_of_window):
    xs = []
    ys = []
    names = []
    for i in graphiks_stuff.keys():
        create_plot(graphiks_stuff[i][0], graphiks_stuff[i][1], graphiks_stuff[i][2], graphiks_stuff[i][3], i, path_to_write, flag_strain, naming_stuff, main_class_prop, ind_of_window)
        xs.append(graphiks_stuff[i][0])
        ys.append(graphiks_stuff[i][1])
        names.append(i)
        ind_of_window+=1
    plt.clf()
    colors = ["blue", "orange", "purple", "brown", "pink", "olive", "cyan"]
    plt.title("All values")
    if flag_strain:
        plt.xlabel("Strain, %")
    else:
        plt.xlabel("Time, s")
    plt.ylabel("Stress, N/mm^2")
    for i in range(len(xs)):
        plt.plot(xs[i], ys[i], label=names[i])
    leg = plt.legend(loc='upper right')
    tmp_path = ""
    if flag_strain:
        tmp_path = os.path.join(path_to_write, 'all_in_one_strain.png')
    else:
        tmp_path = os.path.join(path_to_write, 'all_in_one.png')
    plt.savefig(tmp_path)
    return ind_of_window


def calc_area_under_curve(coords_x, coords_y, x_value=1e9): # y[координта ] x_value - до какой координты по x интегрировать
    # # return area
    # print(coords_x)
    # print(coords_y)
    # print(x_value)
    y_tmp = []
    x_tmp = []
    for i in range(len(coords_x)):
        if coords_x[i] < x_value:
            y_tmp.append(coords_y[i])
            x_tmp.append(coords_x[i])
    area = simpson(np.array(y_tmp), np.array(x_tmp))
    return area

def create_new_table(main_class_prop, path, path_to_write, additional_data, meta_data_table_values):
    os.mkdir(path_to_write)
    tmp = xl.load_workbook(path)
    sheets_names = tmp.sheetnames
    new_sheets = {}
    graphiks_stuff = {}
    strain_graphiks_stuff_dict = {}
    for name in sheets_names:
        if name == "обработка":
            df = pd.read_excel(path, sheet_name=[name], header=None)[name]
            new_sheets["обработка"] = df
            continue
        df = pd.read_excel(path, sheet_name=[name], header=None)[name]
        if additional_data[-2]:
            tmp, strain_graphiks_stuff = data_preparation(df, additional_data[-2], additional_data[-1])
        else:
            tmp, strain_graphiks_stuff = data_preparation(df, additional_data[-2], additional_data[-1])
        new_sheets[name] = tmp[0]
        graphiks_stuff[name] = tmp[1:]
        strain_graphiks_stuff_dict[name] = strain_graphiks_stuff


    comparison_name_to_other = {}

    if additional_data[0]:
        comparison_name_to_other = {"а2": "30 mm/min", "а1": "60 mm/min", "а3": "6 mm/min", "а4": "3 mm/min", "а5": "0.6 mm/min", "а0": "200 mm/min"}
    elif additional_data[1]:
        comparison_name_to_other = {"а2": "20 mm/min", "а1": "60 mm/min", "а3": "6 mm/min", "а4": "2 mm/min", "а5": "0.6 mm/min", "а0": "200 mm/min"}
    else:
        comparison_name_to_other = {"а2": "a2", "а1": "a1", "а3": "a3", "а4": "a4", "а5": "a5",
         "а0": "a6"}


    ind_of_already_opened_windows = 0
    if additional_data[2]:
        ind_of_already_opened_windows = create_plots(graphiks_stuff, path_to_write, False, comparison_name_to_other, main_class_prop, ind_of_already_opened_windows)
    if additional_data[3]:
        ind_of_already_opened_windows = create_plots(strain_graphiks_stuff_dict, path_to_write, additional_data[3], comparison_name_to_other, main_class_prop, ind_of_already_opened_windows)
    path_tmp = os.path.join(path_to_write, 'output.xlsx')
    with pd.ExcelWriter(path_tmp) as writer:
        for i in new_sheets.keys():
            new_sheets[i].to_excel(writer, sheet_name=i)
    path_tmp = os.path.join(path_to_write, 'data.xlsx')
    list_of_calculed_data = [["Названия"] + [str(i) for i in graphiks_stuff.keys()], ["Первый максимум [Strain, Stress]"], ["Последний максимум [Strain, Stress]"], ["Начало полки [Strain, Stress]"], ["Площадь до первого максимума"], ["Площадь до последнего максимума"], ["Площадь до начала полки"]]
    for i in strain_graphiks_stuff_dict.keys():
        list_of_calculed_data[1].append(str(strain_graphiks_stuff_dict[i][2][0]))
        list_of_calculed_data[2].append(str(strain_graphiks_stuff_dict[i][2][-1]))
        print(strain_graphiks_stuff_dict[i][3])
        list_of_calculed_data[3].append(str(strain_graphiks_stuff_dict[i][3][0]))
    for i in graphiks_stuff.keys():
        list_of_calculed_data[4].append(str(calc_area_under_curve(graphiks_stuff[i][0], graphiks_stuff[i][1], graphiks_stuff[i][2][0][0])))
        list_of_calculed_data[5].append(str(calc_area_under_curve(graphiks_stuff[i][0], graphiks_stuff[i][1], graphiks_stuff[i][2][-1][0])))
        list_of_calculed_data[6].append(str(calc_area_under_curve(graphiks_stuff[i][0], graphiks_stuff[i][1], graphiks_stuff[i][3][0][0])))
    first_page = pd.DataFrame(list_of_calculed_data)
    second_page = pd.DataFrame(meta_data_table_values)

    writer = pd.ExcelWriter(path_tmp, engine='xlsxwriter')

    frames = {'Calc_data': first_page, 'Meta_data': second_page}
    for i in frames.keys():
        frames[i].to_excel(writer, sheet_name=i)
    # for sheet, frame in frames.iteritems():  # .use .items for python 3.X
    #     frame.to_excel(writer, sheet_name=sheet)

    writer._save()






class ExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.save_file_t = None
        self.additional_windows = [None] * 20
        self.initUI()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setPen(Qt.black)
        painter.drawLine(0, 140, 800, 140)

    def initUI(self):

        validarValor = QDoubleValidator(0.00, 999.99, 2)
        validarValor.setNotation(QDoubleValidator.StandardNotation)
        validarValor.setDecimals(2)

        self.setFixedWidth(800)
        self.setFixedHeight(600)
        self.setWindowTitle('Excel Data Processing')

        self.openButton = QPushButton('Open Excel File', self)
        self.openButton.clicked.connect(self.openFile)
        self.openButton.move(20, 10)
        self.openButton.resize(760, 30)

        self.downloadButton = QPushButton('Select new folder name', self)
        self.downloadButton.clicked.connect(self.downloadFile)
        self.downloadButton.setEnabled(False)
        self.downloadButton.move(20, 50)
        self.downloadButton.resize(760, 30)

        self.processButton = QPushButton('Start Processing', self)
        self.processButton.clicked.connect(self.start_processing)
        self.processButton.setEnabled(False)
        self.processButton.move(20, 90)
        self.processButton.resize(760, 30)

        self.label = QLabel('', self)
        self.name_file = ""
        self.label.move(20, 120)

        self.label_settings = QLabel('Settings', self)
        self.label_settings.move(360, 150)
        myFont = QtGui.QFont()
        myFont.setBold(True)
        myFont.setRawMode(True)
        self.label_settings.setFont(myFont)

        self.label_settings_type_of_graphics = QLabel('Types of graphs:', self)
        self.label_settings_type_of_graphics.move(20, 170)

        self.cb_stress_time = QCheckBox("Stress(Time)", self)
        self.cb_stress_time.move(20, 190)
        self.cb_stress_time.show()

        self.cb_stress_strain = QCheckBox("Stress(Strain)", self)
        self.cb_stress_strain.move(20, 210)
        self.cb_stress_strain.show()

        self.label_settings_geom_params = QLabel('Set geometric parameters:', self)
        self.label_settings_geom_params.move(305, 170)

        self.cb_basic_geom_params = QCheckBox("Basic (Stress will not be recalculated)", self)
        self.cb_basic_geom_params.move(305, 190)
        self.cb_basic_geom_params.show()

        self.or_label = QLabel('or', self)
        self.or_label.move(395, 210)

        self.label_length = QLabel('length (mm):', self)
        self.label_length.move(305, 230)


        self.textbox_length = QLineEdit(self)
        self.textbox_length.move(410, 230)
        self.textbox_length.resize(100, 15)


        self.label_width = QLabel('width (mm):', self)
        self.label_width.move(305, 250)

        self.textbox_width = QLineEdit(self)
        self.textbox_width.move(410, 250)
        self.textbox_width.resize(100, 15)


        self.label_thickness = QLabel('thickness (mm):', self)
        self.label_thickness.move(305, 270)

        self.textbox_thickness = QLineEdit(self)
        self.textbox_thickness.move(410, 270)
        self.textbox_thickness.resize(100, 15)

        self.label_settings_speed_params = QLabel('Speed mode:', self)
        self.label_settings_speed_params.move(630, 170)

        self.cb_speed_mode_one = QCheckBox("Mode 1 (a2=30, a4=3)", self)
        self.cb_speed_mode_one.move(630, 190)
        self.cb_speed_mode_one.setChecked(True)

        self.cb_speed_mode_two = QCheckBox("Mode 2 (a2=20, a4=2)", self)
        self.cb_speed_mode_two.move(630, 210)
        self.cb_speed_mode_two.show()

        self.meta_data_table = QTableWidget(self)
        self.meta_data_table.setColumnCount(2)
        self.meta_data_table.setRowCount(10)
        self.meta_data_table.move(20, 320)
        self.meta_data_table.resize(760, 270)
        self.meta_data_table.show()
        self.meta_data_table.setItem(0, 0, QtWidgets.QTableWidgetItem("Дата"))
        self.meta_data_table.setItem(0, 1, QtWidgets.QTableWidgetItem(str(date.today())))
        self.meta_data_table.setItem(1, 0, QtWidgets.QTableWidgetItem("Экспериментатор (Имя Фамилия)"))

        self.label_meta_data = QLabel('Meta Data', self)
        self.label_meta_data.move(360, 300)
        self.label_meta_data.setFont(myFont)


    def openFile(self):
        options = QFileDialog.Options()
        file, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)", options=options)
        if file:
            # create_new_table(file)
            self.label.setText(f'Ready to process table')
            self.name_file = file
            self.downloadButton.setEnabled(True)

    def downloadFile(self):
        options = QFileDialog.Options()
        save_file, _ = QFileDialog.getSaveFileName(self, "Save New Folder", "",
                                                   options=options)
        self.save_file_t = save_file
        if save_file:
            self.processButton.setEnabled(True)

    def start_processing(self):
        self.label.setText(f'Processing...')
        flag_speed_mode_one = False
        flag_speed_mode_two = False
        flag_stress_time = False
        flag_stress_strain = False
        recalc_stress = False
        geom_params = []
        if self.cb_stress_time.isChecked():
            flag_stress_time = True
        if self.cb_stress_strain.isChecked():
            flag_stress_strain = True
        if self.cb_speed_mode_one.isChecked():
            flag_speed_mode_one = True
        if self.cb_speed_mode_two.isChecked():
            flag_speed_mode_two = True

        if not self.cb_basic_geom_params.isChecked() and len(self.textbox_thickness.text()) != 0 and len(self.textbox_width.text()) != 0:
            recalc_stress = True
            geom_params = [self.textbox_length.text(), float(self.textbox_width.text()), float(self.textbox_thickness.text())]



        if self.save_file_t:
            # print(save_file)
            meta_data_table_values = []

            for row in range(10):
                tmp_list_meta_data = []
                for col in range(2):
                    if self.meta_data_table.item(row, col) is None:
                        break
                    tmp_list_meta_data.append(str(self.meta_data_table.item(row, col).text()))
                if len(tmp_list_meta_data) == 2:
                    meta_data_table_values.append(tmp_list_meta_data)

            comparison_name_to_other = {}
            if flag_speed_mode_one:
                comparison_name_to_other = {"а2": "30 mm/min", "а1": "60 mm/min", "а3": "6 mm/min", "а4": "3 mm/min",
                                            "а5": "0.6 mm/min", "а0": "200 mm/min"}
            elif flag_speed_mode_two:
                comparison_name_to_other = {"а2": "20 mm/min", "а1": "60 mm/min", "а3": "6 mm/min", "а4": "2 mm/min",
                                            "а5": "0.6 mm/min", "а0": "200 mm/min"}
            else:
                comparison_name_to_other = {"а2": "a2", "а1": "a1", "а3": "a3", "а4": "a4", "а5": "a5",
                                            "а0": "a6"}
            for i in comparison_name_to_other.keys():
                meta_data_table_values.append([i, comparison_name_to_other[i]])

            add_data = [flag_speed_mode_one, flag_speed_mode_two, flag_stress_time, flag_stress_strain, recalc_stress, geom_params]
            create_new_table(self, self.name_file, self.save_file_t, add_data, meta_data_table_values)
            # shutil.copyfile(self.new_file, save_file)
            self.label.setText(f'Select table')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelApp()
    ex.show()
    sys.exit(app.exec_())
