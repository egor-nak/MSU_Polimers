import pandas as pd
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QLabel, QVBoxLayout
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






def sum_massives(x, y):
    tmp = []
    for i in range(len(x)):
        tmp.append(x[i] + y[i])
    return tmp

def data_preparation(df, k_press=3):

    df_head_stuff = df.iloc[:3]
    df = df.iloc[3:]
    pos = df[1].tolist()
    force = df[2].tolist()
    time = df[3].tolist()
    exten = df[4].tolist()
    strain = df[5].tolist()
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
    y_smooth = savgol_filter(y, 50, 3)
    y_smooth = y_smooth.tolist()

    local_maxis = []
    for i in range(len(y_smooth)):
        tmp = []
        for j in y_smooth[max(0, i - int(len(y_smooth) * 0.1)): min(len(y_smooth), i + int(len(y_smooth) * 0.1))]:
            tmp.append(j)
        if max(tmp) == y_smooth[i]:
            local_maxis.append([y_smooth[i], x[i]])

    local_maxis.sort()
    local_maxis = local_maxis[::-1]


    local_mins = []
    for i in range(len(y_smooth)):
        tmp = []
        for j in y_smooth[max(0, i - int(len(y_smooth) * 0.05)): min(len(y_smooth), i + int(len(y_smooth) * 0.05))]:
            tmp.append(j)
        if min(tmp) == y_smooth[i]:
            local_mins.append([y_smooth[i], x[i]])

    local_mins.sort()
    local_mins = local_mins[::-1]

    maxi_points = [[i[-1], i[0]] for i in local_maxis]
    mini_points = [[i[-1], i[0]] for i in local_mins]
    maxi_points.sort()
    mini_points.sort()

    del mini_points[-1]
    del mini_points[0]


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
    final_data = np.array(final_data)
    final_data = pd.DataFrame(final_data)
    # print(type(final_data), type(df_head_stuff))
    final_data = pd.concat([df_head_stuff, final_data])
    return [final_data, tmp_X, tmp_Y, maxi_points, mini_points]


def create_plot(x, y, maxi_points, mini_points, name, path_to_write):
    plt.clf()
    plt.title(name)
    plt.xlabel("Time, s")
    plt.ylabel("Stress, N/mm^2")
    plt.plot(x, y)
    for i in maxi_points:
        plt.plot(i[0], i[1], 'o', color="red")
        # plt.annotate("(" + str(i[0]) + ";" + str(i[1]) + ")", (i[0], i[1]))
    for i in mini_points:
        plt.plot(i[0], i[1], 'o', color="green")
        # plt.annotate("(" + str(i[0]) + ";" + str(i[1]) + ")", (i[0], i[1]))
    tmp_path = os.path.join(path_to_write, name+'.png')
    plt.savefig(tmp_path)
    # plt.show()


def create_plots(graphiks_stuff, path_to_write):
    xs = []
    ys = []
    names = []
    for i in graphiks_stuff.keys():
        create_plot(graphiks_stuff[i][0], graphiks_stuff[i][1], graphiks_stuff[i][2], graphiks_stuff[i][3], i, path_to_write)
        xs.append(graphiks_stuff[i][0])
        ys.append(graphiks_stuff[i][1])
        names.append(i)
    plt.clf()
    colors = ["blue", "orange", "purple", "brown", "pink", "olive", "cyan"]
    plt.title("All values")
    plt.xlabel("Time, s")
    plt.ylabel("Stress, N/mm^2")
    for i in range(len(xs)):
        plt.plot(xs[i], ys[i], label=names[i])
    leg = plt.legend(loc='upper right')
    tmp_path = os.path.join(path_to_write, 'all_in_one.png')
    plt.savefig(tmp_path)


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

def create_new_table(path, path_to_write):
    os.mkdir(path_to_write)
    tmp = xl.load_workbook(path)
    sheets_names = tmp.sheetnames
    new_sheets = {}
    graphiks_stuff = {}
    for name in sheets_names:
        if name == "обработка":
            df = pd.read_excel(path, sheet_name=[name], header=None)[name]
            new_sheets["обработка"] = df
            continue
        df = pd.read_excel(path, sheet_name=[name], header=None)[name]
        tmp = data_preparation(df)
        new_sheets[name] = tmp[0]
        graphiks_stuff[name] = tmp[1:]

    create_plots(graphiks_stuff, path_to_write)
    path_tmp = os.path.join(path_to_write, 'output.xlsx')
    with pd.ExcelWriter(path_tmp) as writer:
        for i in new_sheets.keys():
            new_sheets[i].to_excel(writer, sheet_name=i)
    path_tmp = os.path.join(path_to_write, 'areas.txt')
    with open(path_tmp, 'w') as f:
        for i in graphiks_stuff.keys():
            f.write(str(i) + ":"+"\n")
            f.write("Первый максимум [Time, Stress]: "+ str(graphiks_stuff[i][2][0]) + "\n")
            f.write("Последний максимум [Time, Stress]: " + str(graphiks_stuff[i][2][-1]) + "\n")
            f.write("Начало полки [Time, Stress]: "+ str(graphiks_stuff[i][3][0]) + "\n")
            f.write("Площадь до первого максимума: "+ str(calc_area_under_curve(graphiks_stuff[i][0], graphiks_stuff[i][1], graphiks_stuff[i][2][0][0])) + "\n")
            f.write("Площадь до последнего максимума: "+  str(calc_area_under_curve(graphiks_stuff[i][0], graphiks_stuff[i][1], graphiks_stuff[i][2][-1][0])) + "\n")
            f.write("Площадь до начала полки: "+ str(calc_area_under_curve(graphiks_stuff[i][0], graphiks_stuff[i][1], graphiks_stuff[i][3][0][0])) + "\n")
            f.write("\n")



class ExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.save_file_t = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel Data Processing')

        self.openButton = QPushButton('Open Excel File', self)
        self.openButton.clicked.connect(self.openFile)

        self.downloadButton = QPushButton('Select new folder name', self)
        self.downloadButton.clicked.connect(self.downloadFile)
        self.downloadButton.setEnabled(False)

        self.processButton = QPushButton('Start Processing', self)
        self.processButton.clicked.connect(self.start_processing)
        self.processButton.setEnabled(False)

        self.label = QLabel('', self)
        self.name_file = ""



        layout = QVBoxLayout()
        layout.addWidget(self.openButton)
        layout.addWidget(self.downloadButton)
        layout.addWidget(self.processButton)
        layout.addWidget(self.label)


        self.setLayout(layout)

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
        if self.save_file_t:
            # print(save_file)
            create_new_table(self.name_file, self.save_file_t)
            # shutil.copyfile(self.new_file, save_file)
            self.label.setText(f'Select table')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelApp()
    ex.show()
    sys.exit(app.exec_())
