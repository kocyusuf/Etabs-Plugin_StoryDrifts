import comtypes.client
from PyQt5 import QtWidgets, QtGui, QtCore
import sys
import storyDrift_win as sdw
import numpy as np

class StoryDriftCheck(QtWidgets.QMainWindow):
    def __init__(self):
        super(StoryDriftCheck, self).__init__()
        self.ui = sdw.Ui_Form()
        self.ui.setupUi(self)
        ETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
        self.SapModel = ETABSObject.SapModel

        self.ui.kappa_values.setFixedWidth(120)
        self.ui.kappa_values.addItems(["Betonarme, κ=1", "Çelik, κ=0.5"])
        self.ui.condition_values.addItems(["0.008κ", "0.016κ"])
        #   GET LOAD PATTERNS NAMES
        ret_case = self.SapModel.LoadPatterns.GetNameList()
        self.seismic = []
        for i in ret_case[1]:
            ret_loadType = self.SapModel.LoadPatterns.GetLoadType(i)
            if ret_loadType[0] == 5:
                self.seismic.append(i)

        self.ret_story = self.SapModel.Story.GetStories()
        self.story_heights = [story/1000 for story in self.ret_story[3]]
        self.story_heights.pop(0)

        self.ui.get_values_button.clicked.connect(self.get_parameters)
        self.ui.check_button.clicked.connect(self.check_drifts)



    def get_parameters(self):
        #   EARTHQUAKE PARAMETERS
        self.reduction_factor_x = int(self.ui.reduction_factor_x.text())
        self.reduction_factor_y = int(self.ui.reduction_factor_y.text())
        self.importance_factor = float(self.ui.important_factor.text())

        #   SPECTRUM PERIODS FOR DD2
        self.spectrum_short_forDD2 = float(self.ui.dd2_sds.text())
        self.spectrum_one_forDD2 = float(self.ui.dd2_sd1.text())

        #   SPECTRUM PERIODS FOR DD3
        self.spectrum_short_forDD3 = float(self.ui.dd3_sds.text())
        self.spectrum_one_forDD3 = float(self.ui.dd3_sd1.text())

        #   CORNER PERIODS FOR DD2
        self.corner_period_B_forDD2 = self.spectrum_one_forDD2 / self.spectrum_short_forDD2
        self.corner_period_A_forDD2 = 0.2 * self.corner_period_B_forDD2

        #   CORNER PERIODS FOR DD3
        self.corner_period_B_forDD3 = self.spectrum_one_forDD3 / self.spectrum_short_forDD3
        self.corner_period_A_forDD3 = 0.2 * self.corner_period_B_forDD3

        self.spectrum_long_period = 6.0

        self.check_drifts_x_dir()
        self.check_drifts_y_dir()


    def check_drifts_x_dir(self):
        ret1 = self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        ret_case_1 = self.SapModel.Results.Setup.SetCaseSelectedForOutput(self.seismic[0])

        self.NumberResults = 0
        self.Story = []
        self.LoadCase = []
        self.StepType = []
        self.StepNum = []
        self.Direction = []
        self.Drift = []
        self.Label = []
        self.X = []
        self.Y = []
        self.Z = []

        retX = self.SapModel.Results.StoryDrifts(self.NumberResults, self.Story, self.LoadCase, self.StepType,
                                                 self.StepNum, self.Direction, self.Drift, self.Label,
                                                 self.X, self.Y, self.Z)


        self.delta_X = []
        for x in retX[6]:
            deltax = (self.reduction_factor_x / self.importance_factor) * x
            self.delta_X.append(deltax)
        self.drifts_X = [item for item in self.delta_X]
        self.drifts_X = np.reshape(self.drifts_X, (int(len(self.drifts_X) / 3), 3))
        self.max_drift_X = []
        for drift in self.drifts_X :
            self.max_drift_X.append(max(drift))


        # CALCULATE LAMBDA COEFFICIENT
        ret_period = self.SapModel.Results.ModalParticipatingMassRatios()
        self.period_structure_X = ret_period[4][ret_period[5].index(max(ret_period[5]))]

        if 0 <= self.period_structure_X <= self.corner_period_A_forDD2:
            self.Sae_X_dd2 = (0.4 + 0.6 * (self.period_structure_X / self.corner_period_A_forDD2)) * self.spectrum_short_forDD2
        elif self.corner_period_A_forDD2 <= self.period_structure_X <= self.corner_period_B_forDD2:
            self.Sae_X_dd2 = self.spectrum_short_forDD2
        elif self.corner_period_B_forDD2 <= self.period_structure_X <= self.spectrum_long_period:
            self.Sae_X_dd2 = self.spectrum_one_forDD2 / self.period_structure_X
        else:
            self.Sae_X_dd2 = (self.spectrum_one_forDD2 * self.spectrum_long_period) / (self.period_structure_X ** 2)

        if 0 <= self.period_structure_X <= self.corner_period_A_forDD3:
            self.Sae_X_dd3 = (0.4 + 0.6 * (self.period_structure_X / self.corner_period_A_forDD3)) * self.spectrum_short_forDD3
        elif self.corner_period_A_forDD3 <= self.period_structure_X <= self.corner_period_B_forDD3:
            self.Sae_X_dd3 = self.spectrum_short_forDD3
        elif self.corner_period_B_forDD3 <= self.period_structure_X <= self.spectrum_long_period:
            self.Sae_X_dd3 = self.spectrum_one_forDD3 / self.period_structure_X
        else:
            self.Sae_X_dd3 = (self.spectrum_one_forDD3 * self.spectrum_long_period) / (self.period_structure_X ** 2)

        self.lamda_X = self.Sae_X_dd3 / self.Sae_X_dd2
        self.result_X = []
        for index in range(0, len(self.max_drift_X)):
            result = self.lamda_X * self.max_drift_X[index] / self.story_heights[index]
            self.result_X.append(result)

        print("result-x: ", self.result_X)

    def check_drifts_y_dir(self):
        ret1 = self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        ret_case = self.SapModel.Results.Setup.SetCaseSelectedForOutput(self.seismic[1])

        self.NumberResults = 0
        self.Story = []
        self.LoadCase = []
        self.StepType = []
        self.StepNum = []
        self.Direction = []
        self.Drift = []
        self.Label = []
        self.X = []
        self.Y = []
        self.Z = []

        retY = self.SapModel.Results.StoryDrifts(self.NumberResults, self.Story, self.LoadCase, self.StepType,
                                                 self.StepNum, self.Direction, self.Drift, self.Label,
                                                 self.X, self.Y, self.Z)

        self.delta_Y = []
        for y in retY[6]:
            deltay = (self.reduction_factor_y / self.importance_factor) * y
            self.delta_Y.append(deltay)
        self.drifts_Y = [item for item in self.delta_Y]
        self.drifts_Y = np.reshape(self.drifts_Y, (int(len(self.drifts_Y) / 3), 3))
        self.max_drift_Y = []
        for drift in self.drifts_Y :
            self.max_drift_Y.append(max(drift))


        # CALCULATE LAMBDA COEFFICIENT
        ret_period = self.SapModel.Results.ModalParticipatingMassRatios()
        self.period_structure_Y = ret_period[4][ret_period[6].index(max(ret_period[6]))]

        if 0 <= self.period_structure_Y <= self.corner_period_A_forDD2 :
            self.Sae_Y_dd2 = (0.4 + 0.6 * (
                        self.period_structure_Y / self.corner_period_A_forDD2)) * self.spectrum_short_forDD2
        elif self.corner_period_A_forDD2 <= self.period_structure_Y <= self.corner_period_B_forDD2 :
            self.Sae_Y_dd2 = self.spectrum_short_forDD2
        elif self.corner_period_B_forDD2 <= self.period_structure_Y <= self.spectrum_long_period :
            self.Sae_Y_dd2 = self.spectrum_one_forDD2 / self.period_structure_Y
        else :
            self.Sae_Y_dd2 = (self.spectrum_one_forDD2 * self.spectrum_long_period) / (self.period_structure_Y ** 2)


        if 0 <= self.period_structure_Y <= self.corner_period_A_forDD3 :
            self.Sae_Y_dd3 = (0.4 + 0.6 * (
                        self.period_structure_Y / self.corner_period_A_forDD3)) * self.spectrum_short_forDD3
        elif self.corner_period_A_forDD3 <= self.period_structure_Y <= self.corner_period_B_forDD3 :
            self.Sae_Y_dd3 = self.spectrum_short_forDD3
        elif self.corner_period_B_forDD3 <= self.period_structure_Y <= self.spectrum_long_period :
            self.Sae_Y_dd3 = self.spectrum_one_forDD3 / self.period_structure_Y
        else :
            self.Sae_Y_dd3 = (self.spectrum_one_forDD3 * self.spectrum_long_period) / (self.period_structure_Y ** 2)

        self.lamda_Y = self.Sae_Y_dd3 / self.Sae_Y_dd2
        self.result_Y = []
        for index in range(0, len(self.max_drift_Y)):
            result = self.lamda_Y * self.max_drift_Y[index] / self.story_heights[index]
            self.result_Y.append(result)

        print("result-y: ", self.result_Y)



    def check_drifts(self):
        if self.ui.condition_values.currentText() == "0.008κ":
            if self.ui.kappa_values.currentText() == "Betonarme, κ=1":
                condition = 0.008 * 1
                for item in range(0, len(self.result_X)):
                    if self.result_X[item] <= condition and self.result_Y[item] <= condition:
                        message = QtWidgets.QMessageBox()
                        message.setWindowTitle("Göreli Kat Ötelenmesi Kontrolü")
                        message.setText("Göreli kat ötelenmesi kontrolü sağlanmıştır.")
                        x = message.exec_()
                        break
                    else:
                        message = QtWidgets.QMessageBox()
                        message.setWindowTitle("Göreli Kat Ötelenmesi Kontrolü")
                        message.setText("Göreli kat ötelenmesi kontrolü sağlanamamıştır.")
                        x = message.exec_()
                        break

            if self.ui.kappa_values.currentText() == "Çelik, κ=0.5":
                condition = 0.008 * 0.5
                for item in range(0, len(self.result_X)):
                    if self.result_X[item] <= condition and self.result_Y[item] <= condition:
                        message = QtWidgets.QMessageBox()
                        message.setWindowTitle("Göreli Kat Ötelenmesi Kontrolü")
                        message.setText("Göreli kat ötelenmesi kontrolü sağlanmıştır.")
                        x = message.exec_()
                        break
                    else:
                        message = QtWidgets.QMessageBox()
                        message.setWindowTitle("Göreli Kat Ötelenmesi Kontrolü")
                        message.setText("Göreli kat ötelenmesi kontrolü sağlanamamıştır.")
                        x = message.exec_()
                        break

        if self.ui.condition_values.currentText() == "0.016κ" :
            if self.ui.kappa_values.currentText() == "Betonarme, κ=1" :
                condition = 0.016 * 1
                for item in range(0, len(self.result_X)) :
                    if self.result_X[item] <= condition and self.result_Y[item] <= condition :
                        message = QtWidgets.QMessageBox()
                        message.setWindowTitle("Göreli Kat Ötelenmesi Kontrolü")
                        message.setText("Göreli kat ötelenmesi kontrolü sağlanmıştır.")
                        x = message.exec_()
                        break
                    else :
                        message = QtWidgets.QMessageBox()
                        message.setWindowTitle("Göreli Kat Ötelenmesi Kontrolü")
                        message.setText("Göreli kat ötelenmesi kontrolü sağlanamamıştır.")
                        x = message.exec_()
                        break

            if self.ui.kappa_values.currentText() == "Çelik, κ=0.5" :
                condition = 0.016 * 0.5
                for item in range(0, len(self.result_X)) :
                    if self.result_X[item] <= condition and self.result_Y[item] <= condition :
                        message = QtWidgets.QMessageBox()
                        message.setWindowTitle("Göreli Kat Ötelenmesi Kontrolü")
                        message.setText("Göreli kat ötelenmesi kontrolü sağlanmıştır.")
                        x = message.exec_()
                        break
                    else :
                        message = QtWidgets.QMessageBox()
                        message.setWindowTitle("Göreli Kat Ötelenmesi Kontrolü")
                        message.setText("Göreli kat ötelenmesi kontrolü sağlanamamıştır.")
                        x = message.exec_()
                        break


def app():
    app = QtWidgets.QApplication(sys.argv)
    window = StoryDriftCheck()
    window.show()
    sys.exit(app.exec_())

app()