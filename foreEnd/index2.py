# Copyright (C) 2022 The Qt Company Ltd.
# SPDX-License-Identifier: LicenseRef-Qt-Commercial OR BSD-3-Clause

"""PySide6 Active Qt Viewer example"""

import sys
from lib.share import SI

from PyQt5.QtWidgets import qApp
from PySide2.QtAxContainer import QAxSelect, QAxWidget
from PySide2.QtWidgets import (QApplication, QDialog,QAction,
    QMainWindow, QMessageBox, QToolBar)


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        toolBar = QToolBar()
        self.addToolBar(toolBar)
        fileMenu = self.menuBar().addMenu("&File")
        loadAction = QAction("Load...", self, shortcut="Ctrl+L", triggered=self.load)
        fileMenu.addAction(loadAction)
        toolBar.addAction(loadAction)
        exitAction = QAction("E&xit", self, shortcut="Ctrl+Q", triggered=self.close)
        fileMenu.addAction(exitAction)

        aboutMenu = self.menuBar().addMenu("&About")
        aboutQtAct = QAction("About &Qt", self, triggered=qApp.aboutQt)
        aboutMenu.addAction(aboutQtAct)
        self.axWidget = QAxWidget()
        self.setCentralWidget(self.axWidget)

    def load(self):
        axSelect = QAxSelect(self)
        if axSelect.exec() == QDialog.Accepted:
            clsid = axSelect.clsid()
            if not self.axWidget.setControl(clsid):
                QMessageBox.warning(self, "AxViewer", f"Unable to load {clsid}.")


def main():
    SI.mainWin2 = MainWindow()
    availableGeometry = SI.mainWin2.screen().availableGeometry()
    SI.mainWin2.resize(availableGeometry.width() / 3, availableGeometry.height() / 2)
    SI.mainWin2.show()