# -*- coding: utf-8 -*-
"""
Created on Sat Jan  4 20:12:09 2025

@author: young
"""

import sys
import random
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView, QLabel, QSizePolicy
from PyQt5 import QtGui, QtCore
from PyQt5.QtCore import Qt
from datetime import datetime

class RandomizationApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Group Randomizer")
        self.groups = []

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # Output path
        output_layout = QHBoxLayout()
        self.output_line_edit = QLineEdit(self)
        self.output_button = QPushButton("Select Output Path", self)
        self.output_button.clicked.connect(self.select_output_path)
        output_layout.addWidget(self.output_line_edit)
        output_layout.addWidget(self.output_button)

        # Input fields, Add button, and Delete button
        input_layout = QHBoxLayout()
        self.group_name_edit = QLineEdit(self)
        self.group_name_edit.setPlaceholderText("Group Name")
        self.sample_size_edit = QLineEdit(self)
        self.sample_size_edit.setPlaceholderText("Sample Size (positive integer)")
        self.sample_size_edit.setValidator(QtGui.QRegExpValidator(QtCore.QRegExp("^[1-9][0-9]*$")))
        self.add_button = QPushButton("Add", self)
        self.add_button.clicked.connect(self.add_group)
        self.delete_button = QPushButton("Delete", self)
        self.delete_button.clicked.connect(self.delete_selected_row)
        self.delete_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        input_layout.addWidget(self.group_name_edit)
        input_layout.addWidget(self.sample_size_edit)
        input_layout.addWidget(self.add_button)
        input_layout.addWidget(self.delete_button)

        # Table
        self.table_widget = QTableWidget(self)
        self.table_widget.setColumnCount(2)
        self.table_widget.setHorizontalHeaderLabels(["Group Name", "Sample Size"])
        self.table_widget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Randomize layout
        randomize_layout = QHBoxLayout()
        self.seed_line_edit = QLineEdit(self)
        self.seed_line_edit.setPlaceholderText("Random Seed (optional)")
        self.seed_line_edit.setValidator(QtGui.QRegExpValidator(QtCore.QRegExp("^[1-9][0-9]*$")))
        self.seed_line_edit.setFixedWidth(120)
        self.randomize_button = QPushButton("Randomize", self)
        self.randomize_button.clicked.connect(self.randomize_groups)
        self.randomize_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        randomize_layout.addWidget(self.seed_line_edit)
        randomize_layout.addStretch()
        randomize_layout.addWidget(self.randomize_button)
        randomize_layout.addStretch()

        # Add to main layout
        main_layout.addLayout(output_layout)
        main_layout.addLayout(input_layout)
        main_layout.addWidget(self.table_widget)
        main_layout.addLayout(randomize_layout)
        self.setLayout(main_layout)

    def select_output_path(self):
        options = QFileDialog.Options()
        directory = QFileDialog.getExistingDirectory(self, "Select Output Path", options=options)
        if directory:
            self.output_line_edit.setText(directory)

    def add_group(self):
        group_name = self.group_name_edit.text().strip()
        if any(group[0] == group_name for group in self.groups):
            print(f"Error: Group name '{group_name}' already exists.")
            return
        try:
            sample_size = int(self.sample_size_edit.text().strip())
        except ValueError:
            print("Error: Sample size must be a positive integer.")
            return
        self.groups.append((group_name, sample_size))
        self.update_table()
        self.group_name_edit.clear()
        self.sample_size_edit.clear()

    def update_table(self):
        self.table_widget.setRowCount(len(self.groups))
        for row, (group_name, sample_size) in enumerate(self.groups):
            self.table_widget.setItem(row, 0, QTableWidgetItem(group_name))
            self.table_widget.setItem(row, 1, QTableWidgetItem(str(sample_size)))

    def delete_selected_row(self):
        selected_rows = self.table_widget.selectionModel().selectedRows()
        if selected_rows:
            row = selected_rows[0].row()
            del self.groups[row]
            self.update_table()

    def randomize_groups(self):
        participants = []
        for group_name, sample_size in self.groups:
            participants.extend([(group_name, i + 1) for i in range(sample_size)])

        seed_text = self.seed_line_edit.text().strip()
        random_seed = None
        if seed_text:
            try:
                random_seed = int(seed_text)
                random.seed(random_seed)
            except ValueError:
                print("Error: Invalid random seed. Proceeding without seed.")

        random.shuffle(participants)

        data = [[idx + 1, group] for idx, (group, _) in enumerate(participants)]
        df = pd.DataFrame(data, columns=["Participant ID", "Group Name"])

        timestamp = datetime.now().strftime("%Y_%m_%d %H_%M_%S")
        output_file_name = f"{timestamp} RandomizationAllocation.xlsx"
        output_path = self.output_line_edit.text().strip()
        if not output_path:
            print("Error: Output path is not set.")
            return

        output_file = f"{output_path}/{output_file_name}"

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Randomization")
            seed_df = pd.DataFrame({"Random Seed": [random_seed] if random_seed else ["No seed used"]})
            seed_df.to_excel(writer, index=False, sheet_name="Random Seed")

        print(f"Randomization allocation saved to {output_file}")

def main():
    app = QApplication(sys.argv)
    window = RandomizationApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
