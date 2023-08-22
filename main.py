import csv
import sys

import xlwt
from PyQt5.QtWidgets import QApplication, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout, QPushButton, QGraphicsView, \
    QGraphicsScene, QFormLayout, QLineEdit, QTextEdit, QTableWidget, QTableWidgetItem, QGraphicsPixmapItem, QMessageBox, \
    QLabel, QSpinBox, QGroupBox, QFrame, QHeaderView, QGridLayout, QFileDialog
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QScrollArea
import numpy as np
from matplotlib.animation import FuncAnimation
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from scipy.interpolate import Akima1DInterpolator
import matplotlib.pyplot as plt
from PyQt5.QtGui import QPixmap, QImage, QColor, QPalette, QIcon, QFont
import math


class DashboardTab(QWidget):
    def __init__(self):
        super().__init__()

        self.setStyleSheet("background-color: gray;")

        # Create labels to display maximum values
        self.max_rop_label = QLabel("Max ROP:")
        self.max_buckling_label = QLabel("Max Buckling:")
        self.max_torque_label = QLabel("Max Torque:")
        self.max_drag_label = QLabel("Max Drag:")
        self.max_wob_label = QLabel("Max WOB:")
        self.max_rpm_label = QLabel("Max RPM:")

        # Create labels to display the values (with green and italics)
        self.max_rop_value = QLabel()
        self.max_buckling_value = QLabel()
        self.max_torque_value = QLabel()
        self.max_drag_value = QLabel()
        self.max_wob_value = QLabel()
        self.max_rpm_value = QLabel()

        # Create a layout for the dashboard
        dashboard_layout = QGridLayout()

        label_palette = self.max_rop_label.palette()
        label_palette.setColor(self.max_rop_label.foregroundRole(), QColor("yellow"))
        self.max_rop_label.setPalette(label_palette)
        self.max_buckling_label.setPalette(label_palette)
        self.max_torque_label.setPalette(label_palette)
        self.max_drag_label.setPalette(label_palette)
        self.max_wob_label.setPalette(label_palette)
        self.max_rpm_label.setPalette(label_palette)

        dashboard_layout.addWidget(self.max_rop_label, 0, 0)
        dashboard_layout.addWidget(self.max_rop_value, 0, 1)
        dashboard_layout.addWidget(self.max_buckling_label, 0, 2)
        dashboard_layout.addWidget(self.max_buckling_value, 0, 3)
        dashboard_layout.addWidget(self.max_torque_label, 1, 0)
        dashboard_layout.addWidget(self.max_torque_value, 1, 1)
        dashboard_layout.addWidget(self.max_drag_label, 1, 2)
        dashboard_layout.addWidget(self.max_drag_value, 1, 3)
        dashboard_layout.addWidget(self.max_wob_label, 2, 0)
        dashboard_layout.addWidget(self.max_wob_value, 2, 1)
        dashboard_layout.addWidget(self.max_rpm_label, 2, 2)
        dashboard_layout.addWidget(self.max_rpm_value, 2, 3)

        # Create a frame to display conditions
        conditions_frame = QFrame()
        conditions_frame.setStyleSheet("background-color: white; border: 1px solid black;")

        conditions_layout = QVBoxLayout()
        self.torque_condition_label = QLabel("Torque Condition:")
        self.drag_condition_label = QLabel("Drag Condition:")
        self.buckling_condition_label = QLabel("Buckling Condition:")
        self.torque_condition_value = QLabel()
        self.drag_condition_value = QLabel()
        self.buckling_condition_value = QLabel()

        conditions_layout.addWidget(self.torque_condition_label)
        conditions_layout.addWidget(self.torque_condition_value)
        conditions_layout.addWidget(self.drag_condition_label)
        conditions_layout.addWidget(self.drag_condition_value)
        conditions_layout.addWidget(self.buckling_condition_label)
        conditions_layout.addWidget(self.buckling_condition_value)

        conditions_frame.setLayout(conditions_layout)
        dashboard_layout.addWidget(conditions_frame, 3, 0, 1, 4)  # Spanning 1 row, 4 columns

        self.setLayout(dashboard_layout)

        # Set font styles for the labels
        font = QFont("Arial", 16, QFont.Bold)
        self.max_rop_label.setFont(font)
        self.max_buckling_label.setFont(font)
        self.max_torque_label.setFont(font)
        self.max_drag_label.setFont(font)
        self.max_wob_label.setFont(font)
        self.max_rpm_label.setFont(font)

        value_font = QFont("Arial", 14, QFont.Bold)
        value_palette = self.max_rop_value.palette()
        value_palette.setColor(self.max_rop_value.foregroundRole(), QColor("green"))
        self.max_rop_value.setPalette(value_palette)
        self.max_rop_value.setFont(value_font)
        self.max_rop_value.setStyleSheet("font-style: italic;")
        self.max_buckling_value.setPalette(value_palette)
        self.max_buckling_value.setFont(value_font)
        self.max_buckling_value.setStyleSheet("font-style: italic;")
        self.max_torque_value.setPalette(value_palette)
        self.max_torque_value.setFont(value_font)
        self.max_torque_value.setStyleSheet("font-style: italic;")
        self.max_drag_value.setPalette(value_palette)
        self.max_drag_value.setFont(value_font)
        self.max_drag_value.setStyleSheet("font-style: italic;")
        self.max_wob_value.setPalette(value_palette)
        self.max_wob_value.setFont(value_font)
        self.max_wob_value.setStyleSheet("font-style: italic;")
        self.max_rpm_value.setPalette(value_palette)
        self.max_rpm_value.setFont(value_font)
        self.max_rpm_value.setStyleSheet("font-style: italic;")

        condition_font = QFont("Arial", 12, QFont.Bold)
        self.torque_condition_label.setFont(condition_font)
        self.drag_condition_label.setFont(condition_font)
        self.buckling_condition_label.setFont(condition_font)
        self.torque_condition_value.setFont(condition_font)
        self.drag_condition_value.setFont(condition_font)
        self.buckling_condition_value.setFont(condition_font)

    def update_dashboard(self, max_rop, max_buckling, max_torque, max_drag, max_wob, max_rpm):
        self.max_rop_value.setText(f"<font color='white' face='Courier New' size='5'>{max_rop} ft/hr</font>")
        self.max_buckling_value.setText(f"<font color='white' face='Courier New' size='5'>{max_buckling} ft/hr</font>")
        self.max_torque_value.setText(f"<font color='white' face='Courier New' size='5'>{max_torque} lb-ft</font>")
        self.max_drag_value.setText(f"<font color='white' face='Courier New' size='5'>{max_drag} lbf</font>")
        self.max_wob_value.setText(f"<font color='white' face='Courier New' size='5'>{max_wob} lb</font>")
        self.max_rpm_value.setText(f"<font color='white' face='Courier New' size='5'>{max_rpm}</font>")

        # Torque condition
        if max_torque > 100000:
            self.torque_condition_value.setText("<font color='red'>Abnormal Torque</font>")
        else:
            self.torque_condition_value.setText("<font color='green'>Normal Torque</font>")

        # Drag condition
        if max_drag > 50000:
            self.drag_condition_value.setText("<font color='red'>Abnormal Drag forces</font>")
        else:
            self.drag_condition_value.setText("<font color='green'>Normal Drag</font>")

        # Buckling condition
        if max_wob > max_buckling:
            self.buckling_condition_value.setText("<font color='red'>Buckling</font>")
        else:
            self.buckling_condition_value.setText("<font color='green'>No Buckling</font>")

    def reset_dashboard(self):
        self.max_rop_value.setText("")
        self.max_buckling_value.setText("")
        self.max_torque_value.setText("")
        self.max_drag_value.setText("")
        self.max_rpm_value.setText("")
        self.max_wob_value.setText("")



class PathInputWidget(QWidget):
    def __init__(self):
        super().__init__()

        self.num_points_label = QLabel("Number of Targets:")
        self.num_points_input = QLineEdit()
        self.num_points_input.setFixedSize(300, 20)
        self.path_input_button = QPushButton("Add")
        self.path_input_button.setStyleSheet("QPushButton { background-color: green; color: black; }")
        self.path_input_button.setFixedSize(100, 20)

        # Create a scroll area to contain the layout
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)  # Allow the content widget to resize within the scroll area

        # Create a widget to hold the layout
        self.scroll_content = QWidget()
        self.scroll_area.setWidget(self.scroll_content)

        # Create a vertical layout for the scroll content
        self.scroll_layout = QVBoxLayout(self.scroll_content)

        self.scroll_layout.addWidget(self.num_points_label)
        self.scroll_layout.addWidget(self.num_points_input)
        self.scroll_layout.addWidget(self.path_input_button)
        self.setLayout(self.scroll_layout)

        self.path_input_button.clicked.connect(self.collect_path_data)

    def collect_path_data(self):
        try:
            num_points = int(self.num_points_input.text())
        except ValueError:
            return

        self.coords = []
        for i in range(num_points):
            point_label = QLabel(f"Enter coordinates for Targets {i + 1} (x y z):")
            point_input = QLineEdit()
            self.scroll_layout.addWidget(point_label)
            self.scroll_layout.addWidget(point_input)

            # Collect entered coordinates
            self.coords.append(point_input)


class MultilateralWellsTab(QWidget):
    def __init__(self):
        super().__init__()

        self.instruction_label = QLabel("Enter Coordinates using spaces (eg. x y z)")
        font = QFont()
        font.setBold(True)
        font.setItalic(True)
        self.instruction_label.setFont(font)
        self.instruction_label.setStyleSheet("color: red;")

        self.num_paths_label = QLabel("Number of Wells:")
        self.num_paths_input = QLineEdit()
        self.num_paths_input.setFixedSize(300, 20)
        self.add_button = QPushButton("Add Path")
        self.add_button.setStyleSheet("QPushButton { background-color: orange; color: black; }")
        self.add_button.setFixedSize(100, 20)

        self.generate_button = QPushButton("Generate 3D Plot")
        self.generate_button.setStyleSheet("QPushButton { background-color: black; color: white; }")
        self.generate_button.setFixedSize(200, 30)  # Set a fixed size for the button
        self.paths = []

        # Create a scroll area and a widget to hold the content
        self.scroll_area = QScrollArea()
        self.scroll_widget = QWidget()
        self.scroll_widget.setStyleSheet("background-color: white;")
        self.scroll_layout = QVBoxLayout(self.scroll_widget)

        self.scroll_layout.addWidget(self.instruction_label)
        self.scroll_layout.addWidget(self.num_paths_label)
        self.scroll_layout.addWidget(self.num_paths_input)
        self.scroll_layout.addWidget(self.add_button)

        self.scroll_area.setWidget(self.scroll_widget)
        self.scroll_area.setWidgetResizable(True)  # Allow the content widget to resize within the scroll area

        layout = QVBoxLayout(self)
        layout.addWidget(self.scroll_area)
        layout.addWidget(self.generate_button)

        self.add_button.clicked.connect(self.add_paths)
        self.generate_button.clicked.connect(self.generate_plot)

    def add_paths(self):
        try:
            num_paths = int(self.num_paths_input.text())
        except ValueError:
            return

        for _ in range(num_paths):
            path_input_widget = PathInputWidget()
            self.paths.append(path_input_widget)
            self.scroll_layout.addWidget(path_input_widget)  # Add the path widget to the scroll layout

    def generate_plot(self):

        try:
            fig = plt.figure()
            ax = fig.add_subplot(111, projection='3d')

            colors = plt.cm.jet(np.linspace(0, 1, len(self.paths)))
            for path_index, path_widget in enumerate(self.paths):
                coords = path_widget.coords
                if not coords:
                    continue

                x_coords = []
                y_coords = []
                z_coords = []
                for point_input in coords:
                    x, y, z = map(float, point_input.text().split())
                    # Increase x by a factor
                    x_coords.append(x + len(x_coords) * 0.000000001)
                    y_coords.append(y)
                    z_coords.append(z)

                ax.scatter(x_coords, y_coords, z_coords, c=colors[path_index], marker='.')

                akima_interpolator = Akima1DInterpolator(x_coords, y_coords)
                interp_x = np.linspace(min(x_coords), max(x_coords), num=100)
                interp_y = akima_interpolator(interp_x)

                akima_interpolator_z = Akima1DInterpolator(x_coords, z_coords)
                interp_z = akima_interpolator_z(interp_x)

                ax.plot(interp_x, interp_y, interp_z, c=colors[path_index], label=f'Well {path_index + 1}')


            ax.invert_zaxis()
            ax.set_xlabel('X')
            ax.set_ylabel('Y')
            ax.set_zlabel('Z')
            ax.set_title('Multilateral Well-Paths')
            ax.legend()

            # Show a triangle scatter marker at (0, 0, 0)
            ax.scatter([0], [0], [0], c='black', marker='^', s=100,
                       label='Rig Location')  # s determines the size of the marker

            plt.show()

        except Exception as e:
            self.handle_error(f"An error occurred while adding coordinates: {e} \n Enter Valid coordinates")

    def handle_error(self, error_message):
        QMessageBox.warning(self, "Error", error_message)


class DrillstringCalculator:
    def __init__(self):
        random_number = np.random.uniform(2000, 3000)
        num_values = int(random_number / 10)

        self.inclination_values = np.linspace(0, 0.8, num_values)  # degrees
        self.depth_values = np.linspace(0, random_number, num_values)

    # Function to calculate buoyed weight of the drillstring
    def calculate_buoyed_weight(self, mud_weight_initial, mud_weight_final, inner_diameter, outer_diameter,
                                drillstring_weight, inclination_angle, depth):
        return drillstring_weight + 0.0408 * (
                mud_weight_initial * inner_diameter ** 2 - mud_weight_final * outer_diameter ** 2) * math.cos(
            math.radians(inclination_angle)) * depth

    # Function to calculate weight applied to the drill bit
    def calculate_weight_on_bit(self, buoyed_weight, drillstring_length, inclination_angle):
        return buoyed_weight * drillstring_length * math.sin(math.radians(inclination_angle))

    # Function to calculate torque applied to the drillstring
    def calculate_torque(self, weight_on_bit, effective_radius, friction_coefficient, inclination_angle):
        return weight_on_bit * effective_radius * friction_coefficient * math.cos(math.radians(inclination_angle))

    # Function to calculate drag force acting on the drillstring
    def calculate_drag_force(self, friction_factor, annulus_area, pressure_coefficient, projected_area,
                             inclination_angle, depth):
        return friction_factor * annulus_area * math.cos(
            math.radians(inclination_angle)) + pressure_coefficient * projected_area * depth

    # Function to calculate critical buckling load of the drillstring
    def calculate_critical_buckling_load(self, modulus_of_elasticity, moment_of_inertia, depth):
        column_length = depth  # Just an example; replace this with the correct column length calculation
        if column_length == 0:
            return 0  # or any other default value or action you prefer
        return (np.pi ** 2 * modulus_of_elasticity * moment_of_inertia) / column_length ** 2

    # Function to calculate Rate of Penetration (ROP)
    def calculate_rop(self, weight_on_bit, specific_energy, a, b, k, inclination, depth):
        return k * weight_on_bit * specific_energy * np.abs(1 - a * inclination - b * depth)*50

    # Function to calculate RPM from ROP and other parameters
    def calculate_rpm(self, rop, weight_on_bit, specific_energy, a, b, k, inclination, depth):
        denominator = k * weight_on_bit * specific_energy * (1 + a * inclination + b * depth)

        # Handle invalid values in the denominator to avoid RuntimeWarning
        denominator = np.where(denominator == 0, 1e-6, denominator)  # Set zero values to a small positive value

        return rop / denominator

    def calculate_parameters(self):
        # Real data inputs
        mud_weight_initial = 13.8  # lb/gallon
        mud_weight_final = 12.2  # lb/gallon
        inner_diameter = 8.5  # inches
        outer_diameter = 9.0  # inches
        drillstring_weight = 16000.0  # lb
        drillstring_length = 250.0  # feet
        effective_radius = 12.0  # feet
        friction_coefficient = 0.12  # dimensionless
        friction_factor = 0.3  # dimensionless
        annulus_area = 2.5  # square inches
        pressure_coefficient = 0.15  # dimensionless
        projected_area = 35.0  # square inches
        modulus_of_elasticity = 30e6  # psi
        moment_of_inertia = 110.0  # inches^4

        # Constants for ROP calculation
        specific_energy = 0.005  # ft/lb (example value, adjust as needed)
        k = 0.001  # Constant factor (example value, adjust as needed)
        a = 0.01  # Coefficient for inclination (adjust as needed)
        b = 0.005  # Coefficient for depth (adjust as needed)

        # Create empty lists to store the parameter values
        inclination_data = []
        depth_data = []
        buoyed_weight_data = []
        weight_on_bit_data = []
        torque_data = []
        drag_force_data = []
        critical_buckling_load_data = []
        rop_data = []
        rpm_data = []

        # Calculate optimized parameters for each inclination and depth
        counter = 0  # Counter variable to limit the iterations
        for inclination, depth in zip(self.inclination_values, self.depth_values):
            # Calculate the parameters
            buoyed_weight = self.calculate_buoyed_weight(mud_weight_initial, mud_weight_final, inner_diameter,
                                                         outer_diameter, drillstring_weight, inclination,depth)
            weight_on_bit = self.calculate_weight_on_bit(buoyed_weight, drillstring_length, inclination)
            torque = self.calculate_torque(weight_on_bit, effective_radius, friction_coefficient, inclination)
            drag_force = self.calculate_drag_force(friction_factor, annulus_area, pressure_coefficient, projected_area,
                                                   inclination, depth)
            critical_buckling_load = self.calculate_critical_buckling_load(modulus_of_elasticity, moment_of_inertia,
                                                                           depth)
            rop = self.calculate_rop(weight_on_bit, specific_energy, a, b, k, inclination, depth)
            rpm = self.calculate_rpm(rop, weight_on_bit, specific_energy, a, b, k, inclination, depth)

            # Store the parameter values in the respective lists
            inclination_data.append(inclination)
            depth_data.append(depth)
            buoyed_weight_data.append(round(buoyed_weight, 2))
            weight_on_bit_data.append(round(weight_on_bit, 2))
            torque_data.append(round(torque, 2))
            drag_force_data.append(round(drag_force, 2))
            critical_buckling_load_data.append(round(critical_buckling_load, 2))
            rop_data.append(round(rop, 2))
            rpm_data.append(round(rpm, 2))

        return inclination_data, depth_data, buoyed_weight_data, weight_on_bit_data, \
            torque_data, drag_force_data, critical_buckling_load_data, rop_data, rpm_data


# Create an instance of the DrillstringCalculator class
calculator = DrillstringCalculator()

# Calculate and retrieve the parameters
inclination_data, depth_data, buoyed_weight_data, weight_on_bit_data, torque_data, drag_force_data, \
    critical_buckling_load_data, rop_data, rpm_data = calculator.calculate_parameters()


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Wellbore Visualization")

        # Set application icon
        self.setWindowIcon(QIcon("oil_drill"))  # Replace with the actual icon file path

        # Set placeholder text
        placeholder_label = QLabel("Created by Joel")
        placeholder_label.setAlignment(Qt.AlignRight)  # Align the text to the right
        placeholder_label.setStyleSheet("color: yellow;")

        self.layout = QVBoxLayout()
        self.tab_widget = QTabWidget()

        # Set the background color for the whole window
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(0, 128, 128))  # Sea blue color (R, G, B)
        self.setPalette(palette)

        # Create the label for "Well Visualization App"
        self.title_label = QLabel("Well Visualization App")
        self.title_label.setStyleSheet("QLabel { font-size: 24px; color: white; }")
        self.title_label.setAlignment(Qt.AlignCenter)

        # Layout for the title label and tab widget
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.title_label)
        self.layout.addWidget(placeholder_label)  # placeholder label

        self.simulation_tab = QWidget()
        self.tab_widget.addTab(self.simulation_tab, "Simulation/Well Plan")

        self.multilateral_wells_tab = MultilateralWellsTab()

        self.start_button = QPushButton("Start")
        self.start_button.setStyleSheet("QPushButton { background-color: green; color: white; }")
        self.start_button.clicked.connect(self.start_simulation)
        self.start_button.setFixedSize(80, 30)  # Set a fixed size for the button

        self.stop_button = QPushButton("Stop")
        self.stop_button.setStyleSheet("QPushButton { background-color: red; color: white; }")
        self.stop_button.clicked.connect(self.stop_simulation)
        self.stop_button.setEnabled(False)  # Disable the stop button initially
        self.stop_button.setFixedSize(80, 30)  # Set a fixed size for the button

        self.clear_button = QPushButton("Clear")
        self.clear_button.setStyleSheet("QPushButton { background-color: cyan; color: black; }")
        self.clear_button.clicked.connect(self.clear_inputs)
        self.clear_button.setFixedSize(80, 30)  # Set a fixed size for the button

        self.export_button = QPushButton("Export Data")
        self.export_button.setStyleSheet("QPushButton { background-color: black; color: red; }")
        self.export_button.setFixedSize(90, 30)
        self.export_button.clicked.connect(self.export_to_xls)

        self.setup_simulation_tab()
        self.setup_results_tab()
        self.setup_results_2_tab()  # Call the method to set up the "Results 2" tab
        self.multilateral_wells_tab

        self.layout.addWidget(self.tab_widget)
        self.button_layout = QHBoxLayout()
        self.button_layout.addWidget(self.start_button)
        self.button_layout.addWidget(self.stop_button)
        self.button_layout.addWidget(self.clear_button)
        self.button_layout.addWidget(self.export_button)
        self.layout.addLayout(self.button_layout)
        self.setLayout(self.layout)

        self.tab_widget.addTab(self.multilateral_wells_tab, "Multilateral Wells")

        # Add the DashboardTab to your main window's tab widget
        self.dashboard_tab = DashboardTab()
        self.tab_widget.addTab(self.dashboard_tab, "Dashboard")
        # Connect the Clear button to the reset_dashboard method in the DashboardTab
        self.clear_button.clicked.connect(self.dashboard_tab.reset_dashboard)

        self.animation = None

    def setup_simulation_tab(self):
        layout = QHBoxLayout()

        self.plot_widget = QGraphicsView()
        self.plot_widget.setFixedWidth(650)  # Set the maximum width
        layout.addWidget(self.plot_widget)

        form_layout = QFormLayout()
        self.num_targets_edit = QLineEdit()
        self.kop_x_edit = QLineEdit()
        self.kop_y_edit = QLineEdit()
        self.tvd_kop_edit = QLineEdit()
        self.target_coords_edit = QTextEdit()
        self.target_coords_edit.setMaximumHeight(100)  # Set the maximum height to your desired value
        self.target_coords_edit.setPlaceholderText("Enter target coordinates separated by comma (X, Y, Z)")
        self.kop_x_edit.setPlaceholderText("Should be Zero")
        self.kop_y_edit.setPlaceholderText("Should be Zero")

        self.output_table = QTableWidget()
        form_layout.addRow("Number of targets:", self.num_targets_edit)
        form_layout.addRow("KOP X coordinate:", self.kop_x_edit)
        form_layout.addRow("KOP Y coordinate:", self.kop_y_edit)
        form_layout.addRow("TVD at KOP:", self.tvd_kop_edit)
        form_layout.addRow("Target Coordinates:",
                           self.target_coords_edit)  # Use QTextEdit for entering multiple target coordinates
        form_layout.addRow("data output:", self.output_table)
        layout.addLayout(form_layout)

        self.output_table.setColumnCount(7)
        self.output_table.setEditTriggers(QTableWidget.NoEditTriggers)  # Set the edit triggers to NoEditTriggers
        self.output_table.setHorizontalHeaderLabels(['EAST', 'NORTH', 'TVD', 'INC', 'AZI', 'MD', 'DLS (deg)'])

        self.simulation_tab.setLayout(layout)

    def setup_results_tab(self):
        self.results_tab = QWidget()
        self.tab_widget.addTab(self.results_tab, "Results")

        self.results_layout = QVBoxLayout()
        self.results_tab.setLayout(self.results_layout)

        # Create a new QHBoxLayout to stack the QGraphicsView widgets horizontally
        hbox = QHBoxLayout()

        # Create the 2D wellbore plot QGraphicsView
        self.results_view = QGraphicsView()
        # self.results_view.setMaximumHeight(600)
        hbox.addWidget(self.results_view)

        # Create the ROP plot QGraphicsView
        self.ROP_plot_view = QGraphicsView()
        self.results_view.setMinimumHeight(400)
        hbox.addWidget(self.ROP_plot_view)

        # Add the QHBoxLayout to the QVBoxLayout
        self.results_layout.addLayout(hbox)

        # Add a QLabel for the drillstring_table label
        drillstring_label = QLabel("Drilling Output Parameters")
        drillstring_label.setStyleSheet("font-weight: bold; font-size: 14px; margin-bottom: 10px;")
        self.results_layout.addWidget(drillstring_label)

        # Create the drillstring table widget
        self.drillstring_table = QTableWidget()
        self.drillstring_table.setEditTriggers(QTableWidget.NoEditTriggers)  # Set the edit triggers to NoEditTriggers
        self.drillstring_table.setColumnCount(9)
        self.drillstring_table.setHorizontalHeaderLabels(['Inclination (degrees)', 'Depth (feet)', 'Buoyed Weight (lb)',
                                                          'Weight on Bit (lb)', 'Torque (lb-ft)', 'Drag Force (lb)',
                                                          'Crit Buckling Load (lb)', 'ROP (ft/hr)', 'RPM (rev/minute)'])
        self.results_layout.addWidget(self.drillstring_table)

    def setup_results_2_tab(self):
        self.results_2_tab = QWidget()
        self.tab_widget.addTab(self.results_2_tab, "Results 2")

        self.results_2_layout = QGridLayout()
        self.results_2_tab.setLayout(self.results_2_layout)

        # Create QGraphicsView instances for the other frames
        self.buckling_view = QGraphicsView()
        self.results_2_layout.addWidget(self.buckling_view, 0, 1)

        self.WOB_view = QGraphicsView()
        self.results_2_layout.addWidget(self.WOB_view, 0, 2)

        # Create QGraphicsView instances for torque and drag plots
        self.torque_drag_view = QGraphicsView()
        self.results_2_layout.addWidget(self.torque_drag_view, 1,
                                        1)  # Place torque_drag_view in the second row, first column

        # Create QGraphicsView instances for dogleg plot
        self.dogleg_view = QGraphicsView()
        self.results_2_layout.addWidget(self.dogleg_view, 1, 2)  # Place dogleg_view in the second row, second column

    def export_to_xls(self):
        try:
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "",
                                                       "Excel Files (*.xls);;All Files (*)",
                                                       options=options)

            if file_path:
                workbook = xlwt.Workbook(encoding="utf-8")

                # Export Drillstring Table
                drillstring_worksheet = workbook.add_sheet("Drillstring Table")
                drillstring_headers = ['Inclination', 'Depth', 'Buoyed Weight', 'Weight on Bit', 'Torque', 'Drag Force',
                                       'Critical Buckling', 'ROP', 'RPM']
                self.export_table_to_worksheet(self.drillstring_table, drillstring_worksheet, drillstring_headers)

                # Export Output Table
                output_worksheet = workbook.add_sheet("Output Table")
                output_headers = ['EAST', 'NORTH', 'TVD', 'INC', 'AZI', 'MD', 'DLS (deg)']
                self.export_table_to_worksheet(self.output_table, output_worksheet, output_headers)

                workbook.save(file_path)
                QMessageBox.information(self, "Export Complete", "Data exported to Excel file successfully.")

        except Exception as e:
            print("An error occurred:", e)

    def export_table_to_worksheet(self, table_widget, worksheet, headers):
        # Export table headers
        for col_idx, header in enumerate(headers):
            worksheet.write(0, col_idx, header)

        # Export table data
        for row_idx in range(table_widget.rowCount()):
            for col_idx in range(table_widget.columnCount()):
                item = table_widget.item(row_idx, col_idx)
                if item is not None:
                    worksheet.write(row_idx + 1, col_idx, item.text())

    def start_simulation(self):

        try:
            if not self.validate_inputs():
                return

            self.start_button.setEnabled(False)  # Disable the start button
            self.stop_button.setEnabled(True)  # Enable the stop button

            num_targets = int(self.num_targets_edit.text())
            kop_x = float(self.kop_x_edit.text())
            kop_y = float(self.kop_y_edit.text())
            tvd_kop = float(self.tvd_kop_edit.text())

            target_coords_text = self.target_coords_edit.toPlainText()
            target_coords = []
            lines = target_coords_text.split("\n")
            for line in lines:
                coordinates = line.strip()
                if coordinates:
                    x, y, z = map(float, coordinates.split(","))
                    target_coords.append([x, y, z])
            target_coords = np.array(target_coords)
            if len(target_coords) != num_targets:
                self.show_error_message(f"Number of target coordinates does not match the specified number of targets "
                                        f"({num_targets}). Please enter {num_targets} target coordinates.")
                return

            # Surface coordinates
            surface_coords = np.array([0, 0, 0])

            surface_x, surface_y, surface_z = surface_coords
            kop_z = tvd_kop

            targets_x = target_coords[:, 0]
            targets_y = target_coords[:, 1]
            targets_z = target_coords[:, 2]

            interpolating_x = np.insert(targets_x, 0, kop_x)
            interpolating_y = np.insert(targets_y, 0, kop_y)
            interpolating_z = np.insert(targets_z, 0, kop_z)

            def akima1DInterp(x, y, z):
                interp_func_x = Akima1DInterpolator(z, x)
                interp_func_y = Akima1DInterpolator(z, y)
                interp_func_z = Akima1DInterpolator(z, z)

                z_interp = np.arange(z[0], z[-1], 10)
                x_interp = interp_func_x(z_interp)
                y_interp = interp_func_y(z_interp)
                z_interp = interp_func_z(z_interp)

                return x_interp, y_interp, z_interp

            def calAzimuthInc(x, y, z):
                delta_x = np.diff(x)
                delta_y = np.diff(y)
                delta_z = np.diff(z)
                horizontal_distance = np.sqrt(delta_x ** 2 + delta_y ** 2)
                vertical_distance = delta_z

                inc_rad = np.arctan2(horizontal_distance, vertical_distance)
                inc_rad = np.concatenate((np.array([0]), inc_rad))
                inc_deg = np.rad2deg(inc_rad)

                az_rad = np.arctan2(delta_y, delta_x)
                az_rad = np.concatenate((np.array([0]), az_rad))
                az_deg = np.rad2deg(az_rad)

                return ((az_rad, az_deg), (inc_rad, inc_deg))

            def calculateMeasuredDepth(x, y, z):
                delta_x = np.diff(x)
                delta_y = np.diff(y)
                delta_z = np.diff(z)
                segment_lengths = np.sqrt(delta_x ** 2 + delta_y ** 2 + delta_z ** 2)
                measured_depth = np.cumsum(segment_lengths)
                measured_depth = np.concatenate(([0], measured_depth))

                return measured_depth

            def dogleg(inc1, azi1, inc2, azi2):
                dogleg_angle = np.arccos(np.cos(inc2 - inc1) - np.sin(inc1) * np.sin(inc2) * (1 - np.cos(azi2 - azi1)))
                return dogleg_angle/30

            def calculate_deviation(p1, p2, p3):
                #  function to calculate deviation based on hydraulic pressures
                # Replace this with your actual RSS implementation
                deviation_factor = 0.05
                x_deviated = p1 * deviation_factor
                y_deviated = p2 * deviation_factor
                z_deviated = p3 * deviation_factor
                return x_deviated, y_deviated, z_deviated

            x, y, z = akima1DInterp(interpolating_x, interpolating_y, interpolating_z)

            surface_to_kop_zs = np.arange(surface_z, kop_z, 10)
            z = np.concatenate((surface_to_kop_zs, z))
            x = np.concatenate((np.full((len(surface_to_kop_zs),), surface_x), x))
            y = np.concatenate((np.full((len(surface_to_kop_zs),), surface_y), y))

            # Assume the hydraulic pressures are available as p1, p2, and p3
            p1, p2, p3 = 10.0, 15.0, 20.0

            # Calculate the deviation using the RSS function
            x_deviated, y_deviated, z_deviated = calculate_deviation(p1, p2, p3)

            # Deviate the coordinates by the calculated RSS deviation
            x_deviated += x
            y_deviated += y
            z_deviated += z

            inclination, azimuth = calAzimuthInc(x_deviated, y_deviated, z_deviated)
            measured_depth = calculateMeasuredDepth(x_deviated, y_deviated, z_deviated)

            dogleg_angles_rad = np.array([dogleg(inc1, azi1, inc2, azi2) for inc1, azi1, inc2, azi2 in
                                          zip(inclination[0], azimuth[0], inclination[0][1:], azimuth[0][1:])])
            dogleg_angles_deg = np.rad2deg(dogleg_angles_rad)

            # Update the table with data
            self.output_table.setRowCount(len(x_deviated))
            for idx, (x_val, y_val, z_val, inc_val, azi_val, md_val, dls_val) in enumerate(
                    zip(x_deviated, y_deviated, z_deviated, inclination[0], azimuth[0], measured_depth,
                        dogleg_angles_deg)):
                self.output_table.setItem(idx, 0, QTableWidgetItem(f"{x_val:.2f}"))
                self.output_table.setItem(idx, 1, QTableWidgetItem(f"{y_val:.2f}"))
                self.output_table.setItem(idx, 2, QTableWidgetItem(f"{z_val:.2f}"))
                self.output_table.setItem(idx, 3, QTableWidgetItem(f"{inc_val:.2f}"))
                self.output_table.setItem(idx, 4, QTableWidgetItem(f"{azi_val:.2f}"))
                self.output_table.setItem(idx, 5, QTableWidgetItem(f"{md_val:.2f}"))
                self.output_table.setItem(idx, 6, QTableWidgetItem(f"{dls_val:.2f}"))

            # Generate the 2D plot for X against Z
            fig_2d = plt.figure(figsize=(6, 4))  # Set the width to 8 inches and height to 5 inches
            ax_2d = fig_2d.add_subplot(111)
            ax_2d.plot(x, z, c='g', label='well plan')
            ax_2d.scatter(x_deviated, z_deviated, c='pink', label='simulated path')
            ax_2d.scatter(target_coords[:, 0], target_coords[:, 2], c='r', marker='o', label='Targets')
            ax_2d.set_xlabel('HD')
            ax_2d.set_ylabel('TVD')
            ax_2d.set_title('2D View')
            ax_2d.invert_yaxis()
            # Show a triangle scatter marker at (0, 0, 0)
            ax_2d.scatter(0, 0, c='black', marker='^', s=100, label='Rig Location')
            ax_2d.legend()

            # Convert the 2D plot to a QPixmap
            canvas_2d = fig_2d.canvas
            canvas_2d.draw()
            width_2d, height_2d = canvas_2d.get_width_height()
            plot_img_2d = np.frombuffer(canvas_2d.tostring_rgb(), dtype=np.uint8).reshape(height_2d, width_2d, 3)
            qimage_2d = QImage(plot_img_2d, width_2d, height_2d, QImage.Format_RGB888)
            pixmap_2d = QPixmap.fromImage(qimage_2d)

            # Create a QGraphicsPixmapItem with the 2D plot
            pixmap_item_2d = QGraphicsPixmapItem(pixmap_2d)

            # Create a QGraphicsScene and add the 2D plot pixmap item
            scene_2d = QGraphicsScene()
            scene_2d.addItem(pixmap_item_2d)

            # Set the scene in the QGraphicsView widget in the Results tab
            self.results_view.setScene(scene_2d)
            self.results_view.setSceneRect(pixmap_item_2d.boundingRect())

            # Close the 2D figure to avoid memory consumption
            plt.close(fig_2d)

            # Generate the 3D plot on QviewWidget
            fig = plt.figure(figsize=(6, 6))
            ax = fig.add_subplot(111, projection='3d')
            ax.plot(x, y, z, c='g', label='well plan')
            ax.scatter(target_coords[:, 0], target_coords[:, 1], target_coords[:, 2], c='r', marker='o',
                       label='Targets')
            ax.set_xlabel('EAST')
            ax.set_ylabel('NORTH')
            ax.set_zlabel('TVD')
            plt.title('Planned Well Path')
            ax.invert_zaxis()
            # ax.set_xlim(0, 100)
            # ax.set_ylim(0, 100)

            # Show a triangle scatter marker at (0, 0, 0)
            ax.scatter([0], [0], [0], c='black', marker='^', s=100,
                       label='Rig Location')  # s determines the size of the marker
            # Add a legend for the specific marker
            ax.legend()

            # Convert the 3D plot to a QPixmap
            canvas = fig.canvas
            canvas.draw()
            width, height = canvas.get_width_height()
            plot_img = np.frombuffer(canvas.tostring_rgb(), dtype=np.uint8).reshape(height, width, 3)
            qimage = QImage(plot_img, width, height, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimage)

            # Create a QGraphicsPixmapItem with the 3D plot
            pixmap_item = QGraphicsPixmapItem(pixmap)

            # Create a QGraphicsScene and add the 3D plot pixmap item
            scene = QGraphicsScene()
            scene.addItem(pixmap_item)

            # Set the scene in the QGraphicsView widget in the Simulation tab
            self.plot_widget.setScene(scene)
            self.plot_widget.setSceneRect(pixmap_item.boundingRect())

            # Close the 3D figure to avoid memory consumption
            plt.close(fig)

            # Create a new figure and subplot with 3D projection (animated)*********************
            fig_animated = plt.figure()
            ax_animated = fig_animated.add_subplot(111, projection='3d')

            # Set labels for the axes
            ax_animated.set_xlabel('EAST')
            ax_animated.set_ylabel('NORTH')
            ax_animated.set_zlabel('TVD')


            # Invert the z-axis (TVD) for better visualization
            ax_animated.invert_zaxis()

            # Set the title
            plt.title('RSS-Steered Wellpath')


            # Plot the rig location as a big black triangle marker
            rig_marker = ax_animated.scatter([0], [0], [0], c='black', marker='^', s=200, label='Rig Location')

            # Plot the targets
            target_markers = ax_animated.scatter(target_coords[:, 0], target_coords[:, 1], target_coords[:, 2],
                                                 c='r', marker='o', label='Targets')

            well_plan = ax_animated.plot(x, y, z, c='g', label='well plan')

            # Initialize the line for the simulated well path
            line_deviated, = ax_animated.plot([], [], [], c='violet', label='Simulated Path', marker='o')

            # Function to update the animated plot
            def update(frame):
                # Update the line data with each frame
                line_deviated.set_data(x_deviated[:frame], y_deviated[:frame])
                line_deviated.set_3d_properties(z_deviated[:frame])
                return line_deviated,

            # Create the animation
            num_frames = len(x_deviated)
            ani = FuncAnimation(fig_animated, update, frames=num_frames, interval=100, blit=True, repeat=False)

            # Add a legend to the plot
            ax_animated.legend()

            # Show the animation
            plt.show()

            # Create an instance of the DrillstringCalculator class
            calculator = DrillstringCalculator()
            # Calculate the drillstring parameters
            inclination_data, depth_data, buoyed_weight_data, weight_on_bit_data, torque_data, \
                drag_force_data, critical_buckling_load_data, rop_data, rpm_data = calculator.calculate_parameters()

            # Display the drillstring data in the results table
            self.display_drillstring_data(inclination_data, depth_data, buoyed_weight_data, weight_on_bit_data,
                                          torque_data, drag_force_data, critical_buckling_load_data, rop_data, rpm_data)

            # Generate the torque-drag plot
            self.generate_torque_drag_plot(depth_data, torque_data, drag_force_data)

            # Generate the dogleg plot
            self.generate_dogleg_plot(measured_depth, dogleg_angles_deg)

            # Generate the Buckling plot
            self.generate_buckling_plot(depth_data, critical_buckling_load_data)

            # Generate the WOB plot
            self.generate_WOB_plot(depth_data, weight_on_bit_data)

            # Generate the ROP plot
            self.generate_ROP_plot(depth_data, rop_data)

            # Display a QMessageBox to indicate simulation completion
            self.show_simulation_complete_message()

            # Calculate the maximum values
            max_rop = max(rop_data)
            max_buckling = max(critical_buckling_load_data)
            max_torque = max(torque_data)
            max_drag = max (drag_force_data)
            max_wob = max(weight_on_bit_data)
            max_rpm = max(rpm_data)

            # Update the dashboard with the maximum values
            self.dashboard_tab.update_dashboard(max_rop, max_buckling, max_torque,max_drag, max_wob, max_rpm)

        except Exception as e:
            self.handlle_error(f"An error occurred while creating well path : {e} \n Enter Valid Coordinates")

    def handlle_error(self, error_message):
        QMessageBox.warning(self, "Error", error_message)

    def calculate_minimum_curvature(self, I1, I2, Az1, Az2, MD):
        # Convert inclination angles to radians
        I1_rad = np.radians(I1)
        I2_rad = np.radians(I2)

        # Convert azimuth directions to radians
        Az1_rad = np.radians(Az1)
        Az2_rad = np.radians(Az2)

        # Calculate the dog leg angle (β)
        cos_beta = np.cos(I2_rad - I1_rad) - (np.sin(I1_rad) * np.sin(I2_rad) * (1 - np.cos(Az2_rad - Az1_rad)))
        beta = np.arccos(cos_beta)

        # Calculate Ratio Factor (RF)
        RF = 2 / beta * np.tan(beta / 2)

        # Calculate North, East, and TVD
        North = MD / 2 * (np.sin(I1_rad) * np.cos(Az1_rad) + np.sin(I2_rad) * np.cos(Az2_rad)) * RF
        East = MD / 2 * (np.sin(I1_rad) * np.sin(Az1_rad) + np.sin(I2_rad) * np.sin(Az2_rad)) * RF
        TVD = MD / 2 * (np.cos(I1_rad) + np.cos(I2_rad)) * RF

        return North, East, TVD

    def generate_ROP_plot(self, depth_data, rop_data):
        # Add random wobbling effect to ROP data
        amplitude = 0.01  # Adjust the amplitude of the wobbling
        rop_data_wobbly = rop_data + amplitude * np.random.normal(0, 1, len(rop_data))

        # Generate the ROP plot
        fig_rop = plt.figure(figsize=(6, 4))  # Set the width to 7 inches and height to 4 inches
        plt.plot(depth_data, rop_data_wobbly, marker='o', color='blue', label='ROP (ft/hr)')
        plt.xlabel('MD (feet)')
        plt.ylabel('ROP (ft/hr)')
        plt.title('Rate of Penetration (ROP) vs. Depth')
        plt.legend()

        # Convert the ROP plot to a QPixmap
        canvas_rop = fig_rop.canvas
        canvas_rop.draw()
        width_rop, height_rop = canvas_rop.get_width_height()
        plot_img_rop = np.frombuffer(canvas_rop.tostring_rgb(), dtype=np.uint8).reshape(height_rop, width_rop, 3)
        qimage_rop = QImage(plot_img_rop, width_rop, height_rop, QImage.Format_RGB888)
        pixmap_rop = QPixmap.fromImage(qimage_rop)

        # Create a QGraphicsPixmapItem with the ROP plot
        pixmap_item_rop = QGraphicsPixmapItem(pixmap_rop)

        # Create a QGraphicsScene and add the ROP plot pixmap item
        scene_rop = QGraphicsScene()
        scene_rop.addItem(pixmap_item_rop)

        # Set the scene in the QGraphicsView widget for the ROP plot
        self.ROP_plot_view.setScene(scene_rop)
        self.ROP_plot_view.setSceneRect(pixmap_item_rop.boundingRect())

        # Close the ROP figure to avoid memory consumption
        plt.close(fig_rop)

    def generate_WOB_plot(self, depth_data, weight_on_bit_data):

        amplitude = 1000  # Adjust the amplitude of the wobbling
        wob_data_wobbly = (weight_on_bit_data + amplitude * np.random.normal(0, 1, len(weight_on_bit_data)))**2

        # Generate the WOB Plot
        fig_wob = plt.figure(figsize=(6, 3))
        plt.plot(depth_data, wob_data_wobbly, linestyle='-', color='b', linewidth=2.5, marker='')
        plt.xlabel('MD (ft)')
        plt.ylabel('Weight on Bit (WOB) (lb)')
        plt.title('Weight on Bit (WOB)')
        plt.grid(True)
        plt.tight_layout()

        # Convert the WOB Plot to a QPixmap
        canvas_wob = fig_wob.canvas
        canvas_wob.draw()
        width_wob, height_wob = canvas_wob.get_width_height()
        plot_img_wob = np.frombuffer(canvas_wob.tostring_rgb(), dtype=np.uint8).reshape(height_wob, width_wob, 3)
        qimage_wob = QImage(plot_img_wob, width_wob, height_wob, QImage.Format_RGB888)
        pixmap_wob = QPixmap.fromImage(qimage_wob)

        # Create a QGraphicsPixmapItem with the WOB Plot
        pixmap_item_wob = QGraphicsPixmapItem(pixmap_wob)

        # Create a QGraphicsScene and add the WOB Plot pixmap item
        scene_wob = QGraphicsScene()
        scene_wob.addItem(pixmap_item_wob)

        # Set the scene in the QGraphicsView widget for the WOB histogram in "Results 2" tab
        self.WOB_view.setScene(scene_wob)
        self.WOB_view.setSceneRect(pixmap_item_wob.boundingRect())

        # Close the WOB figure to avoid memory consumption
        plt.close(fig_wob)

    def generate_buckling_plot(self, depth_data, critical_buckling_load_data):

        # Generate the Buckling plot
        fig_buck = plt.figure(figsize=(6, 3))
        plt.plot(depth_data, critical_buckling_load_data, linestyle='-', color='g', marker='o')
        plt.xlabel('MD (ft)')
        plt.ylabel('Critical Buckling (lb-force)')
        plt.title('Buckling')
        plt.grid(True)
        plt.tight_layout()

        # Convert the buckling plot to a QPixmap
        canvas_buck = fig_buck.canvas
        canvas_buck.draw()
        width_buck, height_buck = canvas_buck.get_width_height()
        plot_img_buck = np.frombuffer(canvas_buck.tostring_rgb(), dtype=np.uint8).reshape(height_buck, width_buck, 3)
        qimage_buck = QImage(plot_img_buck, width_buck, height_buck, QImage.Format_RGB888)
        pixmap_buck = QPixmap.fromImage(qimage_buck)

        # Create a QGraphicsPixmapItem with the WOB histogram
        pixmap_item_buck = QGraphicsPixmapItem(pixmap_buck)

        # Create a QGraphicsScene and add the WOB histogram pixmap item
        scene_buck = QGraphicsScene()
        scene_buck.addItem(pixmap_item_buck)

        # Set the scene in the QGraphicsView widget for the WOB histogram in "Results 2" tab
        self.buckling_view.setScene(scene_buck)
        self.buckling_view.setSceneRect(pixmap_item_buck.boundingRect())

        # Close the WOB figure to avoid memory consumption
        plt.close(fig_buck)

    def generate_torque_drag_plot(self, depth_data, torque_data, drag_force_data):
        amplitude = 1000  # Adjust the amplitude of the wobbling
        drag_data_wobbly = (drag_force_data + amplitude * np.random.normal(0, 1, len(drag_force_data)))
        tor_data_wobbly = (torque_data + amplitude * np.random.normal(0, 1, len(torque_data)))

        # Generate the Torque and Drag plot
        fig_torque_drag = plt.figure(figsize=(6, 3))
        plt.plot(depth_data, tor_data_wobbly, linestyle='-', color='b', marker='o', label='Torque (lb-ft)',linewidth=2.5)
        plt.plot(depth_data, drag_data_wobbly, linestyle='-', color='r', marker='o', label='Drag Force (lb)',linewidth=2.5)
        plt.xlabel('MD (ft)')
        plt.ylabel('Torque and Drag')
        plt.title('Torque and Drag vs. MD')
        plt.legend()
        plt.grid(True)
        plt.tight_layout()

        # Convert the Torque and Drag plot to QPixmap
        canvas_torque_drag = fig_torque_drag.canvas
        canvas_torque_drag.draw()
        width_torque_drag, height_torque_drag = canvas_torque_drag.get_width_height()
        plot_img_torque_drag = np.frombuffer(canvas_torque_drag.tostring_rgb(), dtype=np.uint8).reshape(
            height_torque_drag, width_torque_drag, 3)
        qimage_torque_drag = QImage(plot_img_torque_drag, width_torque_drag, height_torque_drag, QImage.Format_RGB888)
        pixmap_torque_drag = QPixmap.fromImage(qimage_torque_drag)

        # Create a QGraphicsPixmapItem with the Torque and Drag plot
        pixmap_item_torque_drag = QGraphicsPixmapItem(pixmap_torque_drag)

        # Create a QGraphicsScene and add the Torque and Drag pixmap item
        scene_torque_drag = QGraphicsScene()
        scene_torque_drag.addItem(pixmap_item_torque_drag)

        # Set the scene in the QGraphicsView widget for torque_drag_view
        self.torque_drag_view.setScene(scene_torque_drag)
        self.torque_drag_view.setSceneRect(pixmap_item_torque_drag.boundingRect())

        # Close the Torque and Drag figure to avoid memory consumption
        plt.close(fig_torque_drag)

    def generate_dogleg_plot(self, measured_depth, dogleg_angles_deg):
        # Generate the Dogleg plot
        fig_dogleg = plt.figure(figsize=(6, 3))
        plt.plot(measured_depth[1:], dogleg_angles_deg, linestyle='-', color='g', marker='o',
                 label='Dogleg Angle (deg)')
        plt.xlabel('MD (ft)')
        plt.ylabel('Dogleg Angle (deg)')
        plt.title('Dogleg Angle vs. MD')
        plt.legend()
        plt.grid(True)
        plt.tight_layout()

        # Convert the Dogleg plot to QPixmap
        canvas_dogleg = fig_dogleg.canvas
        canvas_dogleg.draw()
        width_dogleg, height_dogleg = canvas_dogleg.get_width_height()
        plot_img_dogleg = np.frombuffer(canvas_dogleg.tostring_rgb(), dtype=np.uint8).reshape(height_dogleg,
                                                                                              width_dogleg, 3)
        qimage_dogleg = QImage(plot_img_dogleg, width_dogleg, height_dogleg, QImage.Format_RGB888)
        pixmap_dogleg = QPixmap.fromImage(qimage_dogleg)

        # Create a QGraphicsPixmapItem with the Dogleg plot
        pixmap_item_dogleg = QGraphicsPixmapItem(pixmap_dogleg)

        # Create a QGraphicsScene and add the Dogleg pixmap item
        scene_dogleg = QGraphicsScene()
        scene_dogleg.addItem(pixmap_item_dogleg)

        # Set the scene in the QGraphicsView widget for dogleg_view
        self.dogleg_view.setScene(scene_dogleg)
        self.dogleg_view.setSceneRect(pixmap_item_dogleg.boundingRect())

        # Close the Dogleg figure to avoid memory consumption
        plt.close(fig_dogleg)

    def display_drillstring_data(self, inclination_data, depth_data, buoyed_weight_data, weight_on_bit_data,
                                 torque_data, drag_force_data, critical_buckling_load_data, rop_data, rpm_data):
        # Clear the previous data from the drillstring table
        self.drillstring_table.clearContents()
        self.drillstring_table.setRowCount(0)

        # Update the drillstring table with the new data
        self.drillstring_table.setRowCount(len(inclination_data))
        for idx, (inclination, depth) in enumerate(zip(inclination_data, depth_data)):
            # Populate the drillstring table with the data
            self.drillstring_table.setItem(idx, 0, QTableWidgetItem("{:.3f}".format(inclination)))
            self.drillstring_table.setItem(idx, 1, QTableWidgetItem("{:.3f}".format(depth)))
            self.drillstring_table.setItem(idx, 2, QTableWidgetItem("{:.3f}".format(buoyed_weight_data[idx])))
            self.drillstring_table.setItem(idx, 3, QTableWidgetItem("{:.3f}".format(weight_on_bit_data[idx])))
            self.drillstring_table.setItem(idx, 4, QTableWidgetItem("{:.3f}".format(torque_data[idx])))
            self.drillstring_table.setItem(idx, 5, QTableWidgetItem("{:.3f}".format(drag_force_data[idx])))
            self.drillstring_table.setItem(idx, 6, QTableWidgetItem("{:.3f}".format(critical_buckling_load_data[idx])))
            self.drillstring_table.setItem(idx, 7, QTableWidgetItem("{:.3f}".format(rop_data[idx])))
            self.drillstring_table.setItem(idx, 8, QTableWidgetItem("{:.3f}".format(rpm_data[idx])))

    def stop_simulation(self):
        self.stop_button.setEnabled(False)  # Disable the stop button
        self.start_button.setEnabled(True)  # Enable the start button

        if self.animation:
            self.animation[2].stop()  # Stop the animation timer
            self.animation = None  # Reset the animation references

    def clear_inputs(self):
        self.num_targets_edit.clear()
        self.kop_x_edit.clear()
        self.kop_y_edit.clear()
        self.tvd_kop_edit.clear()
        self.target_coords_edit.clear()

        self.clear_outputs()  # Call the clear_outputs method to clear the outputs

        # Clear inputs in the multilateral wells tab
        self.multilateral_wells_tab.num_paths_input.clear()

        # Reset widgets to their original state in the multilateral wells tab
        for path_widget in self.multilateral_wells_tab.paths:
            self.multilateral_wells_tab.scroll_layout.removeWidget(path_widget)
            path_widget.deleteLater()
        self.multilateral_wells_tab.paths = []

    def clear_outputs(self):
        self.drillstring_table.clearContents()
        self.drillstring_table.setRowCount(0)

        self.output_table.clearContents()  # Clear the table contents
        self.output_table.setRowCount(0)  # Reset the row count to 0

        self.results_view.setScene(None)  # Clear the scene in the QGraphicsView widget
        self.ROP_plot_view.setScene(None)

        self.WOB_view.setScene(None)  # Clear the scene in the QGraphicsView widget
        self.WOB_view.setScene(None)

        self.dogleg_view.setScene(None)  # Clear the scene in the QGraphicsView widget
        self.dogleg_view.setScene(None)

        self.torque_drag_view.setScene(None)  # Clear the scene in the QGraphicsView widget
        self.torque_drag_view.setScene(None)

        self.buckling_view.setScene(None)  # Clear the scene in the QGraphicsView widget
        self.buckling_view.setScene(None)

        if self.animation:
            plt.close(self.animation.fig)  # Close the animation figure
            self.animation = None  # Reset the animation reference

        self.plot_widget.setScene(None)  # Clear the scene in the QGraphicsView widget for the unanimated 3D plot

    def validate_inputs(self):
        num_targets = self.num_targets_edit.text()
        kop_x = self.kop_x_edit.text()
        kop_y = self.kop_y_edit.text()
        tvd_kop = self.tvd_kop_edit.text()
        target_coords = self.target_coords_edit.toPlainText()

        if not self.is_float(kop_x):
            self.show_error_message("Invalid KOP X coordinate. Please enter a valid number.")
            return False
        elif float(kop_x) != 0:
            self.show_error_message("KOP X coordinate should be zero.")
            return False

        if not self.is_float(kop_y):
            self.show_error_message("Invalid KOP Y coordinate. Please enter a valid number.")
            return False
        elif float(kop_y) != 0:
            self.show_error_message("KOP Y coordinate should be zero.")
            return False

        if not num_targets.isdigit() or int(num_targets) < 1:
            self.show_error_message("Invalid number of targets. Please enter a positive integer.")
            return False

        if not self.is_float(kop_x):
            self.show_error_message("Invalid KOP X coordinate. Please enter a valid number.")
            return False

        if not self.is_float(kop_y):
            self.show_error_message("Invalid KOP Y coordinate. Please enter a valid number.")
            return False

        if not self.is_float(tvd_kop):
            self.show_error_message("Invalid TVD at KOP. Please enter a valid number.")
            return False

        if not self.validate_target_coords(target_coords):
            self.show_error_message("Invalid target coordinates. Please enter valid coordinates in the format: X, Y, Z")
            return False

        return True

    def is_float(self, value):
        try:
            float(value)
            return True
        except ValueError:
            return False

    def validate_target_coords(self, target_coords):
        lines = target_coords.split("\n")
        for line in lines:
            coordinates = line.strip()
            if coordinates:
                values = coordinates.split(",")
                if len(values) != 3:
                    return False
                for value in values:
                    if not self.is_float(value):
                        return False
        return True

    def show_error_message(self, message):
        error_box = QMessageBox()
        error_box.setIcon(QMessageBox.Warning)
        error_box.setWindowTitle("Input Error")
        error_box.setText(message)
        error_box.exec_()

    def show_simulation_complete_message(self):
        message_box = QMessageBox()
        message_box.setWindowTitle("Simulation")
        message_box.setText("collecting output data to Export....")
        message_box.setIcon(QMessageBox.Information)
        message_box.exec_()

    def closeEvent(self, event):
        self.clear_outputs()

        event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # Set the Fusion style
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    # window.showMaximized()
    sys.exit(app.exec_())
