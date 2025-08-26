# Antenna Design & Analysis Suite

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Python Version](https://img.shields.io/badge/python-3.8%2B-brightgreen)
![Status](https://img.shields.io/badge/status-in%20development-orange)

An automated antenna design and analysis suite using **PyAEDT** to streamline and accelerate simulation workflows in Ansys Electronics Desktop (AEDT).

## About The Project

Designing and analyzing antennas in simulators like Ansys HFSS often involves repetitive manual tasks. This project aims to provide a comprehensive Python-based framework to automate the entire workflow, from parametric geometry creation and simulation setup to advanced post-processing and data extraction.

By leveraging the power of `PyAEDT`, this suite allows engineers and researchers to programmatically control AEDT, enabling rapid prototyping, parametric sweeps, optimization, and repeatable results.

### Key Features

* **Parametric Antenna Generation**: Programmatically create common antenna types (e.g., microstrip patches, dipoles, horns, arrays) with customizable parameters.
* **Automated Simulation Setup**: Automatically configure solution setups, frequency sweeps, boundary conditions (radiation, ports), and mesh operations.
* **Advanced Post-Processing**: Extract and plot S-parameters, VSWR, impedance, gain, directivity, and 2D/3D radiation patterns.
* **Optimization Ready**: Easily integrate with optimization libraries like `scipy.optimize` or `pygmo` to find optimal antenna dimensions based on performance goals.
* **Data Management**: Export results to standard formats such as Touchstone (.sNp), CSV, or JSON for further analysis and documentation.

## Getting Started

Follow these instructions to get a local copy up and running for development and testing.

### Prerequisites

* Python 3.8 or higher
* Ansys Electronics Desktop 2021 R1 or higher (with a valid license)

### Installation

1.  **Clone the repository:**
    ```sh
    git clone [https://github.com/Gecesars/antenna_desindesign.git](https://github.com/Gecesars/antenna_desindesign.git)
    ```
2.  **Navigate to the project directory:**
    ```sh
    cd antenna_desindesign
    ```
3.  **Install the required packages:**
    It's recommended to use a virtual environment.
    ```sh
    pip install -r requirements.txt
    ```
    The `requirements.txt` file should contain:
    ```
    pyaedt
    numpy
    matplotlib
    scipy
    ```

## Usage Example

Here is a basic example of how to create and simulate a simple rectangular microstrip patch antenna.

```python
# Import necessary libraries
from pyaedt import Hfss
import numpy as np

# Project parameters
freq_ghz = 2.4
substrate_height_mm = 1.57
substrate_er = 4.4

# --- 1. Launch AEDT and HFSS ---
# This will launch AEDT in non-graphical mode. Set to False for graphical mode.
with Hfss(specified_version="2023.2", non_graphical=False, new_desktop_session=True) as hfss:

    # --- 2. Define Antenna Geometry Parameters ---
    # (Calculations for a patch antenna would go here)
    patch_width = 38.0  # mm
    patch_length = 29.5 # mm
    feed_offset = 5.0   # mm

    # --- 3. Create Geometry ---
    # Create Substrate
    substrate = hfss.modeler.create_box(
        position=['-50', '-50', '0'],
        dimensions_list=['100', '100', -substrate_height_mm],
        name='Substrate',
        material='FR4_epoxy'
    )
    # Create Ground Plane
    gnd = hfss.modeler.create_rectangle(
        position=['-50', '-50', -substrate_height_mm],
        dimension_list=['100', '100'],
        name='GroundPlane'
    )
    hfss.assign_perfecte_boundary(gnd)

    # Create Patch
    patch = hfss.modeler.create_rectangle(
        position=[-patch_width / 2, -patch_length / 2, 0],
        dimension_list=[patch_width, patch_length],
        name='Patch'
    )
    hfss.assign_perfecte_boundary(patch)

    # Create Port
    port = hfss.modeler.create_rectangle(
        position=[-patch_width / 2, -feed_offset, 0],
        dimension_list=[0, -5, -substrate_height_mm], # Lumped port
        name='Port',
        axis='Y'
    )
    hfss.create_lumped_port_to_sheet(port.name, port_impedance=50)

    # --- 4. Setup Analysis ---
    # Create Radiation Boundary
    hfss.create_open_region(Frequency=f"{freq_ghz}GHz")

    # Add Solution Setup
    setup = hfss.create_setup(setup_name="MainSetup")
    setup.props["Frequency"] = f"{freq_ghz}GHz"
    setup.props["MaximumPasses"] = 10

    # Add Frequency Sweep
    hfss.create_linear_count_sweep(
        setupname="MainSetup",
        unit="GHz",
        freqstart=1.5,
        freqstop=3.5,
        num_of_freq_points=401,
    )

    # --- 5. Run Simulation ---
    hfss.analyze_setup("MainSetup")

    # --- 6. Post-Processing ---
    # Get S-parameter data
    s11_data = hfss.get_solution_data("S(1,1)", "MainSetup : Sweep")
    s11_db = 20 * np.log10(np.abs(s11_data.intrinsics["S(1,1)"]["Mag"]))

    print("Simulation finished. S11 data extracted.")

    # Create a far-field radiation plot
    report = hfss.post.create_report("db(GainTotal)", "3D Polar Plot")
    report.add_all_variations_to_report()

    # --- 7. Save and Close ---
    hfss.save_project()
    print("Project saved. AEDT will be closed.")

```

## Project Structure

A recommended structure for this repository:

```
.
├── .gitignore
├── LICENSE
├── README.md
├── requirements.txt
├── examples/
│   └── patch_antenna_example.py
└── src/
    └── antenna_suite/
        ├── __init__.py
        ├── antennas/
        │   ├── __init__.py
        │   ├── patch.py
        │   └── dipole.py
        ├── analysis/
        │   └── setup.py
        └── postprocessing/
            └── plotting.py
```

## Roadmap

- [ ] Add more parameterized antenna models (Horn, Vivaldi, Helical).
- [ ] Develop a class-based structure for streamlined analysis setups.
- [ ] Implement advanced post-processing functions for beamwidth, axial ratio, etc.
- [ ] Create a simple GUI wrapper using `PyQt` or `Tkinter`.
- [ ] Add comprehensive examples in the form of Jupyter Notebooks.

See the [open issues](https://github.com/Gecesars/antenna_desindesign/issues) for a full list of proposed features (and known issues).

## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1.  Fork the Project
2.  Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3.  Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4.  Push to the Branch (`git push origin feature/AmazingFeature`)
5.  Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Contact

Gecesars - [Your Project Link](https://github.com/Gecesars/antenna_desindesign)

Project Link: [https://github.com/Gecesars/antenna_desindesign](https://github.com/Gecesars/antenna_desindesign)
