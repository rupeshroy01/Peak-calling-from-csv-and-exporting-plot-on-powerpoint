Overview
This Python script is designed to analyze CSV files containing data, identify peaks in the data, and generate a PowerPoint presentation with plots illustrating the original data along with labeled peaks. Additionally, the script exports peak values to CSV files.
Prerequisites
		Python: Make sure you have Python installed on your system.
		Required Python packages: Install necessary packages using the following command:bash

Copy code
pip install pandas numpy matplotlib scipy pillow python-pptx
		


Usage
		Run the Script:
	•	Execute the script by running the following command in your terminal or command prompt:bash

Copy code
python script_name.py
	•	


	•	The script will prompt you to select a folder containing CSV files for analysis.
		Folder Selection:
	•	A file dialog will appear for you to choose the folder containing CSV files. Select the desired folder and click "OK."
		Analysis Process:
	•	The script will analyze each CSV file in the selected folder.
	•	For each file, it subtracts the minimum value from the gray values, plots the original data, identifies peaks, and exports the plot as an image.
	•	If peaks are found, they are labeled on the plot, and peak values are exported to a CSV file.
		PowerPoint Presentation:
	•	After analyzing all CSV files, the script generates a PowerPoint presentation.
	•	Each slide in the presentation includes up to four plots with their corresponding titles.
	•	The PowerPoint file is saved as output_presentation.pptx in the selected folder.
		Output Files:
	•	Original plot images and CSV files containing peak values are saved in the selected folder.
	•	Plot images are named as plot_filename.png.
	•	CSV files are named as peaks_data_filename.csv.
		Additional Information:
	•	The script utilizes the pandas, numpy, matplotlib, scipy, pillow, and python-pptx libraries.
	•	Ensure that your CSV files have columns named 'Distance_(cm)' and 'Gray_Value'.
Notes
	•	If no peaks are found in a CSV file, a message will be printed in the console.
	•	The original plots are saved as images to facilitate review and verification.
