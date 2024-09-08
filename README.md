# Python Automation Project 1
This is an arranged application from my final project for the Python course at college.

## Application Description
This application generates a project progress report in Excel based on task list data. The task list data is initially provided in CSV format, which is downloaded from a work management system. Each record in the CSV represents details and implementation time for a specific task. The application extracts the required information from the CSV file and integrates it into a report template.

![スクリーンショット 2024-09-08 153903](https://github.com/user-attachments/assets/876487fc-2830-49f5-adc6-65a32dcd67ad)


## Purpose of the Automation
The data required for the report is available in a work management system, but the system does not support a custom report format used by the company. As a result, workers are required to create the report manually from downloaded data in the system. Additionally, the system tracks data by individual work units, whereas the report is organized by tasks, leading to potential confusion and a higher risk of errors with more records. This application aims to automate the process, reducing errors and saving time.

### Manual Process
The manual process involves the following steps:

1. Download work time data from the work management system.
2. Open the file in Excel.
3. Copy and paste the relevant data into the report template.

### Data Collection and Integration
To generate the report:

1. Collect data grouped by project number.
2. Insert the project information into the report.
3. Gather data by task number within each project.
4. Add the task information to the report for each task.

### Task list
![スクリーンショット 2024-09-08 152438](https://github.com/user-attachments/assets/fde08b9c-5d08-43f6-a080-3850999122f4)

### Project Progress Report Template
![スクリーンショット 2024-09-08 152539](https://github.com/user-attachments/assets/e0c440e6-466f-42d5-93df-dee2fe2ead92)

![スクリーンショット 2024-09-08 154417](https://github.com/user-attachments/assets/acaf27b7-702b-43ae-9039-98592a43634e)
### Generated report by the application
![スクリーンショット 2024-09-08 152616](https://github.com/user-attachments/assets/c1288e34-5749-48d3-a18f-c7c15f860027)

