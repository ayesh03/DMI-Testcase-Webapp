LPOCIP DMI Testcase Simulator Web Application
Overview

The LPOCIP DMI Testcase Simulator Web Application is a browser-based testing and automation platform designed to simulate, validate, and record communication between the LPOCIP Simulator and the DMI module.
It enables test case execution, real-time serial data exchange, screenshot capture, and Excel report generation — all through a simple web interface.

Live Demo

App Link: https://dmi-testcase-webapp.onrender.com

(Render free plan – initial startup may take 20–40 seconds before the application becomes fully operational.)

Key Features
Simulator Page

COM Port and Baud Rate Configuration

Local Serial Port Controls using Web Serial API

Connect / Disconnect controls for live testing

Real-time logs display and CRC Validation

Test Case Execution with Retry and Timeout logic

Info Data Page

Displays detailed LPOCIP module information such as RFID, GPS, Radio, Cab Info, GSM, and CRC data.

Monitors continuous link status and periodic heartbeat packets.

Start Test Page

Load and select test cases directly from the uploaded Excel sheet

Run Single or Multiple test cases automatically

Progress dialog with live percentage updates

Real-time PASS / FAIL / SKIP indicators

Automatic retry on timeout or ACK failure

Screenshot & Reporting

Captures and reconstructs screenshots transmitted from LPOCIP during test execution

Automatically saves images linked with Test Case IDs

Generates consolidated Excel report with:

Test Case ID

Result Status

Execution Time

Remarks

Embedded Screenshot Image

Notifications

Real-time alerts for connection, report generation, or error states.

Pop-up notifications with auto-dismiss timer.

System Architecture
Layered Design

Presentation Layer:
Web-based user interface developed using HTML, CSS, and JavaScript.

Business Logic Layer:

Test Case Management

CRC Calculation

Retry and Timeout Handling

Communication Layer:

Serial Communication via Web Serial API

WebSocket interface for remote testing

Data Management Layer:

Screenshot Extraction and Auto-Save

Excel Report Generation (with images) via Node.js + ExcelJS

Tech Stack
Category	Technology Used
Frontend	HTML, CSS, JavaScript
Backend	Node.js, Express.js
Excel Integration	ExcelJS
Communication	Web Serial API / WebSocket
Hosting	Render (Backend), GitHub Pages (Frontend)

Installation & Local Setup
Clone the Repository
git clone https://github.com/ayesh03/DMI-Testcase-Webapp.git
cd DMI-Testcase-Webapp

Install Dependencies
npm install

Run the Server
node server.js


Your backend will start at:
👉 http://localhost:5000

Open Frontend

Open index.html in your browser (preferably Chrome or Edge).

Report Generation Flow

Run all selected test cases.

Screenshots are received from the LPOCIP and stored temporarily.

Click “Generate Report” → frontend sends screenshot data to backend.

Backend merges data and screenshots into an Excel file using ExcelJS.

File auto-downloads as Updated_DMI-Testcases-Excelsheet.xlsx.


Notes
Ensure Chrome permissions allow Web Serial access before connecting to the COM port.
The backend server must be running for Excel report generation.
On Render free plan, backend startup may take 20–40 seconds.
