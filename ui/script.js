/**
 * Script for Exam Duty Allocation UI
 * Integrates with the existing allocation logic.
 */

// Sample designation ratio: Professor : Associate : Assistant : Research = 1 : 2 : 2 : 2
const designationRatio = {
    "Professor": 1,
    "Associate": 2,
    "Assistant": 2,
    "Research": 2
};

// Function to parse input JSON safely
function parseJSON(input) {
    try {
        return JSON.parse(input);
    } catch (e) {
        alert("Invalid JSON input");
        return null;
    }
}

// Allocation logic adapted from examDutyAllocator.js with designation ratio consideration
function allocateDutiesWithDesignation(rooms, teachers, examType) {
    const allocations = {};

    // Group teachers by designation
    const teachersByDesignation = {};
    Object.keys(designationRatio).forEach(designation => {
        teachersByDesignation[designation] = teachers.filter(t => t.designation === designation);
    });

    // Flatten teachers list in ratio order
    const orderedTeachers = [];
    const maxCount = Math.max(...Object.values(designationRatio));
    for (let i = 0; i < maxCount; i++) {
        for (const designation of Object.keys(designationRatio)) {
            if (i < designationRatio[designation]) {
                const teacherList = teachersByDesignation[designation];
                if (teacherList && teacherList.length > 0) {
                    // Push one teacher at a time to maintain ratio order
                    if (teacherList.length > i) {
                        orderedTeachers.push(teacherList[i]);
                    }
                }
            }
        }
    }

    // Filter lady and male teachers separately from ordered list
    const ladyTeachers = orderedTeachers.filter(t => t.gender.toLowerCase() === "female");
    const maleTeachers = orderedTeachers.filter(t => t.gender.toLowerCase() === "male");

    let ladyIndex = 0;
    let maleIndex = 0;

    rooms.forEach(room => {
        allocations[room.name] = [];

        // Use existing allocation logic for mid-sem and end-sem with required teachers count
        let requiredTeachers = 1;
        if (examType === "mid-sem") {
            if (room.name === "LT 402") {
                requiredTeachers = room.students > 100 ? 4 : 3;
            } else {
                requiredTeachers = room.students > 20 ? 2 : 1;
            }
        } else if (examType === "end-sem") {
            const agricultureRooms = ["Agriculture 101", "Agriculture 102", "Agriculture 103"];
            if (room.name === "LT 201") {
                requiredTeachers = 2;
            } else if (agricultureRooms.includes(room.name)) {
                requiredTeachers = 2;
            } else if (room.name === "CR 205" || room.name === "CR 604") {
                requiredTeachers = 3;
            } else if (room.name === "New Classroom") {
                requiredTeachers = room.students > 140 ? 4 : 2;
            } else {
                requiredTeachers = room.students > 80 ? 3 : 2;
            }
        }

        // Allocate teachers ensuring at least one lady teacher
        let allocatedLady = false;
        for (let i = 0; i < requiredTeachers; i++) {
            if (!allocatedLady && ladyIndex < ladyTeachers.length) {
                allocations[room.name].push(ladyTeachers[ladyIndex]);
                ladyIndex++;
                allocatedLady = true;
            } else if (maleIndex < maleTeachers.length) {
                allocations[room.name].push(maleTeachers[maleIndex]);
                maleIndex++;
            } else if (ladyIndex < ladyTeachers.length) {
                allocations[room.name].push(ladyTeachers[ladyIndex]);
                ladyIndex++;
            } else {
                allocations[room.name].push({ error: "Not enough teachers available" });
                break;
            }
        }
    });

    return allocations;
}

// Event listener for allocation button with file uploads and exam dates
document.getElementById("allocateBtn").addEventListener("click", () => {
    const teachersFileInput = document.getElementById("teachersFileInput");
    const roomsFileInput = document.getElementById("roomsFileInput");
    const examType = document.getElementById("examTypeSelect").value;
    const examDatesInput = document.getElementById("examDatesInput").value;

    if (teachersFileInput.files.length === 0) {
        alert("Please upload a teachers JSON file.");
        return;
    }
    if (roomsFileInput.files.length === 0) {
        alert("Please upload a rooms JSON file.");
        return;
    }
    if (!examDatesInput.trim()) {
        alert("Please enter examination dates.");
        return;
    }

    const examDates = examDatesInput.split(",").map(d => d.trim()).filter(d => d);

    const teachersReader = new FileReader();
    const roomsReader = new FileReader();

    let teachers = null;
    let rooms = null;

    teachersReader.onload = function(e) {
        try {
            teachers = JSON.parse(e.target.result);
            if (rooms !== null) {
                proceedAllocation();
            }
        } catch {
            alert("Invalid teachers JSON file.");
        }
    };

    roomsReader.onload = function(e) {
        try {
            rooms = JSON.parse(e.target.result);
            if (teachers !== null) {
                proceedAllocation();
            }
        } catch {
            alert("Invalid rooms JSON file.");
        }
    };

    teachersReader.readAsText(teachersFileInput.files[0]);
    roomsReader.readAsText(roomsFileInput.files[0]);

    function proceedAllocation() {
        // Generate shifts based on examDates and examType
        const shifts = examDates.length * (examType === "mid-sem" ? 4 : examType === "end-sem" ? 2 : 1);

        // Create a schedule object with shifts labeled S1, S2, ... and dates
        const dutySchedule = {};
        let shiftCounter = 1;
        for (const date of examDates) {
            const shiftsPerDay = examType === "mid-sem" ? 4 : examType === "end-sem" ? 2 : 1;
            for (let i = 1; i <= shiftsPerDay; i++) {
                dutySchedule[`S${shiftCounter}`] = [];
                shiftCounter++;
            }
        }

        // Call existing allocation function with generated schedule
        const allocations = allocateDutiesFromSchedule(rooms, teachers, dutySchedule);

        // Format output for display
        let output = "";
        for (const shift in allocations) {
            output += `Shift ${shift}:\n`;
            for (const roomId in allocations[shift]) {
                output += `  Room ${roomId}:\n`;
                allocations[shift][roomId].forEach(teacher => {
                    output += `    Teacher: ${teacher.name} (${teacher.gender}, ${teacher.designation})\n`;
                });
            }
            output += "\n";
        }

        document.getElementById("resultsOutput").textContent = output;

        // Generate duty chart CSV for each exam date
        examDates.forEach((date, index) => {
            const examDate = date.trim();
            const dutyChartCSV = generateDutyChartCSV(allocations, teachers, examDate);
            const filename = `Duty_Chart_${examDate.replace(/\//g, '-')}.csv`;
            downloadDutyChart(dutyChartCSV, filename);
        });

        // Create a downloadable file of the detailed results
        const blob = new Blob([output], { type: "text/plain" });
        const url = URL.createObjectURL(blob);
        let downloadLink = document.getElementById("downloadLink");
        if (!downloadLink) {
            downloadLink = document.createElement("a");
            downloadLink.id = "downloadLink";
            downloadLink.textContent = "Download Detailed Allocation Results";
            downloadLink.style.display = "block";
            downloadLink.style.marginTop = "10px";
            document.body.appendChild(downloadLink);
        }
        downloadLink.href = url;
        downloadLink.download = "detailed_allocation_results.txt";
    }
});

// New event listener for Excel allocation button
document.getElementById("allocateExcelBtn").addEventListener("click", () => {
    const fileInput = document.getElementById("excelFileInput");
    const teachersInput = document.getElementById("teachersInput").value;
    const roomsInput = document.getElementById("roomsInput").value;

    const teachers = parseJSON(teachersInput);
    const rooms = parseJSON(roomsInput);

    if (!teachers || !rooms) {
        alert("Please provide valid teachers and rooms JSON.");
        return;
    }

    if (fileInput.files.length === 0) {
        alert("Please upload a duty chart Excel file.");
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        const dutySchedule = parseExcelData(jsonData);

        // Call the new allocation function with the parsed schedule
        const allocations = allocateDutiesFromSchedule(rooms, teachers, dutySchedule);

        // Format output for display
        let output = "";
        for (const shift in allocations) {
            output += `Shift ${shift}:\n`;
            for (const roomId in allocations[shift]) {
                output += `  Room ${roomId}:\n`;
                allocations[shift][roomId].forEach(teacher => {
                    output += `    Teacher: ${teacher.name} (${teacher.gender}, ${teacher.designation})\n`;
                });
            }
            output += "\n";
        }

        document.getElementById("resultsOutput").textContent = output;

        // Generate duty chart CSV
        const examDate = "03-04-2025"; // Default date for Excel-based allocation
        const dutyChartCSV = generateDutyChartCSV(allocations, teachers, examDate);
        const filename = `Duty_Chart_${examDate.replace(/\//g, '-')}.csv`;
        downloadDutyChart(dutyChartCSV, filename);

        // Create a downloadable file of the detailed results
        const blob = new Blob([output], { type: "text/plain" });
        const url = URL.createObjectURL(blob);
        let downloadLink = document.getElementById("downloadLink");
        if (!downloadLink) {
            downloadLink = document.createElement("a");
            downloadLink.id = "downloadLink";
            downloadLink.textContent = "Download Detailed Allocation Results";
            downloadLink.style.display = "block";
            downloadLink.style.marginTop = "10px";
            document.body.appendChild(downloadLink);
        }
        downloadLink.href = url;
        downloadLink.download = "detailed_allocation_results.txt";
    };

    reader.readAsArrayBuffer(file);
});

// Function to generate duty chart CSV output
function generateDutyChartCSV(allocations, teachers, examDate) {
    // Create header information
    const header = [
        '"GRAPHIC ERA HILL UNIVERSITY, DEHRADUN"',
        '"REVISED MID SEMESTER REGULAR EXAMINATION, ' + examDate + '"',
        'Faculty Duty Chart',
        ' Invigilators must report 30 mins before the commencement of the Examination in the Exam cell and be present in the allotted room 15 mins prior to the commencement of the exam.',
        'Duty allotment and distribution of answer sheets in the Main building will be done from the Exam control room on 2nd Floor Pharmacy Lab behind Student Lift',
        '"In case of duty swapping, bring to the notice of the C.O.E a day prior duly approved by the C.O.E and signed application by both the faculties."',
        '',
        '"Reporting Time for Shift S1 is 9:00 AM Shift S2 is 11:00 AM, S3 is 1:00 PM and S4 is 3:00 PM. Invigilators to kindly frisk the students  if need be to check for any mobile phones and smart gadgets or any other material corroborating to Unfair means. Kindly permit students only with hard copy of admit cards in the Examination room. Teachers are required to be present in the designated room for the entire duration of the examination. Internal swapping or exchange of duties to be completely avoided . Shift Timings . S1 :- 9:30 AM - 11:00 AM     S2 :- 11:30 AM - 1:00 PM       S3 :- 1:30 PM - 3:00 PM          S4 :- 3:30 PM- 5:00 PM"',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        'S.No.,Name,Department,Designation,' + examDate + ',,,,,,,,,,,,,,,,,,,,,,,',
        ',,,,S1,S2,S3,S4,,,,,,,,,,,,,,,,,,,,'
    ];

    // Create teacher duty assignments
    const teacherDuties = [];
    let serialNumber = 1;

    teachers.forEach(teacher => {
        const dutyRow = {
            serialNo: serialNumber,
            name: teacher.name,
            department: teacher.department || 'N/A',
            designation: teacher.designation || 'N/A',
            s1: '0',
            s2: '0',
            s3: '0',
            s4: '0'
        };

        // Check which shifts this teacher is assigned to
        Object.keys(allocations).forEach(shift => {
            Object.keys(allocations[shift]).forEach(roomId => {
                const roomTeachers = allocations[shift][roomId];
                const isAssigned = roomTeachers.some(t => t.id === teacher.id);
                if (isAssigned) {
                    dutyRow[shift.toLowerCase()] = '1';
                }
            });
        });

        teacherDuties.push(dutyRow);
        serialNumber++;
    });

    // Convert to CSV format
    const csvRows = header.map(row => row);
    
    teacherDuties.forEach(duty => {
        const row = [
            duty.serialNo,
            duty.name,
            duty.department,
            duty.designation,
            duty.s1,
            duty.s2,
            duty.s3,
            duty.s4
        ];
        csvRows.push(row.join(','));
    });

    return csvRows.join('\n');
}

// Function to create and download duty chart file
function downloadDutyChart(csvContent, filename) {
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Parse Excel data to duty schedule
function parseExcelData(data) {
    const dutySchedule = {};
    let headerIndex = -1;

    // Find header row with shifts (S1, S2, S3, S4)
    for (let i = 0; i < data.length; i++) {
        if (data[i].includes("S1") && data[i].includes("S4")) {
            headerIndex = i;
            break;
        }
    }

    if (headerIndex === -1) {
        alert("Invalid Excel format: Shift headers not found.");
        return {};
    }

    const headers = data[headerIndex];
    const shiftIndices = {};
    headers.forEach((h, idx) => {
        if (["S1", "S2", "S3", "S4"].includes(h)) {
            shiftIndices[h] = idx;
        }
    });

    // Parse rows after header
    for (let i = headerIndex + 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length < headers.length) continue;
        const teacherName = row[1];
        if (!teacherName) continue;

        Object.entries(shiftIndices).forEach(([shift, idx]) => {
            if (row[idx] === 1) {
                if (!dutySchedule[shift]) {
                    dutySchedule[shift] = [];
                }
                dutySchedule[shift].push(teacherName);
            }
        });
    }

    return dutySchedule;
}
